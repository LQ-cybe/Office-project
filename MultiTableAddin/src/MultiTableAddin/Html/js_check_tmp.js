
(function(){
  "use strict";
  // 构建信息由 C# 在加载时注入（window.__mtBuild），无需桥接即可显示，用于核对版本。
  var BUILD = window.__mtBuild || { time:"未知", version:"未知", dll:"未知" };
  (function(){ var c=document.getElementById('content'); if(c) c.innerHTML='<div class="loading">正在加载…<br><span style="font-size:12px;color:var(--text-3)">构建 '+escapeHtml(BUILD.time)+' · '+escapeHtml(BUILD.version)+'</span></div>'; })();

  // ───────── 桥接：C# <-> JS ─────────
  var pending = {};
  var msgId = 1;
  var BRIDGE_TIMEOUT = 20000;
  var listenerAdded = false;
  function webviewReady(){ return !!(window.chrome && window.chrome.webview); }
  function onBridgeMessage(e){
    var raw = e.data, data;
    if(raw && typeof raw === 'object'){ data = raw; }
    else { try{ data = JSON.parse(raw); }catch(err){ return; } }
    if(data && data.id && pending[data.id]){
      var p = pending[data.id]; delete pending[data.id];
      if(p.timer) clearTimeout(p.timer);
      if(data.ok) p.resolve(data.data); else p.reject(new Error(data.error||'bridge error'));
    }
  }
  function ensureListener(){
    if(webviewReady() && !listenerAdded){
      window.chrome.webview.addEventListener('message', onBridgeMessage);
      listenerAdded = true; return true;
    }
    return listenerAdded;
  }
  window.api = {
    call: function(method, params){
      return new Promise(function(resolve, reject){
        if(!webviewReady()){ reject(new Error('WebView2 桥接未就绪')); return; }
        ensureListener();
        var id = 'm' + (msgId++);
        var entry = { resolve: resolve, reject: reject };
        entry.timer = setTimeout(function(){
          if(pending[id]){ delete pending[id]; reject(new Error('与 C# 通信超时（'+(BRIDGE_TIMEOUT/1000)+'s）。请查看日志，并确认：① Excel 已打开含“超级表/表格”的工作簿；② 已完全关闭旧 Excel/WPS 并重装插件；③ 已安装 WebView2 Runtime')); }
        }, BRIDGE_TIMEOUT);
        pending[id] = entry;
        try { window.chrome.webview.postMessage(JSON.stringify({ id:id, method:method, params: params||{} })); }
        catch(err){ clearTimeout(entry.timer); delete pending[id]; reject(err); }
      });
    }
  };

  // ───────── 全局状态 ─────────
  var state = {
    appInfo:null, tables:[], sheet:null, table:null,
    fields:[], rows:[], config:null, views:[],
    activeViewId:null, visibleFields:[], colWidths:{},
    sortField:null, sortOrder:null, filter:'', selectedRow:-1
  };
  var FIELD_ICON = { Table:'▦', Form:'▤', Kanban:'▥', Gallery:'▧', Calendar:'▦', Gantt:'▨', Dashboard:'▣', Chart:'▤' };
  var VIEW_LABEL = { Table:'表格', Form:'表单', Kanban:'看板', Gallery:'画册', Calendar:'日历', Gantt:'甘特', Dashboard:'仪表盘', Chart:'图表' };

  function $(id){ return document.getElementById(id); }
  function escapeHtml(s){ return String(s==null?'':s).replace(/[&<>"']/g, function(c){ return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]; }); }
  function toast(msg){ var t=$('toast'); t.textContent=msg; t.classList.remove('hidden'); clearTimeout(t._t); t._t=setTimeout(function(){ t.classList.add('hidden'); }, 1800); }

  // ───────── 初始化 ─────────
  function waitReady(cb){
    var tries=0;
    (function loop(){ if(webviewReady()){ cb(); return; } if(tries++>150){ showFatal('WebView2 桥接长时间未就绪，请确认已安装 WebView2 Runtime，并完全关闭 Excel/WPS 后重装插件'); return; } setTimeout(loop,100); })();
  }
  function showFatal(msg){ $('content').innerHTML='<div class="placeholder"><div class="big">⚠️</div><div style="max-width:560px;text-align:center">'+escapeHtml(msg)+'</div><button class="btn primary" id="retryBtn" style="margin-top:12px">重试</button></div>'; var b=$('retryBtn'); if(b) b.addEventListener('click', function(){ $('content').innerHTML='<div class="loading">正在加载…</div>'; start(); }); }
  function start(){
    api.call('getAppInfo').then(function(info){ state.appInfo=info; }).catch(function(){});
    api.call('listTables').then(function(tables){
      state.tables = tables||[];
      renderTableSelect();
      if(state.tables.length){ openTable(state.tables[0].sheet, state.tables[0].table); }
      else { $('content').innerHTML='<div class="placeholder"><div class="big">📭</div><div>当前工作簿没有超级表（ListObject）</div><div>请在 Excel/WPS 中把数据区域转为「表格 / 超级表」</div></div>'; }
    }).catch(function(err){ showFatal('加载失败：'+err.message); });
    $('tableSelect').addEventListener('change', function(){ var v=this.value; var t=state.tables.find(function(x){return x.sheet+'/'+x.table===v;}); if(t) openTable(t.sheet, t.table); });
    $('refreshBtn').addEventListener('click', function(){ reloadTables(); });
    $('aboutBtn').addEventListener('click', showAbout);
    $('pinBtn').addEventListener('click', function(){
      state.pinned = !state.pinned;
      $('pinBtn').classList.toggle('active', state.pinned);
      api.call('setTopMost',{value:state.pinned}).catch(function(){});
    });
  }

  function reloadTables(){
    api.call('listTables').then(function(tables){
      state.tables = tables||[];
      renderTableSelect();
      if(state.sheet && state.table){ openTable(state.sheet, state.table); }
      else if(state.tables.length){ openTable(state.tables[0].sheet, state.tables[0].table); }
      else { $('content').innerHTML='<div class="placeholder"><div class="big">📭</div><div>当前工作簿没有超级表（ListObject）</div></div>'; }
      toast('已刷新数据源');
    }).catch(function(err){ toast('刷新失败：'+err.message); });
  }

  function renderTableSelect(){
    var sel=$('tableSelect'); sel.innerHTML='';
    state.tables.forEach(function(t){
      var o=document.createElement('option'); o.value=t.sheet+'/'+t.table;
      // 数据源格式：工作表名称（超级表名称）
      o.textContent=t.sheet+'（'+t.table+'）';
      sel.appendChild(o);
    });
  }

  function openTable(sheet, table, afterRender){
    api.call('openTable',{sheet:sheet, table:table}).then(function(d){
      state.sheet=d.sheet; state.table=d.table; state.fields=applyFieldOverrides(d.fields||[], d.fieldOverrides||[]); state.rows=d.rows;
      state.config=d.config; state.views=d.config.views||[];
      state.sortField=null; state.sortOrder=null; state.filter=''; state.selectedRow=-1;
      state.colWidths={};
      state.fields.forEach(function(f){ state.colWidths[f.name]=defaultWidth(f.type); });
      ensureViewDefaults();
      var prevId = state.activeViewId;
      var tv = state.views.find(function(v){return v.viewId===prevId && v.viewType!=='Chart';}) || state.views.find(function(v){return v.viewType==='Table';}) || state.views.find(function(v){return v.viewType!=='Chart';}) || null;
      state.activeViewId = tv? tv.viewId : null;
      // 每个视图独立维护自己的可见字段；切换数据源时以当前视图为准
      state.visibleFields = (tv && tv.visibleFields && tv.visibleFields.length) ? tv.visibleFields.slice() : (d.visibleFields||[]).slice();
      if($('tableSelect').value !== sheet+'/'+table) $('tableSelect').value = sheet+'/'+table;
      renderSidebar(); renderViewbar(); renderContent();
      if(typeof afterRender==='function') afterRender();
    }).catch(function(err){ toast('打开失败：'+err.message); });
  }

  function defaultWidth(type){
    switch(type){
      case 'Number': case 'Integer': case 'Currency': case 'Percentage': return 110;
      case 'Checkbox': return 70;
      case 'Date': case 'DateTime': return 140;
      case 'Select': case 'Quarter': case 'Phone': return 130;
      case 'Image': return 150;
      default: return 170;
    }
  }
  function applyFieldOverrides(fields, overrides){
    var map={}; (overrides||[]).forEach(function(o){ map[o.name]=o; });
    return fields.map(function(f){ var o=map[f.name]; if(!o) return f; var r={}; for(var k in f) r[k]=f[k]; r.type=o.type||f.type; r.options=(o.options&&o.options.length)?o.options.slice():f.options; r.format=o.format||''; r.required=!!o.required; r.minValue=o.minValue; r.maxValue=o.maxValue; r.minLength=o.minLength; r.maxLength=o.maxLength; r.regex=o.regexPattern||''; r.errorMessage=o.errorMessage||''; return r; });
  }
  function getFieldOverride(name){
    var f=state.fields.find(function(x){return x.name===name;}); return f?{name:f.name, type:f.type, options:(f.options||[]).slice(), format:f.format||'', required:!!f.required, minValue:f.minValue, maxValue:f.maxValue, minLength:f.minLength, maxLength:f.maxLength, regexPattern:f.regex||'', errorMessage:f.errorMessage||''}:null;
  }

  // 为每个视图补全默认专属配置（仅当缺省时）
  function ensureViewDefaults(){
    var changed=false;
    state.views.forEach(function(v){
      if(v.viewType==='Kanban'){
        if(!v.groupBy){ var g=firstField(['Select','Quarter','Text']); v.groupBy=g?g.name:''; changed=true; }
        if(!v.cardMeta){ var tf=firstField(['Text']); v.cardMeta={title:tf?tf.name:'',image:'',description:[]}; changed=true; }
        else if(!v.cardMeta.title){ var tf2=firstField(['Text']); if(tf2) v.cardMeta.title=tf2.name; changed=true; }
      } else if(v.viewType==='Gallery'){
        if(!v.cardMeta){ var im=firstField(['Image']); var tf3=firstField(['Text']); v.cardMeta={title:tf3?tf3.name:'',image:im?im.name:'',description:[]}; changed=true; }
      } else if(v.viewType==='Calendar'){
        if(!v.calendarConfig){ var df=firstField(['Date','DateTime']); var df2=state.fields.filter(isDate)[1]; var tf4=firstField(['Text']); v.calendarConfig={dateField:df?df.name:'',endDateField:df2?df2.name:'',titleField:tf4?tf4.name:''}; changed=true; }
      } else if(v.viewType==='Gantt'){
        if(!v.ganttConfig){ var ds=state.fields.filter(isDate); v.ganttConfig={startField:ds[0]?ds[0].name:'',endField:ds[1]?ds[1].name:ds[0]?ds[0].name:'',labelField:(firstField(['Text'])||{}).name||'',groupField:'',progressField:'',timeDimension:'Month',displayFields:[],colorMode:'field',colorField:'',customColor:'#3370FF',workdaysOnly:false}; changed=true; }
      } else if(v.viewType==='Dashboard'){
        if(!v.dashboardConfig){ var num2=firstField(['Number','Integer','Currency','Percentage']); var dim2=firstField(['Select','Quarter','Text']);
          v.dashboardConfig={columns:2, statCards: num2?[{id:'k1',title:num2.name+' 合计',field:num2.name,aggregation:'Sum',format:'auto',color:''}]:[], charts: dim2&&num2?[{id:'c1',title:dim2.name+' 分布',type:'Column',dimensionField:dim2.name,metricField:num2.name,aggregation:'Sum',timeField:'',timeGroup:'None',seriesField:'',topN:12,gaugeTarget:100,columnSpan:1,height:260}]:[]}; changed=true; }
      }
    });
    // 保证甘特视图始终可用：老配置/旧版本创建时若未生成甘特视图，而数据又存在日期字段，则自动补建
    var hasGantt = state.views.some(function(v){ return v.viewType==='Gantt'; });
    if(!hasGantt){
      var dsAll = state.fields.filter(isDate);
      if(dsAll.length){
        state.views.push({
          viewId: 'gantt-'+Date.now().toString(36),
          viewType: 'Gantt',
          viewName: '甘特视图',
          visibleFields: state.fields.map(function(f){return f.name;}),
          ganttConfig: { startField: dsAll[0].name, endField: dsAll.length>1?dsAll[1].name:dsAll[0].name, labelField: (firstField(['Text'])||{}).name||dsAll[0].name, groupField:'', progressField:'', timeDimension:'Month', displayFields:[], colorMode:'field', colorField:'', customColor:'#3370FF', workdaysOnly:false }
        });
        changed=true;
      }
    }
    if(changed) saveConfig();
  }
  function firstField(types){ return state.fields.find(function(f){ return types.indexOf(f.type)>=0; }); }
  function isDate(f){ return f.type==='Date'||f.type==='DateTime'; }

  // ───────── 侧栏 ─────────
  function renderSidebar(){
    var sb=$('sidebar'); sb.innerHTML='';
    var title=document.createElement('div'); title.className='group-title'; title.textContent='视图'; sb.appendChild(title);
    // 过滤掉已废弃的统计图表视图
    state.views.forEach(function(v){
      if(v.viewType==='Chart') return;
      var it=document.createElement('div'); it.className='view-item'+(v.viewId===state.activeViewId?' active':'');
      it.innerHTML='<span class="ic">'+(FIELD_ICON[v.viewType]||'▦')+'</span><span>'+escapeHtml(v.viewName||VIEW_LABEL[v.viewType]||v.viewType)+'</span>';
      it.addEventListener('click', function(){
        state.activeViewId=v.viewId;
        // 切换视图时恢复该视图自己的可见字段
        state.visibleFields = (v.visibleFields && v.visibleFields.length) ? v.visibleFields.slice() : state.fields.map(function(f){return f.name;});
        state.filter=''; state.sortField=null; state.sortOrder=null;
        renderSidebar(); renderViewbar(); renderContent();
      });
      sb.appendChild(it);
    });
  }

  // ───────── 视图工具条 ─────────
  function renderViewbar(){
    var vb=$('viewbar'); vb.innerHTML='';
    var tools=document.createElement('div'); tools.className='tools';
    var type=activeViewType();
    if(type==='Form'){
      var add2=document.createElement('button'); add2.className='btn primary sm'; add2.textContent='＋ 新增'; add2.addEventListener('click', addRow); tools.appendChild(add2);
      var setf2=document.createElement('button'); setf2.className='btn sm'; setf2.textContent='字段设置'; setf2.addEventListener('click', showFieldSettings); tools.appendChild(setf2);
    } else if(type==='Dashboard'){
      var add3=document.createElement('button'); add3.className='btn sm'; add3.textContent='＋ 图表'; add3.addEventListener('click', function(){ addDashboardItem('chart'); }); tools.appendChild(add3);
      var add4=document.createElement('button'); add4.className='btn sm'; add4.textContent='＋ 指标'; add4.addEventListener('click', function(){ addDashboardItem('kpi'); }); tools.appendChild(add4);
      var vDash=activeView(); var allShow=(vDash.dashboardConfig&&vDash.dashboardConfig.charts||[]).every(function(c){return c.showLabels!==false;});
      var lblToggle=document.createElement('button'); lblToggle.className='btn sm'; lblToggle.textContent=allShow?'隐藏标签':'显示标签';
      lblToggle.addEventListener('click', function(){ var v=activeView(); var charts=v.dashboardConfig?v.dashboardConfig.charts:[]; var newShow=!charts.every(function(c){return c.showLabels!==false;}); charts.forEach(function(c){c.showLabels=newShow;}); saveConfig(); renderContent(); });
      tools.appendChild(lblToggle);
      renderSaveSettings(tools);
    } else {
      renderCommonTools(tools, type);
    }
    vb.appendChild(tools);
  }
  function renderSaveSettings(tools){
    var save=document.createElement('button'); save.className='btn primary sm'; save.textContent='保存设置';
    save.addEventListener('click', function(){ saveConfig(); toast('视图设置已保存'); });
    tools.appendChild(save);
  }
  function renderCommonTools(tools, type){
    var v=activeView();
    var add=document.createElement('button'); add.className='btn primary sm'; add.textContent='＋ 新增'; add.addEventListener('click', addRow); tools.appendChild(add);
    if(type!=='Chart'){
      var fld=document.createElement('button'); fld.className='btn sm'; fld.textContent='字段开关'; fld.addEventListener('click', function(){ toggleFieldPanel(this); }); tools.appendChild(fld);
    }
    if(type!=='Calendar' && type!=='Chart'){
      var setf=document.createElement('button'); setf.className='btn sm'; setf.textContent='字段设置'; setf.addEventListener('click', showFieldSettings); tools.appendChild(setf);
    }
    if(type==='Calendar'){
      var calSet=document.createElement('button'); calSet.className='btn sm'; calSet.textContent='日历设置';
      calSet.addEventListener('click', showCalendarSettings); tools.appendChild(calSet);
    }
    if(type==='Gantt'){
      var ganttSet=document.createElement('button'); ganttSet.className='btn sm'; ganttSet.textContent='甘特图配置';
      ganttSet.addEventListener('click', showGanttSettings); tools.appendChild(ganttSet);
    }
    if(type==='Kanban'){
      // 卡片标题字段直接放在工具栏
      var titleWrap=document.createElement('div'); titleWrap.style.cssText='display:flex;align-items:center;gap:4px';
      var titleSel=document.createElement('select'); titleSel.className='btn sm'; titleSel.style.cssText='height:24px;padding:0 4px';
      titleSel.appendChild(new Option('卡片标题',''));
      state.fields.forEach(function(f){ titleSel.appendChild(new Option(f.name, f.name)); });
      var cm=v&&v.cardMeta? v.cardMeta:{}; titleSel.value=cm.title||'';
      titleSel.addEventListener('change', function(){ v.cardMeta=v.cardMeta||{}; v.cardMeta.title=this.value; v.cardMeta.description=[]; saveConfig(); renderContent(); });
      titleWrap.appendChild(titleSel); tools.appendChild(titleWrap);
    }
    var sf=document.createElement('input'); sf.className='search'; sf.placeholder='查找…'; sf.value=state.filter;
    sf.addEventListener('input', function(){ state.filter=this.value; renderContent(); }); tools.appendChild(sf);
    if(type!=='Calendar' && type!=='Gantt' && type!=='Chart'){
      // 排序
      var sortWrap=document.createElement('div'); sortWrap.style.cssText='display:flex;align-items:center;gap:4px';
      var sortSel=document.createElement('select'); sortSel.className='btn sm'; sortSel.style.cssText='height:24px;padding:0 4px';
      sortSel.appendChild(new Option('不排序',''));
      state.fields.forEach(function(f){ sortSel.appendChild(new Option(f.name, f.name)); });
      sortSel.value=state.sortField||'';
      sortSel.addEventListener('change', function(){ state.sortField=this.value||null; if(state.sortField) state.sortOrder=state.sortOrder||'asc'; renderContent(); });
      var sortDir=document.createElement('button'); sortDir.className='btn sm'; sortDir.textContent=state.sortOrder==='desc'?'降序':'升序'; sortDir.title='切换升降序';
      sortDir.addEventListener('click', function(){ state.sortOrder=(state.sortOrder==='asc'?'desc':'asc'); renderViewbar(); renderContent(); });
      sortWrap.appendChild(sortSel); sortWrap.appendChild(sortDir); tools.appendChild(sortWrap);
    }
    if(type!=='Calendar' && type!=='Chart'){
      // 分组
      var groupWrap=document.createElement('div'); groupWrap.style.cssText='display:flex;align-items:center;gap:4px';
      var groupSel=document.createElement('select'); groupSel.className='btn sm'; groupSel.style.cssText='height:24px;padding:0 4px';
      groupSel.appendChild(new Option('不分组',''));
      state.fields.forEach(function(f){ if(f.type==='Select'||f.type==='Quarter'||f.type==='Text'||f.type==='Checkbox') groupSel.appendChild(new Option(f.name, f.name)); });
      var gv=v?viewGroupField(v):'';
      groupSel.value=gv;
      groupSel.addEventListener('change', function(){ setViewGroupField(v, this.value); renderContent(); });
      groupWrap.appendChild(groupSel); tools.appendChild(groupWrap);
    }
    renderSaveSettings(tools);
  }
  function viewGroupField(v){ if(!v) return ''; if(v.viewType==='Kanban') return v.groupBy||''; if(v.ganttConfig) return v.ganttConfig.groupField||''; return v.groupBy||''; }
  function setViewGroupField(v, val){ if(!v) return; if(v.viewType==='Kanban') v.groupBy=val; else if(v.ganttConfig) v.ganttConfig.groupField=val; else v.groupBy=val; saveConfig(); }
  function activeView(){ return state.views.find(function(x){return x.viewId===state.activeViewId;}); }
  function activeViewType(){ var v=activeView(); return v?v.viewType:null; }

  // ───────── 内容渲染分发 ─────────
  function renderContent(){
    var type=activeViewType();
    if(!type){ $('content').innerHTML='<div class="placeholder"><div class="big">📭</div><div>请选择一个视图</div></div>'; return; }
    if(type!=='Form') closeSettingsPanel();
    // 表格视图保持右侧详情面板常驻；其它视图按原逻辑关闭
    if(type!=='Table') closeDetailPanel();
    if(type==='Table') renderTable();
    else if(type==='Form'){ renderForm(); if(!$('settingsPanel').classList.contains('open')) showFieldSettings(); }
    else if(type==='Kanban') renderKanban();
    else if(type==='Gallery') renderGallery();
    else if(type==='Calendar') renderCalendar();
    else if(type==='Gantt') renderGantt();
    else if(type==='Dashboard') renderDashboard();
    else if(type==='Chart') renderChart(activeView());
    else { $('content').innerHTML='<div class="placeholder"><div class="big">🚧</div><div>视图类型 '+escapeHtml(type)+' 暂未实现</div></div>'; }
    // 表格视图渲染后，若详情面板已打开则同步当前选中行
    if(type==='Table' && state.selectedRow>0 && $('detailPanel').classList.contains('open')){ showDetailPanel(state.selectedRow); }
    // 若右侧详情已打开，高亮左侧对应数据
    syncLeftSelectionHighlight();
  }

  function getVisibleRows(){
    var rows = state.rows.slice();
    var f = state.filter.trim().toLowerCase();
    if(f){ rows = rows.filter(function(r){ return state.visibleFields.some(function(fn){ return String(displayText(fn,r)).toLowerCase().indexOf(f)>=0; }); }); }
    if(state.sortField && state.sortOrder){
      var fn=state.sortField, ord=state.sortOrder==='asc'?1:-1;
      rows.sort(function(a,b){
        var va=rawNum(a.values[fn]), vb=rawNum(b.values[fn]);
        if(va==null && vb==null) return 0;
        if(va==null) return 1; if(vb==null) return -1;
        if(typeof va==='number' && typeof vb==='number') return (va-vb)*ord;
        var sa=String(va), sb=String(vb); return sa<sb?-1*ord:(sa>sb?1*ord:0);
      });
    }
    return rows;
  }
  function rawNum(v){ if(typeof v==='number') return v; if(v==null) return null; var n=Number(v); return isNaN(n)?v:n; }

  // ───────── 表格视图 ─────────
  function renderTable(){
    var wrap=document.createElement('div'); wrap.className='table-wrap';
    var table=document.createElement('table'); table.className='grid';
    var colgroup=document.createElement('colgroup');
    colgroup.appendChild(col('col-rownum',46));
    state.visibleFields.forEach(function(fn){ colgroup.appendChild(col('', state.colWidths[fn]||160)); });
    colgroup.appendChild(col('col-actions',40));
    table.appendChild(colgroup);
    var thead=document.createElement('thead'); var htr=document.createElement('tr');
    htr.appendChild(th('序号', 'col-rownum', null, false));
    state.visibleFields.forEach(function(fn){ var f=state.fields.find(function(x){return x.name===fn;}); htr.appendChild(th(fieldLabel(f), '', f, true)); });
    htr.appendChild(th('', 'col-actions', null, false));
    thead.appendChild(htr); table.appendChild(thead);
    var tbody=document.createElement('tbody');
    var rows=getVisibleRows();
    if(rows.length===0){
      var tr=document.createElement('tr'); var td=document.createElement('td');
      td.colSpan=state.visibleFields.length+2; td.style.textAlign='center'; td.style.color='var(--text-3)';
      td.textContent = state.rows.length===0 ? '暂无数据，点击右上角「＋ 新增」添加记录' : '没有匹配筛选条件的记录';
      tr.appendChild(td); tbody.appendChild(tr);
    }
    var v=activeView();
    var groupBy=v?viewGroupField(v):'';
    var gf=groupBy?state.fields.find(function(f){return f.name===groupBy;}):null;
    if(groupBy && gf){
      var groups={};
      rows.forEach(function(r){ var k=displayText(groupBy,r)||'（空）'; (groups[k]=groups[k]||[]).push(r); });
      var keys=Object.keys(groups);
      if(gf.type==='Select'||gf.type==='Quarter'){ var opts=(gf.options||[]).slice(); opts.push('（空）'); keys=opts.filter(function(k){return groups[k];}); Object.keys(groups).forEach(function(k){ if(keys.indexOf(k)<0) keys.push(k); }); }
      keys.forEach(function(k){
        var gh=document.createElement('tr'); gh.className='group-header';
        var ghd=document.createElement('td'); ghd.className='group-title-cell'; ghd.colSpan=state.visibleFields.length+2;
        ghd.innerHTML='<span class="gdot"></span><span class="gname">'+escapeHtml(k)+'</span><span class="gcount">'+groups[k].length+'</span>';
        gh.appendChild(ghd); tbody.appendChild(gh);
        groups[k].forEach(function(r){ appendTableRow(tbody, r); });
      });
    } else {
      rows.forEach(function(r){ appendTableRow(tbody, r); });
    }
    table.appendChild(tbody); wrap.appendChild(table);
    $('content').innerHTML=''; $('content').appendChild(wrap);
    autoFitColumns();
  }
  function appendTableRow(tbody, r){
    var tr=document.createElement('tr'); tr.dataset.row=JSON.stringify({rowIndex:r.rowIndex});
    if(r.rowIndex===state.selectedRow) tr.className='selected';
    var td0=document.createElement('td'); td0.className='col-rownum'; td0.textContent=r.rowIndex; tr.appendChild(td0);
    state.visibleFields.forEach(function(fn){
      var f=state.fields.find(function(x){return x.name===fn});
      var td=document.createElement('td'); td.className=isNum(f)?'num':'';
      td.dataset.field=fn; td.dataset.rowIndex=r.rowIndex;
      td.innerHTML=cellHtml(f,r);
      td.addEventListener('click', function(e){ e.stopPropagation(); selectRow(r.rowIndex, tr); editCell(td, f, r); });
      tr.appendChild(td);
    });
    var tdA=document.createElement('td'); tdA.className='col-actions';
    var del=document.createElement('span'); del.className='row-del'; del.textContent='🗑'; del.title='删除该行';
    del.addEventListener('click', function(e){ e.stopPropagation(); deleteRow(r.rowIndex); });
    tdA.appendChild(del); tr.appendChild(tdA);
    tr.addEventListener('click', function(){ selectRow(r.rowIndex, tr); });
    tr.addEventListener('contextmenu', function(e){ showRowContextMenu(e, r.rowIndex); });
    tbody.appendChild(tr);
  }
  function autoFitColumns(){
    var table=$('content').querySelector('table.grid'); if(!table) return;
    var cols=table.querySelectorAll('colgroup col');
    var headerCells=table.querySelectorAll('thead th');
    var bodyRows=table.querySelectorAll('tbody tr');
    var sample=Math.min(bodyRows.length, 200);
    var maxCol=state.visibleFields.length+1; // 含序号列和操作列
    for(var i=0;i<=maxCol;i++){
      var w=measureText(headerCells[i]?headerCells[i].textContent:'', 12)+28;
      for(var s=0;s<sample;s++){
        var td=bodyRows[s].children[i]; if(td && td.textContent){ var tw=measureText(td.textContent, 12)+22; if(tw>w) w=tw; }
      }
      // 序号/操作列更紧凑
      if(i===0 || i===maxCol){ w=Math.max(34, Math.min(w, 70)); }
      else {
        var fn=state.visibleFields[i-1]; var f=state.fields.find(function(x){return x.name===fn;});
        var minW=70;
        if(f){ if(f.type==='Date'||f.type==='DateTime') minW=170; else if(f.type==='Select'||f.type==='Quarter') minW=150; else if(f.type==='Currency'||f.type==='Percentage') minW=110; }
        // 日期/下拉列额外加宽，避免编辑控件撑出或换行
        if(f && (f.type==='Date'||f.type==='DateTime'||f.type==='Select'||f.type==='Quarter')) w+=34;
        w=Math.max(minW, Math.min(w, 360));
      }
      if(cols[i]) cols[i].style.width=w+'px';
      if(i>0 && i<maxCol) state.colWidths[state.visibleFields[i-1]]=w;
    }
  }
  var _measureCtx=null;
  function measureText(txt, fontSize){ if(!_measureCtx){ var c=document.createElement('canvas'); _measureCtx=c.getContext('2d'); } _measureCtx.font=(fontSize||12)+'px -apple-system,"PingFang SC","Microsoft YaHei",sans-serif'; return _measureCtx.measureText(txt||'').width; }
  function col(cls,w){ var el=document.createElement('col'); if(cls) el.className=cls; if(w) el.style.width=w+'px'; return el; }
  function th(text, cls, field, sortable){
    var el=document.createElement('th'); if(cls) el.className=cls;
    var inner=document.createElement('div'); inner.className='th-inner';
    var span=document.createElement('span'); span.textContent=text; inner.appendChild(span);
    if(sortable && field){
      var s=document.createElement('span'); s.className='sort'+(state.sortField===field.name?(state.sortOrder||''):'');
      s.textContent=(state.sortField===field.name)?(state.sortOrder==='asc'?'▲':'▼'):'⇅';
      s.title='点击排序'; s.addEventListener('click', function(e){ e.stopPropagation(); toggleSort(field.name); });
      inner.appendChild(s);
      var rs=document.createElement('span'); rs.className='resizer'; rs.addEventListener('mousedown', function(e){ startResize(e, field.name); });
      inner.appendChild(rs);
    }
    el.appendChild(inner); return el;
  }
  function fieldLabel(f){ if(!f) return ''; return f.name; }
  function typeShort(t){ return ({Text:'文本',LongText:'长文',Number:'数字',Integer:'整数',Date:'日期',DateTime:'时间',Select:'单选',Quarter:'季度',Currency:'金额',Percentage:'百分比',Email:'邮箱',Phone:'手机',Url:'网址',Checkbox:'勾选',Image:'图片'})[t]||t; }
  function isNum(f){ return f && (f.type==='Number'||f.type==='Integer'||f.type==='Currency'||f.type==='Percentage'); }
  function displayText(fieldName, row){
    var f=state.fields.find(function(x){return x.name===fieldName;});
    var raw=row.values?row.values[fieldName]:undefined;
    if(f && f.type==='Checkbox') return raw?'✓':'';
    if(row.displayTexts && row.displayTexts[fieldName]!=null && row.displayTexts[fieldName]!=='') return formatValue(f, row.displayTexts[fieldName]);
    return formatValue(f, raw);
  }
  function formatValue(f, v){
    if(v==null || v==='') return '';
    if(!f) return String(v);
    if(f.type==='Checkbox') return v?'✓':'';
    var fmt=f.format?String(f.format).trim():'';
    if((f.type==='Currency'||fmt==='money') && !isNaN(Number(v))) return '¥'+Number(v).toLocaleString('zh-CN',{minimumFractionDigits:2,maximumFractionDigits:2});
    if((f.type==='Percentage'||fmt==='percent') && !isNaN(Number(v))) return (Number(v)*100).toFixed(2)+'%';
    if(fmt==='int' && !isNaN(Number(v))) return Number(v).toLocaleString('zh-CN',{maximumFractionDigits:0});
    if(fmt==='datetime' && (f.type==='Date'||f.type==='DateTime')){ var d=toDate(v); return d?fmtDateTime(d):String(v); }
    if(f.type==='DateTime' && !fmt){ var d2=toDate(v); return d2?fmtDateTime(d2):String(v); }
    if(f.type==='Date' && !fmt){ var d3=toDate(v); return d3?fmtDate(d3):String(v); }
    if(fmt==='phone'){ var s=String(v).replace(/\D/g,''); if(s.length===11) return s.replace(/(\d{3})(\d{4})(\d{4})/,'$1 $2 $3'); }
    if(fmt==='idcard'){ var ss=String(v); if(ss.length===18) return ss.replace(/(\d{6})(\d{4})(\d{4})(\d{4})/,'$1 $2 $3 $4'); }
    return String(v);
  }
  function fmtDateTime(d){ return d.getFullYear()+'-'+pad(d.getMonth()+1)+'-'+pad(d.getDate())+' '+pad(d.getHours())+':'+pad(d.getMinutes()); }
  function cellHtml(f, r){
    var t=f.type; var raw=r.values[f.name];
    if(t==='Checkbox') return '<span class="cell-check">'+(raw?'☑':'☐')+'</span>';
    var disp=displayText(f.name, r);
    if(t==='Url' && raw) return '<a href="'+escapeHtml(raw)+'" target="_blank" style="color:var(--accent)">'+escapeHtml(disp)+'</a>';
    return escapeHtml(disp);
  }
  function toggleSort(field){ if(state.sortField!==field){ state.sortField=field; state.sortOrder='asc'; } else if(state.sortOrder==='asc'){ state.sortOrder='desc'; } else { state.sortField=null; state.sortOrder=null; } renderContent(); }
  function startResize(e, field){
    e.preventDefault(); e.stopPropagation();
    var startX=e.clientX, startW=state.colWidths[field]||160;
    function move(ev){ var w=Math.max(60, startW+(ev.clientX-startX)); state.colWidths[field]=w; var idx=state.visibleFields.indexOf(field); var colEls=$('content').querySelectorAll('colgroup col'); if(colEls[idx+1]) colEls[idx+1].style.width=w+'px'; }
    function up(){ document.removeEventListener('mousemove',move); document.removeEventListener('mouseup',up); }
    document.addEventListener('mousemove',move); document.addEventListener('mouseup',up);
  }

  // 通用就地编辑（表格用）
  function editCell(td, f, r){
    if(td.querySelector('.cell-edit')) return;
    var raw=r.values[f.name];
    var input=buildFieldInput(f, raw);
    td.classList.add('editing'); td.innerHTML=''; td.appendChild(input); input.focus(); if(input.select) try{ input.select(); }catch(_){}
    var committed=false;
    function commit(){
      if(committed) return; committed=true;
      var val=readFieldInput(f, input);
      api.call('updateCell',{rowIndex:r.rowIndex, field:f.name, value:val}).then(function(){
        r.values[f.name]=val;
        td.classList.remove('editing');
        td.innerHTML=cellHtml(f, r);
        toast('已保存');
      }).catch(function(err){
        toast('保存失败：'+err.message);
        td.classList.remove('editing');
        td.innerHTML=cellHtml(f, r);
      });
    }
    input.addEventListener('keydown', function(ev){
      if(ev.key==='Enter' && input.tagName!=='TEXTAREA'){ ev.preventDefault(); commit(); }
      else if(ev.key==='Enter' && ev.ctrlKey && input.tagName==='TEXTAREA'){ ev.preventDefault(); commit(); }
      else if(ev.key==='Escape'){ ev.preventDefault(); td.classList.remove('editing'); td.innerHTML=cellHtml(f, r); }
    });
    input.addEventListener('blur', function(){ commit(); });
    if(f.type==='Select'||f.type==='Quarter'||f.type==='Date'||f.type==='DateTime') input.addEventListener('change', function(){ commit(); });
  }
  function buildFieldInput(f, raw){
    var t=f.type, input;
    if(t==='Checkbox'){ input=document.createElement('input'); input.type='checkbox'; input.className='cell-edit'; input.checked=!!raw; }
    else if(t==='Select'||t==='Quarter'){
      input=document.createElement('select'); input.className='cell-edit';
      var opts=f.options||[]; if(t==='Quarter'&&opts.length===0) opts=['第一季度','第二季度','第三季度','第四季度'];
      opts.forEach(function(o){ var op=document.createElement('option'); op.value=o; op.textContent=o; if(o===raw) op.selected=true; input.appendChild(op); });
      var blank=document.createElement('option'); blank.value=''; blank.textContent='（空）'; if(raw==null||raw==='') blank.selected=true; input.appendChild(blank);
    } else if(t==='DateTime'){
      input=document.createElement('input'); input.type='datetime-local'; input.className='cell-edit date-edit';
      if(raw==null || raw===''){ var n=new Date(); input.value=fmtDateTimeLocal(n, n.getHours(), n.getMinutes()).slice(0,16); }
      else { var d=toDate(raw); input.value=d?fmtDateTimeLocal(d, d.getHours(), d.getMinutes()).slice(0,16):(typeof raw==='string'?raw.slice(0,16):''); }
      enhanceDateInput(input);
    } else if(t==='Date'){
      input=document.createElement('input'); input.type='date'; input.className='cell-edit date-edit';
      if(raw==null || raw==='') input.value=fmtDate(new Date());
      else input.value=(typeof raw==='string')?raw.slice(0,10):(raw&&raw.toISOString?raw.toISOString().slice(0,10):'');
      enhanceDateInput(input);
    } else if(t==='LongText'){
      input=document.createElement('textarea'); input.className='cell-edit'; input.value=raw==null?'':String(raw);
    } else if(t==='Percentage'){
      input=document.createElement('input'); input.type='number'; input.className='cell-edit'; input.step='0.01';
      input.value=(raw==null||raw==='')?'':String(Number(raw)*100);
    } else if(isNum(f)){ input=document.createElement('input'); input.type='number'; input.className='cell-edit'; input.step=(t==='Integer')?'1':'0.01'; if(raw!=null) input.value=raw; }
    else { input=document.createElement('input'); input.type='text'; input.className='cell-edit'; input.value=raw==null?'':String(raw); }
    return input;
  }
  function enhanceDateInput(input){
    function isToday(v){ if(!v) return false; var d=new Date(v), n=new Date(); return d.getFullYear()===n.getFullYear() && d.getMonth()===n.getMonth() && d.getDate()===n.getDate(); }
    function updateStyle(){ if(isToday(input.value)){ input.style.background='var(--accent)'; input.style.color='#fff'; input.style.borderColor='var(--accent)'; } else { input.style.background=''; input.style.color=''; input.style.borderColor=''; } }
    updateStyle(); input.addEventListener('input', updateStyle); input.addEventListener('change', updateStyle);
    input.addEventListener('wheel', function(e){ e.preventDefault(); if(!input.value) return; var d=new Date(input.value); if(isNaN(d)) return; var delta=e.deltaY>0?1:-1; if(e.ctrlKey){ d.setFullYear(d.getFullYear()+delta); } else if(e.shiftKey){ d.setMonth(d.getMonth()+delta*3); } else { d.setDate(d.getDate()+delta); } input.value=fmtDate(d); updateStyle(); }, {passive:false});
  }
  function readFieldInput(f, input){
    var t=f.type;
    if(t==='Checkbox') return input.checked;
    if(t==='Select'||t==='Quarter') return input.value;
    if(t==='Percentage') return input.value===''?null:(Number(input.value)/100);
    if(isNum(f)) return input.value===''?null:Number(input.value);
    if(t==='Date'||t==='DateTime') return input.value?input.value:null;
    return input.value;
  }
  function autoHeightTextarea(ta){
    ta.classList.add('auto-height');
    function lineHeight(){ var st=getComputedStyle(ta); return parseInt(st.lineHeight)||18; }
    function resize(){ ta.style.height='auto'; ta.style.height=(ta.scrollHeight+lineHeight())+'px'; }
    ta.addEventListener('input', resize);
    setTimeout(resize, 0);
  }
  function validateField(f, val){
    if(f && f.required && (val==null || val==='')) return f.name+' 为必填项';
    if(isNum(f) && val!=null && val!==''){ var n=Number(val); if(isNaN(n)) return f.name+' 必须是数字'; if(f.minValue!=null && n<f.minValue) return f.name+' 不能小于 '+f.minValue; if(f.maxValue!=null && n>f.maxValue) return f.name+' 不能大于 '+f.maxValue; }
    if(typeof val==='string'){ if(f.minLength!=null && val.length<f.minLength) return f.name+' 至少 '+f.minLength+' 个字符'; if(f.maxLength!=null && val.length>f.maxLength) return f.name+' 最多 '+f.maxLength+' 个字符'; if(f.regex && !(new RegExp(f.regex).test(val))) return f.errorMessage||(f.name+' 格式不正确'); }
    return '';
  }

  // ───────── 表单视图 ─────────
  function renderForm(){
    var wrap=document.createElement('div'); wrap.className='form-wrap';
    var filtered=getVisibleRows();
    var idx=state.formIndex||0; if(idx>=filtered.length) idx=filtered.length-1; if(idx<0) idx=0;
    var r=filtered[idx];
    // 导航
    var nav=document.createElement('div'); nav.className='form-nav';
    nav.innerHTML='<span style="color:var(--text-3)">第 '+(filtered.length?idx+1:0)+' / '+filtered.length+' 条</span>';
    var prev=document.createElement('button'); prev.className='btn sm'; prev.textContent='‹ 上一条'; prev.addEventListener('click', function(){ state.formIndex=idx-1; renderForm(); }); nav.appendChild(prev);
    var next=document.createElement('button'); next.className='btn sm'; next.textContent='下一条 ›'; next.addEventListener('click', function(){ state.formIndex=idx+1; renderForm(); }); nav.appendChild(next);
    var del=document.createElement('button'); del.className='btn sm'; del.textContent='🗑 删除'; del.addEventListener('click', function(){ if(r) deleteRow(r.rowIndex); }); nav.appendChild(del);
    wrap.appendChild(nav);
    if(!r){ wrap.innerHTML+='<div class="placeholder"><div class="big">📝</div><div>暂无记录</div></div>'; $('content').innerHTML=''; $('content').appendChild(wrap); return; }
    var card=document.createElement('div'); card.className='form-card';
    var inputs=[];
    state.visibleFields.forEach(function(fn){
      var f=state.fields.find(function(x){return x.name===fn});
      var row=document.createElement('div'); row.className='form-row';
      var lab=document.createElement('div'); lab.className='flabel'; lab.textContent=(f?f.name:fn)+(f&&f.required?' *':''); lab.title=f?f.name:''; row.appendChild(lab);
      var val=document.createElement('div'); val.className='fval';
      var input=buildFieldInput(f, r.values[fn]);
      input.dataset.field=fn; val.appendChild(input); row.appendChild(val); card.appendChild(row);
      if(input.tagName==='TEXTAREA') autoHeightTextarea(input);
      inputs.push({field:f, input:input});
    });
    wrap.appendChild(card);
    var actRow=document.createElement('div'); actRow.className='form-row';
    var actLab=document.createElement('div'); actLab.className='flabel'; actLab.textContent=''; actRow.appendChild(actLab);
    var actVal=document.createElement('div'); actVal.className='fval';
    var actions=document.createElement('div'); actions.className='form-actions';
    var saveBtn=document.createElement('button'); saveBtn.className='btn primary'; saveBtn.textContent='保存';
    saveBtn.addEventListener('click', function(){
      var todo=[]; var valid=true;
      inputs.forEach(function(it){ var val=readFieldInput(it.field, it.input); var err=validateField(it.field, val); if(err){ valid=false; toast(err); it.input.focus(); } else if(!deepEqual(val, r.values[it.field.name])) todo.push({field:it.field.name, value:val}); });
      if(!valid) return;
      if(todo.length===0){ toast('没有修改'); return; }
      var pending=todo.length;
      todo.forEach(function(u){ api.call('updateCell',{rowIndex:r.rowIndex, field:u.field, value:u.value}).then(function(){ r.values[u.field]=u.value; pending--; if(pending===0){ toast('已保存 '+todo.length+' 个字段'); renderContent(); } }).catch(function(err){ toast('保存失败：'+err.message); }); });
    });
    var cancelBtn=document.createElement('button'); cancelBtn.className='btn'; cancelBtn.textContent='取消';
    cancelBtn.addEventListener('click', function(){ renderForm(); });
    var resetBtn=document.createElement('button'); resetBtn.className='btn sm'; resetBtn.textContent='重置';
    resetBtn.addEventListener('click', function(){ inputs.forEach(function(it){ it.input=replaceInput(it.input, buildFieldInput(it.field, r.values[it.field.name])); }); });
    actions.appendChild(resetBtn); actions.appendChild(cancelBtn); actions.appendChild(saveBtn);
    actVal.appendChild(actions); actRow.appendChild(actVal);
    wrap.appendChild(actRow);
    $('content').innerHTML=''; $('content').appendChild(wrap);
    if(state._focusForm){ state._focusForm=false; var fi=$('content').querySelector('.form-row .cell-edit'); if(fi) setTimeout(function(){ fi.focus(); if(fi.select) try{fi.select();}catch(_){} },0); }
  }
  function replaceInput(oldEl, newEl){ var p=oldEl.parentNode; if(p){ p.replaceChild(newEl, oldEl); return newEl; } return oldEl; }
  function deepEqual(a,b){ if(a===b) return true; if(a==null || b==null) return a==b; return JSON.stringify(a)===JSON.stringify(b); }

  // 浮动字段编辑器（表单/看板/画册/日历/甘特复用）
  function openFieldEditor(row, field, anchor){
    closeEditor();
    var input=buildFieldInput(field, row.values[field.name]);
    var pop=document.createElement('div'); pop.className='editor-pop';
    var rect=anchor.getBoundingClientRect();
    pop.style.left=Math.min(rect.left, window.innerWidth-220)+'px';
    pop.style.top=(rect.bottom+4)+'px';
    pop.appendChild(input); document.body.appendChild(pop); input.focus(); if(input.select) try{input.select();}catch(_){}
    function commit(){ var val=readFieldInput(field, input); api.call('updateCell',{rowIndex:row.rowIndex, field:field.name, value:val}).then(function(){ row.values[field.name]=val; closeEditor(); renderContent(); toast('已保存'); }).catch(function(err){ toast('保存失败：'+err.message); closeEditor(); renderContent(); }); }
    function cancel(){ closeEditor(); }
    input.addEventListener('keydown', function(ev){ if(ev.key==='Enter'){ ev.preventDefault(); commit(); } else if(ev.key==='Escape'){ ev.preventDefault(); cancel(); } });
    input.addEventListener('blur', function(){ commit(); });
    state._editorPop=pop;
  }
  function closeEditor(){ if(state._editorPop){ state._editorPop.remove(); state._editorPop=null; } }

  // 记录详情模态（看板/画册/日历/甘特编辑整条）
  function openRecordModal(rowIndex){
    var r=state.rows.find(function(x){return x.rowIndex===rowIndex;}); if(!r) return;
    var ov=document.createElement('div'); ov.className='overlay';
    var modal=document.createElement('div'); modal.className='modal';
    modal.innerHTML='<div class="modal-head"><span>编辑记录 #'+r.rowIndex+'</span><button class="btn sm" id="xClose">✕</button></div>';
    var body=document.createElement('div'); body.className='modal-body';
    // 模态框编辑记录也显示所有字段
    state.fields.forEach(function(f){
      var fr=document.createElement('div'); fr.className='form-field';
      var lab=document.createElement('label'); lab.textContent=f.name+(f.required?' *':''); fr.appendChild(lab);
      var ctrl=document.createElement('div'); ctrl.className='ctrl';
      var input=buildFieldInput(f, r.values[f.name]);
      input.dataset.field=f.name; ctrl.appendChild(input); fr.appendChild(ctrl); body.appendChild(fr);
      if(input.tagName==='TEXTAREA') autoHeightTextarea(input);
    });
    modal.appendChild(body);
    var foot=document.createElement('div'); foot.className='modal-foot';
    var save=document.createElement('button'); save.className='btn primary'; save.textContent='保存';
    save.addEventListener('click', function(){
      var ok=true;
      body.querySelectorAll('.ctrl input,.ctrl select').forEach(function(inp){
        if(!ok) return; var f=state.fields.find(function(x){return x.name===inp.dataset.field;}); if(!f) return;
        var val=readFieldInput(f, inp);
        // 同步保存（用 Promise 串行以免并发过大）
        api.call('updateCell',{rowIndex:r.rowIndex, field:inp.dataset.field, value:val}).then(function(){ r.values[inp.dataset.field]=val; }).catch(function(err){ ok=false; toast('保存失败：'+err.message); });
      });
      setTimeout(function(){ ov.remove(); renderContent(); if(ok) toast('已保存'); }, 400);
    });
    var cancel=document.createElement('button'); cancel.className='btn'; cancel.textContent='取消'; cancel.addEventListener('click', function(){ ov.remove(); });
    foot.appendChild(cancel); foot.appendChild(save); modal.appendChild(foot);
    ov.appendChild(modal); ov.addEventListener('click', function(e){ if(e.target===ov) ov.remove(); });
    document.getElementById('overlayHost').appendChild(ov);
    modal.querySelector('#xClose').addEventListener('click', function(){ ov.remove(); });
  }
  function showDetailPanel(rowIndex){
    closeSettingsPanel();
    var r=state.rows.find(function(x){return x.rowIndex===rowIndex;}); if(!r) return;
    var dp=$('detailPanel'); dp.classList.add('open'); dp.innerHTML='';
    var head=document.createElement('div'); head.className='dp-head';
    head.innerHTML='<span>详情</span>';
    var x=document.createElement('button'); x.className='btn sm'; x.id='dpClose'; x.textContent='✕'; head.appendChild(x); dp.appendChild(head);
    var body=document.createElement('div'); body.className='dp-body';
    var inputs=[];
    // 右侧详情显示所有字段，不受字段开关控制
    state.fields.forEach(function(f){
      var fr=document.createElement('div'); fr.className='form-field';
      var lab=document.createElement('label'); lab.textContent=f.name+(f.required?' *':''); fr.appendChild(lab);
      var ctrl=document.createElement('div'); ctrl.className='ctrl';
      var input=buildFieldInput(f, r.values[f.name]); input.dataset.field=f.name; ctrl.appendChild(input); fr.appendChild(ctrl); body.appendChild(fr);
      if(input.tagName==='TEXTAREA') autoHeightTextarea(input);
      inputs.push({field:f, input:input});
    });
    dp.appendChild(body);
    var foot=document.createElement('div'); foot.className='dp-foot';
    var cancel=document.createElement('button'); cancel.className='btn'; cancel.textContent='取消';
    cancel.addEventListener('click', closeDetailPanel);
    var del=document.createElement('button'); del.className='btn sm'; del.textContent='🗑 删除';
    del.addEventListener('click', function(){ deleteRow(r.rowIndex); closeDetailPanel(); });
    var save=document.createElement('button'); save.className='btn primary'; save.textContent='保存';
    save.addEventListener('click', function(){
      var todo=[]; var valid=true;
      inputs.forEach(function(it){ var val=readFieldInput(it.field, it.input); var err=validateField(it.field, val); if(err){ valid=false; toast(err); it.input.focus(); } else if(!deepEqual(val, r.values[it.field.name])) todo.push({field:it.field.name, value:val}); });
      if(!valid) return;
      if(todo.length===0){ toast('没有修改'); return; }
      var pending=todo.length;
      todo.forEach(function(u){ api.call('updateCell',{rowIndex:r.rowIndex, field:u.field, value:u.value}).then(function(){ r.values[u.field]=u.value; pending--; if(pending===0){ toast('已保存 '+todo.length+' 个字段'); renderContent(); } }).catch(function(err){ toast('保存失败：'+err.message); }); });
    });
    foot.appendChild(cancel); foot.appendChild(del); foot.appendChild(save); dp.appendChild(foot);
    x.addEventListener('click', closeDetailPanel);
    syncLeftSelectionHighlight();
  }
  // 日历/甘特等视图中新建记录草稿面板（点击保存后才真正 addRow）
  function showDraftDetailPanel(prefillValues, contextDate, contextHour, contextMin){
    closeSettingsPanel();
    var dp=$('detailPanel'); dp.classList.add('open'); dp.innerHTML='';
    var head=document.createElement('div'); head.className='dp-head';
    head.innerHTML='<span>新建记录</span>';
    var x=document.createElement('button'); x.className='btn sm'; x.id='dpClose'; x.textContent='✕'; head.appendChild(x); dp.appendChild(head);
    var body=document.createElement('div'); body.className='dp-body';
    var inputs=[];
    state.fields.forEach(function(f){
      var fr=document.createElement('div'); fr.className='form-field';
      var lab=document.createElement('label'); lab.textContent=f.name+(f.required?' *':''); fr.appendChild(lab);
      var ctrl=document.createElement('div'); ctrl.className='ctrl';
      var input=buildFieldInput(f, prefillValues[f.name]!==undefined?prefillValues[f.name]:(f.type==='Date'||f.type==='DateTime'?(contextDate?fmtDate(contextDate):''):'')); input.dataset.field=f.name; ctrl.appendChild(input); fr.appendChild(ctrl); body.appendChild(fr);
      if(input.tagName==='TEXTAREA') autoHeightTextarea(input);
      inputs.push({field:f, input:input});
    });
    dp.appendChild(body);
    var foot=document.createElement('div'); foot.className='dp-foot';
    var cancel=document.createElement('button'); cancel.className='btn'; cancel.textContent='取消';
    cancel.addEventListener('click', closeDetailPanel);
    var save=document.createElement('button'); save.className='btn primary'; save.textContent='保存';
    save.addEventListener('click', function(){
      var values={}; var valid=true;
      inputs.forEach(function(it){ var val=readFieldInput(it.field, it.input); var err=validateField(it.field, val); if(err){ valid=false; toast(err); it.input.focus(); } else { values[it.field.name]=val; } });
      if(!valid) return;
      api.call('addRow',{values:values}).then(function(res){
        if(res&&res.rowIndex>0){
          openTable(state.sheet, state.table, function(){ state.selectedRow=res.rowIndex; showDetailPanel(res.rowIndex); toast('已添加记录'); });
        } else toast('添加失败');
      }).catch(function(err){ toast('添加失败：'+err.message); });
    });
    foot.appendChild(cancel); foot.appendChild(save); dp.appendChild(foot);
    x.addEventListener('click', closeDetailPanel);
  }

  function showDateDetailPanel(date){
    closeSettingsPanel();
    var dp=$('detailPanel'); dp.classList.add('open'); dp.innerHTML='';
    var head=document.createElement('div'); head.className='dp-head';
    head.innerHTML='<span>'+fmtDate(date)+' 记录</span>';
    var x=document.createElement('button'); x.className='btn sm'; x.textContent='✕'; x.addEventListener('click', closeDetailPanel); head.appendChild(x);
    dp.appendChild(head);
    var body=document.createElement('div'); body.className='dp-body';
    var cfg=(activeView().calendarConfig||{});
    var evs=state.rows.filter(function(r){ var dv=toDate(r.values[cfg.dateField]); return dv && isSameDay(dv, date); });
    if(evs.length===0){ body.innerHTML='<div style="color:var(--text-3);padding:10px 0">当日无记录</div>'; }
    evs.forEach(function(r){
      var card=document.createElement('div'); card.className='cal-event'; card.style.padding='8px 10px'; card.style.marginBottom='6px';
      card.textContent=displayText(cfg.titleField||(firstField(['Text'])||{}).name||'', r)||('记录 #'+r.rowIndex);
      card.addEventListener('click', function(){ showDetailPanel(r.rowIndex); });
      body.appendChild(card);
    });
    dp.appendChild(body);
  }
  function closeDetailPanel(){ var dp=$('detailPanel'); if(dp){ dp.classList.remove('open'); dp.innerHTML=''; } }
  // 打开右侧详情时，高亮左侧对应的数据控件
  function syncLeftSelectionHighlight(){
    if(state.selectedRow<=0) return;
    var dp=$('detailPanel'); if(!dp || !dp.classList.contains('open')) return;
    // 表格
    var trs=$('content').querySelectorAll('tbody tr'); trs.forEach(function(tr){ try{ var d=JSON.parse(tr.dataset.row||'{}'); tr.classList.toggle('selected', d.rowIndex===state.selectedRow); }catch(_){} });
    // 看板/画册
    $('content').querySelectorAll('.kanban-card,.gallery-card').forEach(function(el){ el.classList.toggle('selected', parseInt(el.dataset.row,10)===state.selectedRow); });
    // 日历事件
    $('content').querySelectorAll('.cal-day-event').forEach(function(el){ el.classList.toggle('selected', parseInt(el.dataset.row,10)===state.selectedRow); });
    // 甘特
    $('content').querySelectorAll('.gantt-row').forEach(function(el){ el.classList.toggle('selected', parseInt(el.dataset.row,10)===state.selectedRow); });
  }

  // ───────── 看板视图 ─────────
  function renderKanban(){
    var v=activeView(); var cfg=v||{};
    var groupBy=cfg.groupBy; var groupField=state.fields.find(function(f){return f.name===groupBy;});
    if(!groupBy){ $('content').innerHTML='<div class="placeholder"><div class="big">▥</div><div>请先在「不分组」下拉框中选择分组字段</div></div>'; return; }
    var titleField=cfg.cardMeta&&cfg.cardMeta.title?cfg.cardMeta.title:firstField(['Text']);
    var titleName=(typeof titleField==='object'?titleField.name:titleField)||'';
    // 分组取值
    var groups={}; state.rows.forEach(function(r){ var k=displayText(groupBy,r)||'（空）'; (groups[k]=groups[k]||[]).push(r); });
    var keys=Object.keys(groups);
    // 若该分组字段是 Select/Quarter，按选项顺序优先
    if(groupField && (groupField.type==='Select'||groupField.type==='Quarter')){
      var opts=(groupField.options||[]).slice(); opts.push('（空）'); keys=opts.filter(function(k){return groups[k];}); Object.keys(groups).forEach(function(k){ if(keys.indexOf(k)<0) keys.push(k); });
    }
    var wrap=document.createElement('div'); wrap.className='kanban-wrap';
    var cols=document.createElement('div'); cols.className='kanban-cols';
    var groupColors=['#3370FF','#14C9C9','#FF7D00','#F76965','#9FDB60','#CA62E8','#FFC53D','#6E4AE0'];
    keys.forEach(function(k, idx){
      var col=document.createElement('div'); col.className='kanban-col'; col.dataset.group=k;
      var color=groupColors[idx % groupColors.length];
      col.style.setProperty('--group-color', color);
      col.style.backgroundColor=color+'10';
      col.style.borderColor=color;
      var head=document.createElement('div'); head.className='kanban-col-head'; head.innerHTML='<span>'+escapeHtml(k)+'</span><span class="cnt">'+groups[k].length+'</span>'; col.appendChild(head);
      var body=document.createElement('div'); body.className='kanban-col-body';
      groups[k].forEach(function(r){
        var card=document.createElement('div'); card.className='kanban-card'+(r.rowIndex===state.selectedRow?' selected':''); card.dataset.row=r.rowIndex;
        var title=titleName?displayText(titleName,r):('记录 #'+r.rowIndex);
        var html='<div class="ktitle">'+escapeHtml(title)+'</div>';
        // 字段开关：显示除标题外的可见字段
        var chips=[];
        state.visibleFields.forEach(function(fn){ if(fn===titleName) return; var f=state.fields.find(function(x){return x.name===fn;}); if(!f) return; var val=displayText(fn,r); if(val==='') return; chips.push('<span class="kfchip" title="'+escapeHtml(f.name)+'">'+escapeHtml(f.name+': '+val)+'</span>'); });
        if(chips.length) html+='<div class="kfields">'+chips.join('')+'</div>';
        card.innerHTML=html;
        card.addEventListener('click', function(){ selectRow(r.rowIndex); showDetailPanel(r.rowIndex); });
        enableKanbanDrag(card, r, groupBy);
        body.appendChild(card);
      });
      col.addEventListener('mouseenter', function(){ if(state._kanbanDragCard) col.classList.add('dropover'); });
      col.addEventListener('mouseleave', function(){ col.classList.remove('dropover'); });
      col.appendChild(body); cols.appendChild(col);
    });
    wrap.appendChild(cols); $('content').innerHTML=''; $('content').appendChild(wrap);
  }
  // 看板自定义拖拽（WebView2 中 HTML5 DnD 不稳定，改用鼠标事件）
  function enableKanbanDrag(card, r, groupBy){
    card.addEventListener('mousedown', function(e){
      if(e.button!==0) return;
      var startX=e.clientX, startY=e.clientY, moved=false;
      var rect=card.getBoundingClientRect();
      var ghost=card.cloneNode(true); ghost.classList.add('dragging'); ghost.style.position='fixed'; ghost.style.left=rect.left+'px'; ghost.style.top=rect.top+'px'; ghost.style.width=rect.width+'px'; ghost.style.zIndex='100'; ghost.style.pointerEvents='none'; document.body.appendChild(ghost);
      card.classList.add('dragging');
      state._kanbanDragCard=card; state._kanbanDragRow=r.rowIndex; state._kanbanDragGroupBy=groupBy;
      function move(em){
        var dx=em.clientX-startX, dy=em.clientY-startY;
        if(!moved && (Math.abs(dx)>3 || Math.abs(dy)>3)) moved=true;
        if(moved){ ghost.style.left=(rect.left+dx)+'px'; ghost.style.top=(rect.top+dy)+'px'; }
      }
      function up(em){
        document.removeEventListener('mousemove', move); document.removeEventListener('mouseup', up);
        ghost.remove(); card.classList.remove('dragging');
        var targetCol=null;
        if(moved){
          ghost.style.display='none';
          var el=document.elementFromPoint(em.clientX, em.clientY);
          if(el) targetCol=el.closest('.kanban-col');
        }
        state._kanbanDragCard=null; state._kanbanDragRow=null; state._kanbanDragGroupBy=null;
        document.querySelectorAll('.kanban-col.dropover').forEach(function(c){ c.classList.remove('dropover'); });
        if(moved && targetCol){
          var newVal=targetCol.dataset.group==='（空）'?'':targetCol.dataset.group;
          var oldVal=r.values[groupBy];
          if(String(oldVal||'')!==String(newVal||'')){
            var rr=state.rows.find(function(x){return x.rowIndex===r.rowIndex;}); if(rr) rr.values[groupBy]=newVal;
            renderContent();
            api.call('updateCell',{rowIndex:r.rowIndex, field:groupBy, value:newVal}).then(function(){ toast('已移动到「'+targetCol.dataset.group+'」'); }).catch(function(err){ toast('移动失败：'+err.message+'，正在刷新'); openTable(state.sheet, state.table); });
          }
        }
      }
      document.addEventListener('mousemove', move); document.addEventListener('mouseup', up);
    });
  }

  // ───────── 画册视图 ─────────
  function renderGallery(){
    var v=activeView(); var cfg=v||{};
    if(!cfg.cardMeta){ $('content').innerHTML='<div class="placeholder"><div class="big">▧</div><div>请先在「字段设置」中设置图片与标题字段</div></div>'; return; }
    var imgField=cfg.cardMeta.image, titleField=cfg.cardMeta.title;
    var grid=document.createElement('div'); grid.className='gallery-grid';
    getVisibleRows().forEach(function(r){
      var card=document.createElement('div'); card.className='gallery-card'+(r.rowIndex===state.selectedRow?' selected':''); card.dataset.row=r.rowIndex;
      var img=document.createElement('div'); img.className='gimg';
      var imgVal=imgField?String(r.values[imgField]||'').trim():'';
      if(imgVal){ var im=document.createElement('img'); im.src=imgVal; im.alt=''; img.innerHTML=''; img.appendChild(im); }
      else img.textContent='（无图片）';
      card.appendChild(img);
      var body=document.createElement('div'); body.className='gbody';
      var t=titleField?displayText(titleField,r):('记录 #'+r.rowIndex);
      var tb=document.createElement('div'); tb.className='gtitle'; tb.textContent=t; body.appendChild(tb);
      // 字段开关：显示除图片/标题外的可见字段
      var fields=[];
      state.visibleFields.forEach(function(fn){ if(fn===imgField || fn===titleField) return; var f=state.fields.find(function(x){return x.name===fn;}); if(!f) return; var val=displayText(fn,r); if(val==='') return; fields.push('<div class="gfld"><span class="k">'+escapeHtml(f.name)+'</span> '+escapeHtml(val)+'</div>'); });
      if(fields.length){ var gf=document.createElement('div'); gf.className='gfields'; gf.innerHTML=fields.join(''); body.appendChild(gf); }
      card.appendChild(body);
      card.addEventListener('click', function(){ selectRow(r.rowIndex); showDetailPanel(r.rowIndex); });
      grid.appendChild(card);
    });
    $('content').innerHTML=''; $('content').appendChild(grid);
  }

  // ───────── 日历视图 ─────────
  var calState={ year:0, month:0, view:'month', day:1 };
  function renderCalendar(){
    var v=activeView(); var cfg=v.calendarConfig||{};
    if(!cfg.dateField){ $('content').innerHTML='<div class="placeholder"><div class="big">▦</div><div>请先在「日历设置」中选择日期字段</div></div>'; return; }
    var now=new Date();
    if(!calState.year){ calState={year:now.getFullYear(), month:now.getMonth(), view:'month', day:now.getDate()}; }
    var y=calState.year, m=calState.month, view=calState.view||'month';
    var wrap=document.createElement('div'); wrap.className='calendar-wrap';
    wrap.style.position='relative';
    var head=document.createElement('div'); head.className='calendar-head';
    var title='';
    if(view==='month') title=y+' 年 '+(m+1)+' 月';
    else if(view==='week'){ var ws=weekStart(y,m,calState.day); title=fmtDate(ws)+' ~ '+fmtDate(new Date(ws.getTime()+6*86400000)); }
    else if(view==='day') title=y+' 年 '+(m+1)+' 月 '+calState.day+' 日';
    head.innerHTML='<button class="btn sm" id="calPrev">‹</button><span class="cal-title">'+title+'</span><button class="btn sm" id="calNext">›</button><span class="cal-spacer"></span><button class="btn sm" id="calToday">今天</button><button class="btn sm'+(view==='month'?' active':'')+'" id="calMonth" title="月视图">月</button><button class="btn sm'+(view==='week'?' active':'')+'" id="calWeek" title="周视图">周</button><button class="btn sm'+(view==='day'?' active':'')+'" id="calDay" title="日视图">日</button>';
    wrap.appendChild(head);
    // + 添加记录
    var addBtn=document.createElement('button'); addBtn.className='btn primary sm cal-add'; addBtn.textContent='＋';
    addBtn.title='添加记录（日期为 '+fmtDate(now)+'）';
    addBtn.addEventListener('click', function(){ var vals={}; vals[cfg.dateField]=fmtDate(now); api.call('addRow',{values:vals}).then(function(res){ if(res&&res.rowIndex>0){ var newIdx=res.rowIndex; openTable(state.sheet, state.table, function(){ state.selectedRow=newIdx; showDetailPanel(newIdx); toast('已添加记录，请在右侧编辑后保存'); }); } }).catch(function(err){ toast('添加失败：'+err.message); }); });
    wrap.appendChild(addBtn);
    // 事件按日期分组
    var evMap={};
    state.rows.forEach(function(r){ var dv=toDate(r.values[cfg.dateField]); if(!dv) return; var key=dv.getFullYear()+'-'+(dv.getMonth()+1)+'-'+dv.getDate(); (evMap[key]=evMap[key]||[]).push(r); });
    if(view==='month') renderMonthView(wrap, y, m, evMap, cfg, now);
    else if(view==='week') renderWeekView(wrap, y, m, calState.day, cfg, now);
    else renderDayView(wrap, y, m, calState.day, cfg, now);
    $('content').innerHTML=''; $('content').appendChild(wrap);
    var step=(view==='day'?1:(view==='week'?7:1));
    var unit=(view==='day'?'d':(view==='week'?'d':'m'));
    $('calPrev').addEventListener('click', function(){ moveCal(-step, unit); renderCalendar(); });
    $('calNext').addEventListener('click', function(){ moveCal(step, unit); renderCalendar(); });
    $('calToday').addEventListener('click', function(){ var n=new Date(); calState={year:n.getFullYear(),month:n.getMonth(),view:view,day:n.getDate()}; renderCalendar(); });
    $('calDay').addEventListener('click', function(){ calState.view='day'; renderCalendar(); });
    $('calWeek').addEventListener('click', function(){ calState.view='week'; renderCalendar(); });
    $('calMonth').addEventListener('click', function(){ calState.view='month'; renderCalendar(); });
  }
  function moveCal(n, unit){
    var d=new Date(calState.year, calState.month, calState.day||1);
    if(unit==='m'){ d.setMonth(d.getMonth()+n); }
    else { d.setDate(d.getDate()+n); }
    calState.year=d.getFullYear(); calState.month=d.getMonth(); calState.day=d.getDate();
  }
  function weekStart(y,m,day){
    var d=new Date(y,m,day||1); var dow=(d.getDay()+6)%7; return new Date(d.getTime()-dow*86400000);
  }
  function isSameDay(a,b){ return a.getFullYear()===b.getFullYear() && a.getMonth()===b.getMonth() && a.getDate()===b.getDate(); }
  function renderMonthView(wrap, y, m, evMap, cfg, now){
    var first=new Date(y, m, 1); var startDow=(first.getDay()+6)%7;
    var daysInMonth=new Date(y, m+1, 0).getDate();
    var prevDays=new Date(y, m, 0).getDate();
    var grid=document.createElement('div'); grid.className='calendar-grid';
    ['一','二','三','四','五','六','日'].forEach(function(d){ var c=document.createElement('div'); c.className='cal-dow'; c.textContent=d; grid.appendChild(c); });
    var cells=[];
    for(var i=0;i<startDow;i++) cells.push({out:true, d:prevDays-startDow+1+i, date:new Date(y,m-1,prevDays-startDow+1+i)});
    for(var d=1;d<=daysInMonth;d++) cells.push({out:false, d:d, date:new Date(y,m,d)});
    var tail=0; while(cells.length%7!==0){ tail++; cells.push({out:true, d:tail, date:new Date(y,m+1,tail)}); }
    cells.forEach(function(c){
      var cell=document.createElement('div'); cell.className='cal-cell'+(c.out?' out':'');
      var isToday=!c.out && isSameDay(c.date, now);
      cell.innerHTML='<div class="dnum">'+(isToday?'<span class="today">'+c.d+'</span>':c.d)+'</div>';
      if(!c.out){
        var key=c.date.getFullYear()+'-'+(c.date.getMonth()+1)+'-'+c.date.getDate();
        var list=evMap[key]||[];
        list.slice(0,5).forEach(function(r){
          var ev=document.createElement('div'); ev.className='cal-event'; ev.textContent=displayText(cfg.titleField||(firstField(['Text'])||{}).name||'', r)||('记录 #'+r.rowIndex);
          ev.title=cfg.titleField?displayText(cfg.titleField,r):'';
          ev.addEventListener('click', function(e){ e.stopPropagation(); showDetailPanel(r.rowIndex); });
          cell.appendChild(ev);
        });
        if(list.length>5){
          var more=document.createElement('div'); more.className='cal-event cal-more'; more.textContent='还有 '+(list.length-5)+' 条记录'; more.title='共 '+list.length+' 条';
          more.addEventListener('click', function(e){ e.stopPropagation(); showDateDetailPanel(c.date); });
          cell.appendChild(more);
        }
        cell.addEventListener('click', function(e){
          if(e.target.closest('.cal-event')) return;
          var df=state.fields.find(function(x){return x.name===cfg.dateField;});
          var vals={}; vals[cfg.dateField]=fmtDate(c.date)+(df&&df.type==='DateTime'?'T09:00:00':'');
          showDraftDetailPanel(vals, c.date, 9, 0);
        });
      }
      grid.appendChild(cell);
    });
    wrap.appendChild(grid);
  }
  function renderWeekView(wrap, y, m, day, cfg, now){
    var start=weekStart(y,m,day);
    var days=[]; for(var i=0;i<7;i++) days.push(new Date(start.getTime()+i*86400000));
    renderTimeGrid(wrap, days, cfg, now, true);
  }
  function renderDayView(wrap, y, m, day, cfg, now){
    renderTimeGrid(wrap, [new Date(y,m,day)], cfg, now, false);
  }
  function renderTimeGrid(wrap, days, cfg, now, isWeek){
    var timeline=document.createElement('div'); timeline.className='cal-timeline';
    var axis=document.createElement('div'); axis.className='cal-time-axis';
    for(var h=0;h<24;h++){ var lbl=document.createElement('div'); lbl.className='cal-time-label'; lbl.textContent=(h<10?'0':'')+h+':00'; axis.appendChild(lbl); }
    timeline.appendChild(axis);
    var cols=document.createElement('div'); cols.className='cal-week-cols';
    days.forEach(function(d, idx){
      var col=document.createElement('div'); col.className='cal-day-col'+(isSameDay(d,now)?' today':'');
      var hd=document.createElement('div'); hd.className='cal-day-col-head';
      if(isWeek){ hd.textContent=['周一','周二','周三','周四','周五','周六','周日'][idx]+' '+d.getDate()+'日'; }
      else { hd.textContent=(isSameDay(d,now)?'今天 · ':'')+fmtDate(d); }
      hd.addEventListener('click', (function(date){ return function(){ showDateDetailPanel(date); }; })(d));
      col.appendChild(hd);
      var allday=document.createElement('div'); allday.className='cal-allday';
      var body=document.createElement('div'); body.className='cal-day-col-body';
      body.dataset.date=fmtDate(d);
      for(var hh=0;hh<24;hh++){ var slot=document.createElement('div'); slot.className='cal-slot'; slot.style.top=(hh*40)+'px'; body.appendChild(slot); }
      // 点击空白时间格：打开右侧草稿详情，保存后才真正新增
      body.addEventListener('click', function(e){
        if(e.target.closest('.cal-day-event')) return;
        var rect=body.getBoundingClientRect();
        var y=e.clientY-rect.top+body.scrollTop;
        var hour=Math.floor(y/40), min=Math.round((y%40)/40*60);
        hour=Math.max(0,Math.min(23,hour)); min=Math.max(0,Math.min(59,min));
        var df=state.fields.find(function(x){return x.name===cfg.dateField;});
        var isDt=df&&df.type==='DateTime';
        var vals={}; vals[cfg.dateField]=isDt?fmtDateTimeLocal(d, hour, min):fmtDate(d);
        if(cfg.endDateField){ var ef=state.fields.find(function(x){return x.name===cfg.endDateField;}); vals[cfg.endDateField]=(ef&&ef.type==='DateTime')?fmtDateTimeLocal(d, hour+1, min):fmtDate(d); }
        showDraftDetailPanel(vals, d, hour, min);
      });
      var evs=state.rows.filter(function(r){ var dv=toDate(r.values[cfg.dateField]); return dv && isSameDay(dv,d); });
      evs.forEach(function(r){ renderCalEvent(body, allday, r, cfg, d); });
      col.appendChild(allday); col.appendChild(body); cols.appendChild(col);
    });
    // 当前时间红线：放在今天的列内，圆点位于当前日期列
    days.forEach(function(d, idx){
      if(!isSameDay(d,now)) return;
      var col=cols.children[idx]; if(!col) return;
      var body=col.querySelector('.cal-day-col-body'); if(!body) return;
      var line=document.createElement('div'); line.className='cal-current-line';
      line.style.top=(now.getHours()*40 + now.getMinutes()/60*40)+'px';
      body.appendChild(line);
    });
    timeline.appendChild(cols); wrap.appendChild(timeline);
  }
  function renderCalEvent(body, alldayContainer, r, cfg, dayDate){
    var dv=toDate(r.values[cfg.dateField]); if(!dv) return;
    var endDv=cfg.endDateField?toDate(r.values[cfg.endDateField]):null;
    var title=displayText(cfg.titleField||(firstField(['Text'])||{}).name||'', r)||('记录 #'+r.rowIndex);
    var timeStr=pad(dv.getHours())+':'+pad(dv.getMinutes())+(endDv?' - '+pad(endDv.getHours())+':'+pad(endDv.getMinutes()):'');
    var ev=document.createElement('div'); ev.className='cal-day-event'+(r.rowIndex===state.selectedRow?' selected':''); ev.dataset.row=r.rowIndex;
    ev.innerHTML='<span class="evt-time">'+timeStr+'</span><span class="evt-title">'+escapeHtml(title)+'</span>';
    ev.title=title+'\n'+timeStr;
    var isAllday=!endDv && dv.getHours()===0 && dv.getMinutes()===0;
    if(isAllday){
      ev.classList.add('allday'); ev.textContent=timeStr+' '+title;
      alldayContainer.appendChild(ev);
    } else {
      ev.style.top=(dv.getHours()*40 + dv.getMinutes()/60*40)+'px';
      var endTop=endDv?(endDv.getHours()*40 + endDv.getMinutes()/60*40):(dv.getHours()*40+40);
      var h=Math.max(20, endTop-(dv.getHours()*40 + dv.getMinutes()/60*40));
      ev.style.height=h+'px';
      makeCalEventDraggable(ev, body, r, cfg);
      body.appendChild(ev);
    }
    ev.addEventListener('click', function(e){ e.stopPropagation(); selectRow(r.rowIndex); showDetailPanel(r.rowIndex); });
  }
  function makeCalEventDraggable(ev, body, r, cfg){
    var resize=document.createElement('div'); resize.className='cal-event-resize'; ev.appendChild(resize);
    // 拖动整体移动时间
    ev.addEventListener('mousedown', function(e){
      if(e.target===resize || e.target.classList.contains('cal-event-resize')) return;
      e.preventDefault();
      var startY=e.clientY, startTop=parseFloat(ev.style.top)||0;
      var origDv=new Date(r.values[cfg.dateField]); var origEnd=cfg.endDateField?toDate(r.values[cfg.endDateField]):null;
      var duration=origEnd?(origEnd.getTime()-origDv.getTime()):3600000;
      function move(em){
        var dy=em.clientY-startY; var newTop=startTop+dy;
        newTop=Math.max(0, Math.min(920, newTop));
        ev.style.top=newTop+'px';
        var totalMin=Math.round(newTop/40*60);
        var hour=Math.floor(totalMin/60), min=totalMin%60;
        hour=Math.max(0,Math.min(23,hour)); min=Math.max(0,Math.min(59,min));
        var endTotal=Math.min(1439, totalMin+Math.round(duration/60000));
        var eh=Math.floor(endTotal/60), emin=endTotal%60;
        var timeSpan=ev.querySelector('.evt-time'); if(timeSpan){
          timeSpan.textContent=pad(hour)+':'+pad(min)+(cfg.endDateField?' - '+pad(eh)+':'+pad(emin):'');
        }
      }
      function up(em){
        document.removeEventListener('mousemove', move); document.removeEventListener('mouseup', up);
        var dy=em.clientY-startY; var newTop=startTop+dy;
        newTop=Math.max(0, Math.min(920, newTop));
        var totalMin=Math.round(newTop/40*60);
        var hour=Math.floor(totalMin/60), min=totalMin%60;
        hour=Math.max(0,Math.min(23,hour)); min=Math.max(0,Math.min(59,min));
        var base=toDate(r.values[cfg.dateField])||new Date();
        var newStart=new Date(base.getFullYear(), base.getMonth(), base.getDate(), hour, min);
        var newEnd=new Date(newStart.getTime()+duration);
        updateCalEventTime(r, cfg, newStart, newEnd);
      }
      document.addEventListener('mousemove', move); document.addEventListener('mouseup', up);
    });
    // 拖动底部调整结束时间
    resize.addEventListener('mousedown', function(e){
      e.preventDefault(); e.stopPropagation();
      var startY=e.clientY, startH=parseFloat(ev.style.height)||38;
      function move(em){ var dy=em.clientY-startY; var newH=Math.max(20, startH+dy); ev.style.height=newH+'px'; }
      function up(em){
        document.removeEventListener('mousemove', move); document.removeEventListener('mouseup', up);
        var newH=Math.max(20, startH+(em.clientY-startY));
        var startDv=toDate(r.values[cfg.dateField]); if(!startDv) return;
        var extraMin=Math.round(newH/40*60);
        var newEnd=new Date(startDv.getTime()+extraMin*60000);
        updateCalEventTime(r, cfg, startDv, newEnd);
      }
      document.addEventListener('mousemove', move); document.addEventListener('mouseup', up);
    });
  }
  function updateCalEventTime(r, cfg, newStart, newEnd){
    var sv=fmtDateTimeLocal(newStart, newStart.getHours(), newStart.getMinutes());
    var updates=[{field:cfg.dateField, value:sv}];
    if(cfg.endDateField && newEnd){ updates.push({field:cfg.endDateField, value:fmtDateTimeLocal(newEnd, newEnd.getHours(), newEnd.getMinutes())}); }
    // 先本地更新再保存
    r.values[cfg.dateField]=sv;
    if(cfg.endDateField && newEnd) r.values[cfg.endDateField]=fmtDateTimeLocal(newEnd, newEnd.getHours(), newEnd.getMinutes());
    renderContent();
    var pending=updates.length;
    updates.forEach(function(u){
      api.call('updateCell',{rowIndex:r.rowIndex, field:u.field, value:u.value}).then(function(){ pending--; if(pending===0) toast('时间已更新'); }).catch(function(err){ toast('保存失败：'+err.message); openTable(state.sheet, state.table); });
    });
  }
  function fmtDateTimeLocal(d, h, m){ return d.getFullYear()+'-'+pad(d.getMonth()+1)+'-'+pad(d.getDate())+'T'+pad(h)+':'+pad(m)+':00'; }
  function toDate(v){ if(v==null||v==='') return null; if(v instanceof Date) return v; if(typeof v==='string'){ var d=new Date(v); return isNaN(d)?null:d; } return null; }

  // ───────── 甘特视图 ─────────
  var ganttState={ dim:'month', anchorDate:null };
  function ganttRange(dim, anchor, dataMin, dataMax){
    var span={week:42, month:180, quarter:540, year:1460}[dim]||180; // 总天数
    var half=span/2;
    var min=new Date(anchor.getTime()-half*86400000);
    var max=new Date(anchor.getTime()+half*86400000);
    if(dataMin && dataMax){
      if(max<dataMax) max=new Date(dataMax.getTime()+86400000);
      if(min>dataMin) min=new Date(dataMin.getTime());
    }
    return {min:min, max:max};
  }
  function ganttUnitStart(d, dim){
    var x=new Date(d);
    if(dim==='year') return new Date(x.getFullYear(),0,1);
    if(dim==='quarter') return new Date(x.getFullYear(), Math.floor(x.getMonth()/3)*3, 1);
    if(dim==='month') return new Date(x.getFullYear(), x.getMonth(), 1);
    if(dim==='week'){ var dow=(x.getDay()+6)%7; return new Date(x.getTime()-dow*86400000); }
    return new Date(x.getFullYear(), x.getMonth(), x.getDate());
  }
  function ganttStep(x, dim){
    var n=new Date(x);
    if(dim==='year') n.setFullYear(n.getFullYear()+1);
    else if(dim==='quarter') n.setMonth(n.getMonth()+3);
    else if(dim==='month') n.setMonth(n.getMonth()+1);
    else if(dim==='week') n.setDate(n.getDate()+7);
    else n.setDate(n.getDate()+1);
    return n;
  }
  function ganttTickLabel(t, dim){
    if(dim==='year') return ''+t.getFullYear();
    if(dim==='quarter') return t.getFullYear()+'年Q'+(Math.floor(t.getMonth()/3)+1);
    if(dim==='month') return t.getFullYear()+'-'+pad(t.getMonth()+1);
    if(dim==='week') return pad(t.getMonth()+1)+'/'+pad(t.getDate());
    return pad(t.getMonth()+1)+'-'+pad(t.getDate());
  }
  function ganttTicks(min, max, dim){
    var ticks=[]; var t=ganttUnitStart(min, dim); var guard=0;
    while(t<=max && guard<3000){ ticks.push({date:t, pos:(t-min)/86400000/((max-min)/86400000)*100, label:ganttTickLabel(t,dim)}); var nx=ganttStep(t,dim); if(nx<=t) break; t=nx; guard++; }
    return ticks;
  }
  var GANTT_PALETTE=['#4f86f7','#f7824f','#3fb27f','#b56ff7','#f74f86','#f7c14f','#4fc7f7','#8f7bf7','#7fb56f','#f76f6f'];
  var GANTT_CUSTOM_PALETTE=['#3370FF','#14C9C9','#FF7D00','#F76965','#9FDB60','#CA62E8','#FFC53D','#6E4AE0','#86909C','#1D2129'];
  function ganttColorFor(cfg, r){
    if(!cfg) return null;
    if(cfg.colorMode==='custom' && cfg.customColor) return cfg.customColor;
    if(!cfg.colorField) return null;
    var f=state.fields.find(function(x){return x.name===cfg.colorField;}); if(!f) return null;
    var val=displayText(cfg.colorField, r); if(val===null||val===undefined||val==='') return null;
    var idx;
    if(f.type==='Select'||f.type==='Quarter'){ var opts=f.options||[]; idx=opts.indexOf(val); if(idx<0) idx=opts.length; }
    else { idx=0; var sv=''+val; for(var i=0;i<sv.length;i++) idx=(idx+sv.charCodeAt(i))%GANTT_PALETTE.length; }
    return GANTT_PALETTE[idx%GANTT_PALETTE.length];
  }
  function renderGantt(){
    var v=activeView(); var cfg=v.ganttConfig||{};
    if(!cfg.startField){ $('content').innerHTML='<div class="placeholder"><div class="big">▨</div><div>请点击「甘特图配置」选择起止日期字段</div></div>'; return; }
    var rows=state.rows.filter(function(r){ return toDate(r.values[cfg.startField])&&toDate(r.values[cfg.endField||cfg.startField]); });
    if(rows.length===0){ $('content').innerHTML='<div class="placeholder"><div class="big">▨</div><div>没有有效的起止日期数据</div></div>'; return; }
    var dataMin=null,dataMax=null;
    rows.forEach(function(r){ var s=toDate(r.values[cfg.startField]), e=toDate(r.values[cfg.endField||cfg.startField]); if(s&&(!dataMin||s<dataMin))dataMin=s; if(e&&(!dataMax||e>dataMax))dataMax=e; });
    if(!ganttState.anchorDate) ganttState.anchorDate=new Date();
    var dim=ganttState.dim||'month';
    var range=ganttRange(dim, ganttState.anchorDate, dataMin, dataMax);
    var min=range.min, max=range.max;
    var totalDays=Math.max(1,(max-min)/86400000);
    var ticks=ganttTicks(min,max,dim);
    var wrap=document.createElement('div'); wrap.className='gantt-wrap';
    var head=document.createElement('div'); head.className='gantt-head';
    head.innerHTML='<span class="title">甘特图</span><span class="spacer"></span>';
    var todayBtn=document.createElement('button'); todayBtn.className='btn sm'; todayBtn.textContent='今天'; todayBtn.title='回到今天'; todayBtn.addEventListener('click', function(){ ganttState.anchorDate=new Date(); renderGantt(); }); head.appendChild(todayBtn);
    [['week','周'],['month','月'],['quarter','季'],['year','年']].forEach(function(d){ var b=document.createElement('button'); b.className='btn sm'+(dim===d[0]?' active':''); b.textContent=d[1]; b.addEventListener('click', function(){ ganttState.dim=d[0]; renderGantt(); }); head.appendChild(b); });
    wrap.appendChild(head);
    var chart=document.createElement('div'); chart.className='gantt-chart';
    var axis=document.createElement('div'); axis.className='gantt-axis';
    ticks.forEach(function(tk){ var el=document.createElement('div'); el.className='gantt-tick'; el.style.left=tk.pos+'%'; el.textContent=tk.label; axis.appendChild(el); });
    chart.appendChild(axis);
    var grid=document.createElement('div'); grid.className='gantt-grid';
    ticks.forEach(function(tk){ var el=document.createElement('div'); el.className='gantt-gl'; el.style.left=tk.pos+'%'; grid.appendChild(el); });
    chart.appendChild(grid);
    // 今天标识线
    var today=new Date();
    if(today>=min && today<=max){
      var tPos=(today-min)/86400000/totalDays*100;
      var tLine=document.createElement('div'); tLine.className='gantt-today-line';
      tLine.style.left=(168 + tPos*(chart.offsetWidth?chart.offsetWidth-168:0)/100)+'px';
      tLine.style.left='calc(168px + '+tPos+'% - 168px)'; // 更简洁：用百分比但偏移标签区
      tLine.style.left='calc(168px + (100% - 168px) * '+tPos/100+')';
      chart.appendChild(tLine);
    }
    function renderRow(r){
      var s=toDate(r.values[cfg.startField]), e=toDate(r.values[cfg.endField||cfg.startField]);
      var left=(s-min)/86400000/totalDays*100;
      var width=Math.max(1.5,(e-s)/86400000/totalDays*100);
      // 左侧标签优先按字段开关显示可见字段，未设置时回退到标题字段
      var labelFields=(state.visibleFields||[]).filter(function(fn){ return fn!==cfg.startField && fn!==cfg.endField && fn!==cfg.progressField; });
      if(labelFields.length===0){ labelFields=[cfg.labelField||(firstField(['Text'])||{}).name].filter(Boolean); }
      var label=labelFields.map(function(fn){ return displayText(fn,r); }).filter(Boolean).join(' | ');
      if(!label) label='记录 #'+r.rowIndex;
      var row=document.createElement('div'); row.className='gantt-row'+(r.rowIndex===state.selectedRow?' selected':''); row.dataset.row=r.rowIndex;
      var lab=document.createElement('div'); lab.className='gantt-label'; lab.textContent=label; lab.title=label; row.appendChild(lab);
      var track=document.createElement('div'); track.className='gantt-track';
      var bar=document.createElement('div'); bar.className='gantt-bar'; bar.style.left=left+'%'; bar.style.width=width+'%';
      bar.title=label+'\n'+fmtDate(s)+' → '+fmtDate(e);
      var gcolor=ganttColorFor(cfg, r);
      if(gcolor) bar.style.background=gcolor;
      var prog=cfg.progressField?Number(r.values[cfg.progressField]):NaN;
      if(!isNaN(prog)){ var pf=document.createElement('div'); pf.className='gantt-progress'; pf.style.width=Math.max(0,Math.min(100,prog))+'%'; pf.title='进度 '+Math.round(prog)+'%'; if(gcolor) pf.style.background='rgba(255,255,255,0.55)'; bar.appendChild(pf); }
      bar.addEventListener('click', function(){ selectRow(r.rowIndex); showDetailPanel(r.rowIndex); });
      track.appendChild(bar); row.appendChild(track);
      return row;
    }
    if(cfg.groupField){
      var gf=state.fields.find(function(f){return f.name===cfg.groupField;});
      var groups={}; rows.forEach(function(r){ var raw=r.values[cfg.groupField]; var k=(raw==null||raw==='')?'（空）':formatValue(gf, raw); (groups[k]=groups[k]||[]).push(r); });
      var keys=Object.keys(groups);
      if(gf&&(gf.type==='Select'||gf.type==='Quarter')){ var opts=(gf.options||[]).slice(); opts.push('（空）'); keys=opts.filter(function(k){return groups[k];}); Object.keys(groups).forEach(function(k){ if(keys.indexOf(k)<0) keys.push(k); }); }
      keys.forEach(function(k){ var gh=document.createElement('div'); gh.className='gantt-group'; gh.textContent=k+'（'+groups[k].length+'）'; chart.appendChild(gh); groups[k].forEach(function(r){ chart.appendChild(renderRow(r)); }); });
    } else {
      rows.forEach(function(r){ chart.appendChild(renderRow(r)); });
    }
    wrap.appendChild(chart);
    var tip=document.createElement('div'); tip.className='gantt-tip'; tip.textContent='时间范围：'+fmtDate(min)+' 至 '+fmtDate(max)+'（共 '+rows.length+' 条任务，点击条可编辑）';
    wrap.appendChild(tip);
    $('content').innerHTML=''; $('content').appendChild(wrap);
  }
  function fmtDate(d){ return d? (d.getFullYear()+'-'+pad(d.getMonth()+1)+'-'+pad(d.getDate())) : ''; }
  function pad(n){ return n<10?'0'+n:''+n; }

  // ───────── 仪表盘 / 图表 ─────────
  function aggregate(rows, field, mode){
    var vals=rows.map(function(r){ var v=r.values[field]; var n=Number(v); return isNaN(n)?null:n; }).filter(function(x){return x!=null;});
    if(mode==='Count') return rows.length;
    if(vals.length===0) return 0;
    if(mode==='Average') return vals.reduce(function(a,b){return a+b;},0)/vals.length;
    if(mode==='Max') return Math.max.apply(null,vals);
    if(mode==='Min') return Math.min.apply(null,vals);
    if(mode==='DistinctCount'){ var s={}; vals.forEach(function(x){s[x]=1;}); return Object.keys(s).length; }
    return vals.reduce(function(a,b){return a+b;},0); // Sum
  }
  function fmtNum(n, format){
    if(format==='money'||format==='percent') {} // 仅简单处理
    if(Math.abs(n)>=10000) return (n/10000).toFixed(1)+'万';
    return (Math.round(n*100)/100).toLocaleString('zh-CN');
  }
  function renderDashboard(){
    var v=activeView(); var cfg=v.dashboardConfig||{statCards:[],charts:[],columns:2};
    var wrap=document.createElement('div'); wrap.className='dash-wrap';
    var kg=document.createElement('div'); kg.className='kpi-grid';
    (cfg.statCards||[]).forEach(function(k){
      var val=aggregate(state.rows, k.field, k.aggregation);
      var card=document.createElement('div'); card.className='kpi-card'; card.style.borderTop='3px solid '+(k.color||'#3370FF');
      card.innerHTML='<div style="display:flex;align-items:center;justify-content:space-between;"><div class="kpi-title">'+escapeHtml(k.title||k.field)+'</div><button class="btn sm" data-kpi="'+k.id+'">设置</button></div><div class="kpi-val">'+escapeHtml(fmtNum(val,k.format))+'</div>';
      card.querySelector('button').addEventListener('click', function(e){ e.stopPropagation(); editDashboardItem('kpi', k.id); });
      kg.appendChild(card);
    });
    wrap.appendChild(kg);
    var cg=document.createElement('div'); cg.className='chart-grid';
    (cfg.charts||[]).forEach(function(ch){
      var box=buildChartBox(ch);
      var h4=box.querySelector('h4'); if(h4) h4.remove();
      var head=document.createElement('div'); head.style.cssText='display:flex;align-items:center;justify-content:space-between;margin-bottom:6px;';
      var t=document.createElement('h4'); t.style.margin='0'; t.textContent=ch.title||'图表'; head.appendChild(t);
      var btns=document.createElement('div'); btns.style.cssText='display:flex;gap:6px;';
      var set=document.createElement('button'); set.className='btn sm'; set.textContent='设置'; set.addEventListener('click', function(e){ e.stopPropagation(); editDashboardItem('chart', ch.id); });
      var del=document.createElement('button'); del.className='btn sm'; del.textContent='删除'; del.addEventListener('click', function(e){ e.stopPropagation(); if(confirm('确定删除图表「'+(ch.title||'图表')+'」？')){ var v=activeView(); v.dashboardConfig.charts=(v.dashboardConfig.charts||[]).filter(function(x){return x.id!==ch.id;}); saveConfig(); renderContent(); toast('已删除'); } });
      btns.appendChild(set); btns.appendChild(del); head.appendChild(btns);
      box.insertBefore(head, box.firstChild);
      cg.appendChild(box);
    });
    wrap.appendChild(cg);
    if((cfg.statCards||[]).length===0 && (cfg.charts||[]).length===0){ wrap.innerHTML='<div class="placeholder"><div class="big">▣</div><div>仪表盘尚未配置，点击右上角「＋ 指标 / ＋ 图表」添加</div></div>'; }
    $('content').innerHTML=''; $('content').appendChild(wrap);
  }
  function addDashboardItem(kind){
    var v=activeView(); v.dashboardConfig=v.dashboardConfig||{statCards:[],charts:[],columns:2};
    if(kind==='kpi'){ var num=firstField(['Number','Integer','Currency','Percentage']); var id='k'+Date.now(); v.dashboardConfig.statCards.push({id:id,title:num?num.name+' 合计':'指标',field:num?num.name:'',aggregation:'Sum',format:'auto',color:''}); saveConfig(); renderContent(); editDashboardItem('kpi', id); }
    else { var dim=firstField(['Select','Quarter','Text']); var num2=firstField(['Number','Integer','Currency','Percentage']); var id='c'+Date.now(); v.dashboardConfig.charts.push({id:id,title:'图表',type:'Column',dimensionField:dim?dim.name:'',metricField:num2?num2.name:'',aggregation:'Sum',timeField:'',timeGroup:'None',seriesField:'',topN:12,gaugeTarget:100,columnSpan:1,height:260,showLabels:true}); saveConfig(); renderContent(); editDashboardItem('chart', id); }
  }
  function editDashboardItem(kind, id){
    var v=activeView(); var cfg=v.dashboardConfig=v.dashboardConfig||{statCards:[],charts:[]}; var item;
    if(kind==='kpi') item=(cfg.statCards||[]).find(function(x){return x.id===id;});
    else item=(cfg.charts||[]).find(function(x){return x.id===id;});
    if(!item) return;
    var ov=document.createElement('div'); ov.className='overlay';
    var modal=document.createElement('div'); modal.className='modal';
    modal.innerHTML='<div class="modal-head"><span>'+(kind==='kpi'?'指标设置':'图表设置')+'</span><button class="btn sm" id="xDashSet">✕</button></div>';
    var body=document.createElement('div'); body.className='modal-body';
    function fieldSelect(label, value, types){
      var row=document.createElement('div'); row.className='setting-row';
      var lab=document.createElement('label'); lab.textContent=label; row.appendChild(lab);
      var sel=document.createElement('select'); sel.appendChild(new Option('（无）',''));
      state.fields.forEach(function(f){ if(!types || types.indexOf(f.type)>=0){ sel.appendChild(new Option(f.name, f.name)); } });
      sel.value=value||''; row.appendChild(sel); return {row:row, sel:sel};
    }
    function textRow(label, value){
      var row=document.createElement('div'); row.className='setting-row';
      var lab=document.createElement('label'); lab.textContent=label; row.appendChild(lab);
      var inp=document.createElement('input'); inp.type='text'; inp.value=value||''; row.appendChild(inp); return {row:row, inp:inp};
    }
    var titleRow=textRow(kind==='kpi'?'指标名称':'图表标题', item.title);
    body.appendChild(titleRow.row);
    if(kind==='kpi'){
      var fsel=fieldSelect('统计字段', item.field, ['Number','Integer','Currency','Percentage']); body.appendChild(fsel.row);
      var AGG_NAMES={'Sum':'求和','Count':'计数','Average':'平均','Max':'最大','Min':'最小','DistinctCount':'去重计数'};
      var arow=document.createElement('div'); arow.className='setting-row';
      var alab=document.createElement('label'); alab.textContent='聚合方式'; arow.appendChild(alab);
      var asel=document.createElement('select'); Object.keys(AGG_NAMES).forEach(function(x){ asel.appendChild(new Option(AGG_NAMES[x], x)); }); asel.value=item.aggregation||'Sum'; arow.appendChild(asel); body.appendChild(arow);
      var fmtRow=textRow('显示格式', item.format||''); fmtRow.inp.placeholder='auto / money / percent / int'; body.appendChild(fmtRow.row);
      var colorRow=textRow('颜色', item.color||''); colorRow.inp.placeholder='#3370FF'; body.appendChild(colorRow.row);
      body.appendChild(makeApply(function(){ item.title=titleRow.inp.value; item.field=fsel.sel.value; item.aggregation=asel.value; item.format=fmtRow.inp.value.trim(); item.color=colorRow.inp.value.trim(); }));
    } else {
      var CHART_TYPES={'Column':'柱状图','Bar':'条形图','Line':'折线图','Area':'面积图','Pie':'饼图','Doughnut':'环形图'};
      var trow=document.createElement('div'); trow.className='setting-row';
      var tlab=document.createElement('label'); tlab.textContent='图表类型'; trow.appendChild(tlab);
      var tsel=document.createElement('select'); Object.keys(CHART_TYPES).forEach(function(x){ tsel.appendChild(new Option(CHART_TYPES[x], x)); }); tsel.value=item.type||'Column'; trow.appendChild(tsel); body.appendChild(trow);
      var dsel=fieldSelect('维度字段', item.dimensionField, ['Select','Quarter','Text','Date','DateTime']); body.appendChild(dsel.row);
      var msel=fieldSelect('度量字段', item.metricField, ['Number','Integer','Currency','Percentage']); body.appendChild(msel.row);
      var AGG_NAMES={'Sum':'求和','Count':'计数','Average':'平均','Max':'最大','Min':'最小','DistinctCount':'去重计数'};
      var arow=document.createElement('div'); arow.className='setting-row';
      var alab=document.createElement('label'); alab.textContent='聚合方式'; arow.appendChild(alab);
      var asel=document.createElement('select'); Object.keys(AGG_NAMES).forEach(function(x){ asel.appendChild(new Option(AGG_NAMES[x], x)); }); asel.value=item.aggregation||'Sum'; arow.appendChild(asel); body.appendChild(arow);
      var TG_NAMES={'None':'无','Year':'按年','Quarter':'按季度','Month':'按月','Week':'按周','Day':'按日'};
      var tgrow=document.createElement('div'); tgrow.className='setting-row';
      var tglab=document.createElement('label'); tglab.textContent='时间维度'; tgrow.appendChild(tglab);
      var tgsel=document.createElement('select'); Object.keys(TG_NAMES).forEach(function(x){ tgsel.appendChild(new Option(TG_NAMES[x], x)); }); tgsel.value=item.timeGroup||'None'; tgrow.appendChild(tgsel); body.appendChild(tgrow);
      var hrow=textRow('高度', String(item.height||260)); hrow.inp.type='number'; body.appendChild(hrow.row);
      var showRow=document.createElement('div'); showRow.className='setting-row';
      var showLab=document.createElement('label'); showLab.textContent='显示标签'; showRow.appendChild(showLab);
      var showWrap=document.createElement('div'); showWrap.style.cssText='flex:1;display:flex;align-items:center;';
      var showCb=document.createElement('input'); showCb.type='checkbox'; showCb.checked=item.showLabels!==false; showCb.style.cssText='width:16px;height:16px;margin:0';
      showWrap.appendChild(showCb); showRow.appendChild(showWrap); body.appendChild(showRow);
      body.appendChild(makeApply(function(){ item.title=titleRow.inp.value; item.type=tsel.value; item.dimensionField=dsel.sel.value; item.metricField=msel.sel.value; item.aggregation=asel.value; item.timeGroup=tgsel.value; item.height=parseInt(hrow.inp.value)||260; item.showLabels=showCb.checked; }));
    }
    modal.appendChild(body);
    var foot=document.createElement('div'); foot.className='modal-foot';
    var del=document.createElement('button'); del.className='btn sm'; del.textContent='删除';
    del.addEventListener('click', function(){ if(kind==='kpi') cfg.statCards=(cfg.statCards||[]).filter(function(x){return x.id!==id;}); else cfg.charts=(cfg.charts||[]).filter(function(x){return x.id!==id;}); saveConfig(); renderContent(); ov.remove(); });
    var close=document.createElement('button'); close.className='btn'; close.textContent='关闭'; close.addEventListener('click', function(){ ov.remove(); });
    foot.appendChild(del); foot.appendChild(close); modal.appendChild(foot);
    ov.appendChild(modal); ov.addEventListener('click', function(e){ if(e.target===ov) ov.remove(); });
    document.getElementById('overlayHost').appendChild(ov);
    modal.querySelector('#xDashSet').addEventListener('click', function(){ ov.remove(); });
  }
  function renderChart(v){
    var cfg=v.chartConfig; if(!cfg){ $('content').innerHTML='<div class="placeholder"><div class="big">▤</div><div>图表未配置</div></div>'; return; }
    var wrap=document.createElement('div'); wrap.className='dash-wrap';
    var box=buildChartBox(cfg);
    var head=document.createElement('div'); head.style.cssText='display:flex;align-items:center;justify-content:space-between;margin-bottom:6px;';
    var t=document.createElement('h4'); t.style.margin='0'; t.textContent=cfg.title||'图表'; head.appendChild(t);
    var btns=document.createElement('div'); btns.style.cssText='display:flex;gap:6px;';
    var set=document.createElement('button'); set.className='btn sm'; set.textContent='设置'; set.addEventListener('click', function(){ showViewSettings(); });
    var del=document.createElement('button'); del.className='btn sm'; del.textContent='删除'; del.addEventListener('click', function(){ if(confirm('确定删除当前统计图表视图？')){ var idx=state.views.findIndex(function(x){return x.viewId===v.viewId;}); if(idx>=0){ state.views.splice(idx,1); saveConfig(); renderSidebar(); renderContent(); toast('已删除'); } } });
    btns.appendChild(set); btns.appendChild(del); head.appendChild(btns);
    box.insertBefore(head, box.firstChild);
    wrap.appendChild(box);
    $('content').innerHTML=''; $('content').appendChild(wrap);
  }
  function buildChartBox(cfg){
    var box=document.createElement('div'); box.className='chart-box';
    box.innerHTML='<h4>'+escapeHtml(cfg.title||'图表')+'</h4>';
    var svg=document.createElementNS('http://www.w3.org/2000/svg','svg'); svg.setAttribute('width','100%'); svg.setAttribute('height',(cfg.height||260));
    drawChart(svg, cfg);
    box.appendChild(svg);
    return box;
  }
  function drawChart(svg, cfg){
    var W=360, H=cfg.height||260, padL=40, padB=28, padT=10, padR=10;
    while(svg.firstChild) svg.removeChild(svg.firstChild);
    // 分组聚合
    var groups=groupAggregate(state.rows, cfg);
    if(groups.length===0){ var t=document.createElementNS('http://www.w3.org/2000/svg','text'); t.setAttribute('x',W/2); t.setAttribute('y',H/2); t.setAttribute('text-anchor','middle'); t.textContent='无数据'; svg.appendChild(t); return; }
    var colors=['#3370FF','#14C9C9','#FF7D00','#F76965','#9FDB60','#CA62E8','#FFC53D','#6E4AE0'];
    var maxV=Math.max.apply(null, groups.map(function(g){return g.value;}).concat([1]));
    if(cfg.type==='Pie'||cfg.type==='Doughnut'){
      var cx=W/2, cy=H/2, r=Math.min(W,H)/2-20, total=groups.reduce(function(a,b){return a+b.value;},0)||1;
      var ang=-Math.PI/2;
      groups.forEach(function(g,i){
        var a2=ang+g.value/total*Math.PI*2;
        var path=document.createElementNS('http://www.w3.org/2000/svg','path');
        var x1=cx+r*Math.cos(ang), y1=cy+r*Math.sin(ang), x2=cx+r*Math.cos(a2), y2=cy+r*Math.sin(a2);
        var large=(a2-ang)>Math.PI?1:0;
        var d='M '+cx+' '+cy+' L '+x1+' '+y1+' A '+r+' '+r+' 0 '+large+' 1 '+x2+' '+y2+' Z';
        path.setAttribute('d',d); path.setAttribute('fill',colors[i%colors.length]); path.setAttribute('stroke','#fff');
        svg.appendChild(path); ang=a2;
      });
      if(cfg.type==='Doughnut'){
        var hole=document.createElementNS('http://www.w3.org/2000/svg','circle');
        hole.setAttribute('cx',cx); hole.setAttribute('cy',cy); hole.setAttribute('r',r*0.55);
        hole.setAttribute('fill','#fff');
        svg.appendChild(hole);
      }
      // 图例
      var ly=padT;
      groups.slice(0,8).forEach(function(g,i){ var lg=document.createElementNS('http://www.w3.org/2000/svg','text'); lg.setAttribute('x',10); lg.setAttribute('y',ly+ (i+1)*16); lg.setAttribute('font-size','11'); lg.textContent=g.label+' ('+fmtNum(g.value)+')'; svg.appendChild(lg); });
      return;
    }
    var plotW=W-padL-padR, plotH=H-padT-padB;
    var n=groups.length;
    var drawPadL=padL, drawPlotW=plotW;
    if(cfg.type==='Bar'){ drawPadL=70; drawPlotW=W-drawPadL-padR; } // 横向条形需要更宽左侧标签区
    var bw=drawPlotW/n;
    // 轴线
    var axis=document.createElementNS('http://www.w3.org/2000/svg','line'); axis.setAttribute('x1',drawPadL); axis.setAttribute('y1',padT+plotH); axis.setAttribute('x2',W-padR); axis.setAttribute('y2',padT+plotH); axis.setAttribute('stroke','#C9CDD4'); svg.appendChild(axis);
    if(cfg.type==='Bar'){
      var yaxis=document.createElementNS('http://www.w3.org/2000/svg','line'); yaxis.setAttribute('x1',drawPadL); yaxis.setAttribute('y1',padT); yaxis.setAttribute('x2',drawPadL); yaxis.setAttribute('y2',padT+plotH); yaxis.setAttribute('stroke','#C9CDD4'); svg.appendChild(yaxis);
    }
    groups.forEach(function(g,i){
      if(cfg.type==='Bar'){ // 横向
        var band=plotH/n;
        var bh=band*0.55;
        var yy=padT+i*band+band*0.5;
        var r=document.createElementNS('http://www.w3.org/2000/svg','rect'); r.setAttribute('x',drawPadL); r.setAttribute('y',yy-bh/2); r.setAttribute('height',bh); r.setAttribute('width',drawPlotW*g.value/maxV); r.setAttribute('fill',colors[i%colors.length]); svg.appendChild(r);
        if(cfg.showLabels!==false){
          var lbl=document.createElementNS('http://www.w3.org/2000/svg','text'); lbl.setAttribute('x',drawPadL-6); lbl.setAttribute('y',yy+4); lbl.setAttribute('text-anchor','end'); lbl.setAttribute('font-size','10'); lbl.textContent=String(g.label).slice(0,10); svg.appendChild(lbl);
          var valLbl=document.createElementNS('http://www.w3.org/2000/svg','text'); valLbl.setAttribute('x',drawPadL+drawPlotW*g.value/maxV+4); valLbl.setAttribute('y',yy+4); valLbl.setAttribute('font-size','9'); valLbl.textContent=fmtNum(g.value); svg.appendChild(valLbl);
        }
      } else if(cfg.type==='Line'||cfg.type==='Area'){
        // 折线在循环后统一画
      } else {
        var x=drawPadL+i*bw+bw/2;
        var h=plotH*g.value/maxV, y=padT+plotH-h;
        var rect=document.createElementNS('http://www.w3.org/2000/svg','rect'); rect.setAttribute('x',x-bw*0.3); rect.setAttribute('y',y); rect.setAttribute('width',bw*0.6); rect.setAttribute('height',h); rect.setAttribute('fill',colors[i%colors.length]); svg.appendChild(rect);
        if(cfg.showLabels!==false){
          var lbl=document.createElementNS('http://www.w3.org/2000/svg','text'); lbl.setAttribute('x',x); lbl.setAttribute('y',padT+plotH+16); lbl.setAttribute('text-anchor','middle'); lbl.setAttribute('font-size','10'); lbl.textContent=String(g.label).slice(0,8); svg.appendChild(lbl);
          var valLbl=document.createElementNS('http://www.w3.org/2000/svg','text'); valLbl.setAttribute('x',x); valLbl.setAttribute('y',y-4); valLbl.setAttribute('text-anchor','middle'); valLbl.setAttribute('font-size','9'); valLbl.textContent=fmtNum(g.value); svg.appendChild(valLbl);
        }
      }
    });
    if(cfg.type==='Line'||cfg.type==='Area'){
      var pts=groups.map(function(g,i){ return [drawPadL+i*bw+bw/2, padT+plotH-plotH*g.value/maxV]; });
      if(cfg.type==='Area'){ var ar=document.createElementNS('http://www.w3.org/2000/svg','polygon'); ar.setAttribute('points', drawPadL+','+(padT+plotH)+' '+pts.map(function(p){return p[0]+','+p[1];}).join(' ')+' '+(drawPadL+drawPlotW)+','+(padT+plotH)); ar.setAttribute('fill','rgba(51,112,255,.15)'); svg.appendChild(ar); }
      var pl=document.createElementNS('http://www.w3.org/2000/svg','polyline'); pl.setAttribute('points',pts.map(function(p){return p[0]+','+p[1];}).join(' ')); pl.setAttribute('fill','none'); pl.setAttribute('stroke','#3370FF'); pl.setAttribute('stroke-width','2'); svg.appendChild(pl);
    }
  }
  function groupAggregate(rows, cfg){
    var field=cfg.dimensionField;
    if(cfg.timeField && (cfg.timeGroup&&cfg.timeGroup!=='None')) field=cfg.timeField;
    if(!field && cfg.timeGroup && cfg.timeGroup!=='None'){ var dateField=firstField(['Date','DateTime']); if(dateField) field=dateField.name; }
    if(!field) return rows.map(function(r,i){ return {label:'#'+i, value:Number(r.values[cfg.metricField]||0)||0}; });
    var map={};
    rows.forEach(function(r){
      var key;
      var dv=toDate(r.values[field]);
      if(dv){ key=fmtDate(dv); if(cfg.timeGroup==='Month') key=dv.getFullYear()+'-'+pad(dv.getMonth()+1); else if(cfg.timeGroup==='Year') key=dv.getFullYear(); else if(cfg.timeGroup==='Quarter') key=dv.getFullYear()+'Q'+Math.floor(dv.getMonth()/3+1); }
      else key=String(r.values[field]||'（空）');
      map[key]=(map[key]||0)+1;
    });
    // 聚合度量
    var out=Object.keys(map).map(function(k){
      var sub=rows.filter(function(r){ var dv=toDate(r.values[field]); var kk; if(dv){ kk=fmtDate(dv); if(cfg.timeGroup==='Month') kk=dv.getFullYear()+'-'+pad(dv.getMonth()+1); else if(cfg.timeGroup==='Year') kk=dv.getFullYear(); else if(cfg.timeGroup==='Quarter') kk=dv.getFullYear()+'Q'+Math.floor(dv.getMonth()/3+1); } else kk=String(r.values[field]||'（空）'); return kk===k; });
      return {label:k, value:aggregate(sub, cfg.metricField, cfg.aggregation)};
    });
    if(cfg.topN && out.length>cfg.topN){ out.sort(function(a,b){return b.value-a.value;}); var rest=out.slice(cfg.topN-1); out=out.slice(0,cfg.topN-1); var sum=rest.reduce(function(a,b){return a+b.value;},0); out.push({label:'其他', value:sum}); }
    return out;
  }

  // ───────── 日历设置弹窗（保留原来视图设置中的两个功能） ─────────
  function showCalendarSettings(){
    var v=activeView(); if(!v) return;
    var cfg=v.calendarConfig||{};
    var ov=document.createElement('div'); ov.className='overlay';
    var modal=document.createElement('div'); modal.className='modal';
    modal.innerHTML='<div class="modal-head"><span>日历设置</span><button class="btn sm" id="xCalSet">✕</button></div>';
    var body=document.createElement('div'); body.className='modal-body';
    function fieldSelect(label, value, types){
      var row=document.createElement('div'); row.className='setting-row';
      var lab=document.createElement('label'); lab.textContent=label; row.appendChild(lab);
      var sel=document.createElement('select'); sel.appendChild(new Option('（无）',''));
      state.fields.forEach(function(f){ if(!types || types.indexOf(f.type)>=0){ sel.appendChild(new Option(f.name, f.name)); } });
      sel.value=value||''; row.appendChild(sel); return {row:row, sel:sel};
    }
    var cd=fieldSelect('开始日期字段', cfg.dateField, ['Date','DateTime']); body.appendChild(cd.row);
    var cde=fieldSelect('结束日期字段（可选）', cfg.endDateField, ['Date','DateTime']); body.appendChild(cde.row);
    var ct=fieldSelect('标题字段', cfg.titleField, null); body.appendChild(ct.row);
    body.appendChild(makeApply(function(){ v.calendarConfig=v.calendarConfig||{}; v.calendarConfig.dateField=cd.sel.value; v.calendarConfig.endDateField=cde.sel.value; v.calendarConfig.titleField=ct.sel.value; }));
    modal.appendChild(body);
    var foot=document.createElement('div'); foot.className='modal-foot';
    var close=document.createElement('button'); close.className='btn'; close.textContent='关闭'; close.addEventListener('click', function(){ ov.remove(); });
    foot.appendChild(close); modal.appendChild(foot);
    ov.appendChild(modal); ov.addEventListener('click', function(e){ if(e.target===ov) ov.remove(); });
    document.getElementById('overlayHost').appendChild(ov);
    modal.querySelector('#xCalSet').addEventListener('click', function(){ ov.remove(); });
  }

  // ───────── 甘特图配置弹窗 ─────────
  function showGanttSettings(){
    var v=activeView(); if(!v) return;
    var cfg=v.ganttConfig=v.ganttConfig||{};
    var ov=document.createElement('div'); ov.className='overlay';
    var modal=document.createElement('div'); modal.className='modal';
    modal.innerHTML='<div class="modal-head"><span>甘特图配置</span><button class="btn sm" id="xGanttSet">✕</button></div>';
    var body=document.createElement('div'); body.className='modal-body';
    function fieldSelect(label, value, types){
      var row=document.createElement('div'); row.className='setting-row';
      var lab=document.createElement('label'); lab.textContent=label; row.appendChild(lab);
      var sel=document.createElement('select'); sel.appendChild(new Option('（无）',''));
      state.fields.forEach(function(f){ if(!types || types.indexOf(f.type)>=0){ sel.appendChild(new Option(f.name, f.name)); } });
      sel.value=value||''; row.appendChild(sel); return {row:row, sel:sel};
    }
    var gs=fieldSelect('开始日期', cfg.startField, ['Date','DateTime']); body.appendChild(gs.row);
    var ge=fieldSelect('结束日期', cfg.endField, ['Date','DateTime']); body.appendChild(ge.row);
    var gl=fieldSelect('标题显示字段', cfg.labelField, null); body.appendChild(gl.row);
    // 颜色模式
    var cmRow=document.createElement('div'); cmRow.className='setting-row';
    var cmLab=document.createElement('label'); cmLab.textContent='颜色显示'; cmRow.appendChild(cmLab);
    var cmSel=document.createElement('select');
    cmSel.appendChild(new Option('自定义颜色','custom'));
    cmSel.appendChild(new Option('按字段分色','field'));
    cmSel.value=cfg.colorMode||'field'; cmRow.appendChild(cmSel); body.appendChild(cmRow);
    var cf=fieldSelect('分色字段', cfg.colorField, ['Select','Quarter','Text']); body.appendChild(cf.row);
    // 自定义颜色选择
    var cpRow=document.createElement('div'); cpRow.className='setting-row';
    var cpLab=document.createElement('label'); cpLab.textContent='自定义颜色'; cpRow.appendChild(cpLab);
    var cpWrap=document.createElement('div'); cpWrap.style.cssText='flex:1;display:flex;gap:6px;flex-wrap:wrap;align-items:center;';
    var cpInput=document.createElement('input'); cpInput.type='text'; cpInput.value=cfg.customColor||'#3370FF'; cpInput.style.cssText='flex:1;min-width:80px;height:28px;';
    cpWrap.appendChild(cpInput);
    GANTT_CUSTOM_PALETTE.forEach(function(c){
      var sq=document.createElement('span'); sq.style.cssText='width:20px;height:20px;border-radius:4px;background:'+c+';cursor:pointer;border:1px solid var(--border);';
      sq.addEventListener('click', function(){ cpInput.value=c; });
      cpWrap.appendChild(sq);
    });
    cpRow.appendChild(cpWrap); body.appendChild(cpRow);
    function updateColorUI(){ var isCustom=cmSel.value==='custom'; cf.row.style.display=isCustom?'none':'flex'; cpRow.style.display=isCustom?'flex':'none'; }
    cmSel.addEventListener('change', updateColorUI); updateColorUI();
    // 仅计算工作日
    var wdRow=document.createElement('div'); wdRow.className='setting-row';
    var wdLab=document.createElement('label'); wdLab.textContent='仅计算工作日'; wdRow.appendChild(wdLab);
    var wdWrap=document.createElement('div'); wdWrap.style.cssText='flex:1;display:flex;align-items:center;';
    var wdCb=document.createElement('input'); wdCb.type='checkbox'; wdCb.checked=!!cfg.workdaysOnly; wdCb.style.cssText='width:16px;height:16px;margin:0';
    wdWrap.appendChild(wdCb); wdRow.appendChild(wdWrap); body.appendChild(wdRow);
    body.appendChild(makeApply(function(){
      v.ganttConfig.startField=gs.sel.value; v.ganttConfig.endField=ge.sel.value; v.ganttConfig.labelField=gl.sel.value;
      v.ganttConfig.colorMode=cmSel.value; v.ganttConfig.colorField=cf.sel.value; v.ganttConfig.customColor=cpInput.value.trim()||'#3370FF'; v.ganttConfig.workdaysOnly=wdCb.checked;
    }));
    modal.appendChild(body);
    var foot=document.createElement('div'); foot.className='modal-foot';
    var close=document.createElement('button'); close.className='btn'; close.textContent='关闭'; close.addEventListener('click', function(){ ov.remove(); });
    foot.appendChild(close); modal.appendChild(foot);
    ov.appendChild(modal); ov.addEventListener('click', function(e){ if(e.target===ov) ov.remove(); });
    document.getElementById('overlayHost').appendChild(ov);
    modal.querySelector('#xGanttSet').addEventListener('click', function(){ ov.remove(); });
  }

  // ───────── 视图设置面板 ─────────
  function showViewSettings(){
    var v=activeView(); if(!v) return;
    var t=v.viewType;
    var ov=document.createElement('div'); ov.className='overlay';
    var modal=document.createElement('div'); modal.className='modal';
    modal.innerHTML='<div class="modal-head"><span>「'+(v.viewName||VIEW_LABEL[t]||t)+'」设置</span><button class="btn sm" id="xSet">✕</button></div>';
    var body=document.createElement('div'); body.className='modal-body';
    function fieldSelect(label, value, types){
      var row=document.createElement('div'); row.className='setting-row';
      var lab=document.createElement('label'); lab.textContent=label; row.appendChild(lab);
      var sel=document.createElement('select');
      sel.appendChild(new Option('（无）',''));
      state.fields.forEach(function(f){ if(!types || types.indexOf(f.type)>=0){ sel.appendChild(new Option(f.name, f.name)); } });
      sel.value=value||''; row.appendChild(sel); return {row:row, sel:sel};
    }
    if(t==='Kanban'){
      var tf=fieldSelect('卡片标题', v.cardMeta&&v.cardMeta.title, null); body.appendChild(tf.row);
      body.appendChild(makeApply(function(){ v.cardMeta=v.cardMeta||{}; v.cardMeta.title=tf.sel.value; v.cardMeta.description=[]; }));
    } else if(t==='Gallery'){
      var gi=fieldSelect('图片字段', v.cardMeta&&v.cardMeta.image, ['Image']); body.appendChild(gi.row);
      var gt=fieldSelect('标题字段', v.cardMeta&&v.cardMeta.title, null); body.appendChild(gt.row);
      var gd=fieldSelect('描述字段', (v.cardMeta&&v.cardMeta.description&&v.cardMeta.description[0])||'', null); body.appendChild(gd.row);
      body.appendChild(makeApply(function(){ v.cardMeta=v.cardMeta||{}; v.cardMeta.image=gi.sel.value; v.cardMeta.title=gt.sel.value; v.cardMeta.description=[gd.sel.value].filter(Boolean); }));
    } else if(t==='Calendar'){
      var cd=fieldSelect('日期字段', v.calendarConfig&&v.calendarConfig.dateField, ['Date','DateTime']); body.appendChild(cd.row);
      var ct=fieldSelect('标题字段', v.calendarConfig&&v.calendarConfig.titleField, null); body.appendChild(ct.row);
      body.appendChild(makeApply(function(){ v.calendarConfig=v.calendarConfig||{}; v.calendarConfig.dateField=cd.sel.value; v.calendarConfig.titleField=ct.sel.value; }));
    } else if(t==='Gantt'){
      var gs=fieldSelect('开始日期', v.ganttConfig&&v.ganttConfig.startField, ['Date','DateTime']); body.appendChild(gs.row);
      var ge=fieldSelect('结束日期', v.ganttConfig&&v.ganttConfig.endField, ['Date','DateTime']); body.appendChild(ge.row);
      var gl=fieldSelect('标签字段', v.ganttConfig&&v.ganttConfig.labelField, null); body.appendChild(gl.row);
      var gp=fieldSelect('进度字段', v.ganttConfig&&v.ganttConfig.progressField, ['Number','Integer','Currency','Percentage']); body.appendChild(gp.row);
      var gg=fieldSelect('分组字段', v.ganttConfig&&v.ganttConfig.groupField, ['Select','Quarter','Text']); body.appendChild(gg.row);
      var gco=fieldSelect('着色字段（按状态分色）', v.ganttConfig&&v.ganttConfig.colorField, ['Select','Quarter','Text']); body.appendChild(gco.row);
      body.appendChild(makeApply(function(){ v.ganttConfig=v.ganttConfig||{}; v.ganttConfig.startField=gs.sel.value; v.ganttConfig.endField=ge.sel.value; v.ganttConfig.labelField=gl.sel.value; v.ganttConfig.progressField=gp.sel.value; v.ganttConfig.groupField=gg.sel.value; v.ganttConfig.colorField=gco.sel.value; }));
    } else if(t==='Chart'){
      var chTitle=document.createElement('div'); chTitle.style.cssText='font-weight:600;margin-bottom:6px'; chTitle.textContent='图表配置'; body.appendChild(chTitle);
      var CHART_TYPES={'Column':'柱状图','Bar':'条形图','Line':'折线图','Area':'面积图','Pie':'饼图','Doughnut':'环形图'};
      var cc=fieldSelect('类型', '', null); cc.sel.innerHTML=''; Object.keys(CHART_TYPES).forEach(function(x){ cc.sel.appendChild(new Option(CHART_TYPES[x], x)); }); cc.sel.value=v.chartConfig&&v.chartConfig.type||'Column'; body.appendChild(cc.row);
      var cd2=fieldSelect('维度字段', v.chartConfig&&v.chartConfig.dimensionField, ['Select','Quarter','Text','Date','DateTime']); body.appendChild(cd2.row);
      var cm=fieldSelect('度量字段', v.chartConfig&&v.chartConfig.metricField, ['Number','Integer','Currency','Percentage']); body.appendChild(cm.row);
      var AGG_NAMES={'Sum':'求和','Count':'计数','Average':'平均','Max':'最大','Min':'最小','DistinctCount':'去重计数'};
      var ca=fieldSelect('聚合方式', '', null); ca.sel.innerHTML=''; Object.keys(AGG_NAMES).forEach(function(x){ ca.sel.appendChild(new Option(AGG_NAMES[x], x)); }); ca.sel.value=v.chartConfig&&v.chartConfig.aggregation||'Sum'; body.appendChild(ca.row);
      var TG_NAMES={'None':'无','Year':'按年','Quarter':'按季度','Month':'按月','Week':'按周','Day':'按日'};
      var ctg=fieldSelect('时间维度', v.chartConfig&&v.chartConfig.timeGroup, null); ctg.sel.innerHTML=''; Object.keys(TG_NAMES).forEach(function(x){ ctg.sel.appendChild(new Option(TG_NAMES[x], x)); }); ctg.sel.value=v.chartConfig&&v.chartConfig.timeGroup||'None'; body.appendChild(ctg.row);
      body.appendChild(makeApply(function(){ v.chartConfig=v.chartConfig||{}; v.chartConfig.type=cc.sel.value; v.chartConfig.dimensionField=cd2.sel.value; v.chartConfig.metricField=cm.sel.value; v.chartConfig.aggregation=ca.sel.value; v.chartConfig.timeGroup=ctg.sel.value; if(!v.chartConfig.title) v.chartConfig.title='图表'; }));
    }
    modal.appendChild(body);
    var foot=document.createElement('div'); foot.className='modal-foot';
    var close=document.createElement('button'); close.className='btn'; close.textContent='关闭'; close.addEventListener('click', function(){ ov.remove(); });
    foot.appendChild(close); modal.appendChild(foot);
    ov.appendChild(modal); ov.addEventListener('click', function(e){ if(e.target===ov) ov.remove(); });
    document.getElementById('overlayHost').appendChild(ov);
    modal.querySelector('#xSet').addEventListener('click', function(){ ov.remove(); });
  }

  // ───────── 字段设置面板（右侧抽屉） ─────────
  function closeSettingsPanel(){ var sp=$('settingsPanel'); if(sp){ sp.classList.remove('open'); sp.innerHTML=''; } }
  function showFieldSettings(){
    closeDetailPanel();
    var sp=$('settingsPanel'); sp.innerHTML=''; sp.classList.add('open');
    var head=document.createElement('div'); head.className='dp-head';
    head.innerHTML='<span>字段设置</span>';
    var x=document.createElement('button'); x.className='btn sm'; x.textContent='✕';
    x.addEventListener('click', function(){ saveConfig(); closeSettingsPanel(); });
    head.appendChild(x); sp.appendChild(head);
    var body=document.createElement('div'); body.className='dp-body';
    // 字段选择组合框
    var selWrap=document.createElement('div'); selWrap.className='form-field';
    var selLab=document.createElement('label'); selLab.textContent='字段'; selWrap.appendChild(selLab);
    var selCtrl=document.createElement('div'); selCtrl.className='ctrl';
    var sel=document.createElement('select');
    state.fields.forEach(function(f){ sel.appendChild(new Option((f.required?'* ':'')+f.name+' ('+typeShort(f.type)+')', f.name)); });
    selCtrl.appendChild(sel); selWrap.appendChild(selCtrl); body.appendChild(selWrap);
    var formWrap=document.createElement('div'); formWrap.className='settings-form';
    var currentField=null;
    function buildForm(f){
      formWrap.innerHTML='';
      if(!f){ formWrap.innerHTML='<div class="placeholder" style="height:auto;padding:20px 0"><div>请选择字段</div></div>'; return; }
      currentField=f;
      var minInp, maxInp, minLen, maxLen;
      function addRow(label, ctrl){ var fr=document.createElement('div'); fr.className='form-field'; var lab=document.createElement('label'); lab.textContent=label; fr.appendChild(lab); var c=document.createElement('div'); c.className='ctrl'; c.appendChild(ctrl); fr.appendChild(c); formWrap.appendChild(fr); }
      var nameInp=document.createElement('input'); nameInp.value=f.name; nameInp.disabled=true; addRow('字段名', nameInp);
      var typeSel=document.createElement('select'); ['Text','LongText','Number','Integer','Date','DateTime','Select','Quarter','Currency','Percentage','Email','Phone','Url','Checkbox','Image'].forEach(function(t){ typeSel.appendChild(new Option(typeShort(t), t)); }); typeSel.value=f.type; addRow('类型', typeSel);
      var fmtInp=document.createElement('input'); fmtInp.value=f.format||''; fmtInp.placeholder='如 money / percent / int / phone / idcard / datetime'; addRow('显示格式', fmtInp);
      var reqWrap=document.createElement('div'); reqWrap.style.cssText='display:flex;align-items:center;gap:6px';
      var reqCb=document.createElement('input'); reqCb.type='checkbox'; reqCb.checked=!!f.required; reqCb.style.cssText='width:16px;height:16px;margin:0'; reqWrap.appendChild(reqCb);
      var reqLab=document.createElement('span'); reqLab.textContent='必填'; reqLab.style.fontSize='12px'; reqWrap.appendChild(reqLab);
      addRow('', reqWrap);
      if(isNum({type:typeSel.value})){
        minInp=document.createElement('input'); minInp.type='number'; minInp.value=f.minValue!=null?f.minValue:''; addRow('最小值', minInp);
        maxInp=document.createElement('input'); maxInp.type='number'; maxInp.value=f.maxValue!=null?f.maxValue:''; addRow('最大值', maxInp);
      }
      if(typeSel.value==='Text'||typeSel.value==='LongText'||typeSel.value==='Email'||typeSel.value==='Phone'||typeSel.value==='Url'){
        minLen=document.createElement('input'); minLen.type='number'; minLen.value=f.minLength!=null?f.minLength:''; addRow('最小长度', minLen);
        maxLen=document.createElement('input'); maxLen.type='number'; maxLen.value=f.maxLength!=null?f.maxLength:''; addRow('最大长度', maxLen);
      }
      var regexInp=document.createElement('input'); regexInp.value=f.regex||''; addRow('正则验证', regexInp);
      var errInp=document.createElement('input'); errInp.value=f.errorMessage||''; addRow('错误提示', errInp);
      if(typeSel.value==='Select'||typeSel.value==='Quarter'){
        var optsArea=document.createElement('textarea'); optsArea.value=(f.options||[]).join('\n'); addRow('选项（每行一个）', optsArea);
      }
      typeSel.addEventListener('change', function(){ buildForm(f); });
      var apply=document.createElement('button'); apply.className='btn primary'; apply.textContent='保存该字段';
      apply.addEventListener('click', function(){
        f.type=typeSel.value; f.format=fmtInp.value.trim();
        f.required=reqCb.checked;
        if(isNum({type:typeSel.value})){ f.minValue=minInp&&minInp.value!==''?Number(minInp.value):null; f.maxValue=maxInp&&maxInp.value!==''?Number(maxInp.value):null; }
        else { f.minValue=null; f.maxValue=null; }
        if(typeSel.value==='Text'||typeSel.value==='LongText'||typeSel.value==='Email'||typeSel.value==='Phone'||typeSel.value==='Url'){ f.minLength=minLen&&minLen.value!==''?parseInt(minLen.value):null; f.maxLength=maxLen&&maxLen.value!==''?parseInt(maxLen.value):null; }
        else { f.minLength=null; f.maxLength=null; }
        f.regex=regexInp.value.trim(); f.errorMessage=errInp.value.trim();
        if(typeSel.value==='Select'||typeSel.value==='Quarter'){ var opts=formWrap.querySelector('textarea'); f.options=opts?opts.value.split(/\r?\n/).map(function(s){return s.trim();}).filter(Boolean):[]; if(typeSel.value==='Quarter'&&f.options.length===0) f.options=['第一季度','第二季度','第三季度','第四季度']; }
        syncFieldToConfig(f); renderCombo(); renderContent(); toast('字段「'+f.name+'」已保存');
      });
      var rowAct=document.createElement('div'); rowAct.style.cssText='display:flex;justify-content:flex-end;margin-top:12px'; rowAct.appendChild(apply); formWrap.appendChild(rowAct);
    }
    function syncFieldToConfig(f){
      state.config.fieldOverrides=state.config.fieldOverrides||[];
      var ov=state.config.fieldOverrides.find(function(x){return x.name===f.name;});
      if(!ov){ ov={name:f.name}; state.config.fieldOverrides.push(ov); }
      ov.type=f.type; ov.format=f.format||''; ov.options=(f.options||[]).slice();
      ov.required=!!f.required; ov.minValue=f.minValue; ov.maxValue=f.maxValue; ov.minLength=f.minLength; ov.maxLength=f.maxLength; ov.regexPattern=f.regex||''; ov.errorMessage=f.errorMessage||''; ov.userDefined=true;
    }
    function renderCombo(){
      var v=sel.value;
      sel.innerHTML='';
      state.fields.forEach(function(f){ sel.appendChild(new Option((f.required?'* ':'')+f.name+' ('+typeShort(f.type)+')', f.name)); });
      sel.value=v || (currentField?currentField.name:(state.fields[0]?state.fields[0].name:''));
    }
    sel.addEventListener('change', function(){ currentField=state.fields.find(function(x){return x.name===sel.value;}); buildForm(currentField); });
    body.appendChild(formWrap);
    sp.appendChild(body);
    var foot=document.createElement('div'); foot.className='dp-foot';
    var close=document.createElement('button'); close.className='btn primary'; close.textContent='完成';
    close.addEventListener('click', function(){ saveConfig(); closeSettingsPanel(); });
    foot.appendChild(close); sp.appendChild(foot);
    currentField=state.fields[0]; if(currentField){ sel.value=currentField.name; buildForm(currentField); }
  }

  function makeApply(fn){
    var row=document.createElement('div'); row.style.cssText='margin-top:10px';
    var btn=document.createElement('button'); btn.className='btn primary'; btn.textContent='应用';
    btn.addEventListener('click', function(){ fn(); saveConfig(); renderContent(); var ov=document.querySelector('.overlay'); if(ov) ov.remove(); toast('设置已保存'); });
    row.appendChild(btn); return row;
  }

  // ───────── 增 / 删 / 选 ─────────
  function addRow(){
    api.call('addRow',{values:{}}).then(function(res){
      if(res && res.rowIndex>0){
        var newIdx=res.rowIndex;
        openTable(state.sheet, state.table, function(){
          var type=activeViewType();
          state.selectedRow=newIdx;
          if(type==='Form'){
            state.formIndex=getVisibleRows().findIndex(function(x){return x.rowIndex===newIdx;});
            if(state.formIndex<0) state.formIndex=0;
            state._focusForm=true;
            renderContent();
            toast('已新增一行');
          } else if(type==='Table'){
            renderContent();
            showDetailPanel(newIdx);
            toast('已新增一行');
          } else {
            showDetailPanel(newIdx);
            toast('已新增一行');
          }
        });
      } else toast('新增失败');
    }).catch(function(err){ toast('新增失败：'+err.message); });
  }
  function deleteRow(rowIndex){
    if(!confirm('确定删除第 '+rowIndex+' 行？')) return;
    api.call('deleteRow',{rowIndex:rowIndex}).then(function(){ openTable(state.sheet, state.table); toast('已删除'); }).catch(function(err){ toast('删除失败：'+err.message); });
  }
  function insertRow(rowIndex, after){
    api.call('insertRow',{rowIndex:rowIndex, after:!!after}).then(function(res){
      if(res && res.rowIndex>0){
        openTable(state.sheet, state.table, function(){
          state.selectedRow=res.rowIndex;
          renderContent();
          showDetailPanel(res.rowIndex);
          toast('已插入一行');
        });
      } else toast('插入失败');
    }).catch(function(err){ toast('插入失败：'+err.message); });
  }
  function showRowContextMenu(e, rowIndex){
    e.preventDefault(); hideContextMenu();
    var menu=$('ctxMenu'); menu.innerHTML='';
    var items=[
      {label:'在上方插入行', action:function(){ insertRow(rowIndex, false); }},
      {label:'在下方插入行', action:function(){ insertRow(rowIndex, true); }},
      {sep:true},
      {label:'查看详情', action:function(){ showDetailPanel(rowIndex); }},
      {label:'删除记录', action:function(){ deleteRow(rowIndex); }}
    ];
    items.forEach(function(it){
      if(it.sep){ var s=document.createElement('div'); s.className='ctx-sep'; menu.appendChild(s); }
      else { var d=document.createElement('div'); d.className='ctx-item'; d.textContent=it.label; d.addEventListener('click', function(){ hideContextMenu(); it.action(); }); menu.appendChild(d); }
    });
    menu.style.left=e.clientX+'px'; menu.style.top=e.clientY+'px';
    menu.classList.remove('hidden');
    setTimeout(function(){ document.addEventListener('click', hideContextMenu, {once:true}); }, 0);
  }
  function hideContextMenu(){ $('ctxMenu').classList.add('hidden'); }
  function selectRow(rowIndex, tr){
    var changed=state.selectedRow!==rowIndex;
    state.selectedRow=rowIndex;
    var rows=$('content').querySelectorAll('tbody tr'); rows.forEach(function(x){ x.classList.remove('selected'); });
    if(tr) tr.classList.add('selected');
    api.call('selectRow',{rowIndex:rowIndex}).catch(function(){});
    // 右侧详情面板若已打开，且焦点行变化，则跟随刷新
    var dp=$('detailPanel'); if(changed && dp && dp.classList.contains('open')){ showDetailPanel(rowIndex); }
  }

  // ───────── 字段显隐面板 ─────────
  function toggleFieldPanel(btn){
    var p=$('fieldPanel');
    if(!p.classList.contains('hidden')){ p.classList.add('hidden'); return; }
    p.innerHTML='';
    state.fields.forEach(function(f){
      var lab=document.createElement('label');
      var cb=document.createElement('input'); cb.type='checkbox'; cb.checked=state.visibleFields.indexOf(f.name)>=0;
      cb.addEventListener('change', function(){
        if(this.checked){ if(state.visibleFields.indexOf(f.name)<0) state.visibleFields.push(f.name); }
        else { state.visibleFields=state.visibleFields.filter(function(x){return x!==f.name;}); }
        if(state.visibleFields.length===0) state.visibleFields.push(f.name);
        var av=activeView(); if(av) av.visibleFields=state.visibleFields.slice();
        saveConfig(); renderContent();
      });
      lab.appendChild(cb); var span=document.createElement('span'); span.textContent=fieldLabel(f); lab.appendChild(span);
      p.appendChild(lab);
    });
    var rect=btn.getBoundingClientRect();
    p.style.left=rect.left+'px';
    p.style.top=(rect.bottom+4)+'px';
    p.classList.remove('hidden');
    p.onmouseleave=function(){ p.classList.add('hidden'); };
    setTimeout(function(){
      function docClick(e){ if(!p.contains(e.target) && e.target!==btn){ p.classList.add('hidden'); document.removeEventListener('click', docClick); } }
      document.addEventListener('click', docClick);
    }, 10);
  }

  // ───────── 配置保存 ─────────
  function saveConfig(){
    if(!state.config) return;
    var cfg=JSON.parse(JSON.stringify(state.config));
    // 同步当前活动视图的可见字段与排序
    var av=cfg.views.find(function(v){return v.viewId===state.activeViewId;}) || cfg.views[0];
    if(av){ av.visibleFields=state.visibleFields.slice(); av.sort=[]; if(state.sortField&&state.sortOrder) av.sort.push({field:state.sortField,order:state.sortOrder}); }
    // 同步各视图的专属配置（从 state.views 回写）
    state.views.forEach(function(sv){ var cv=cfg.views.find(function(x){return x.viewId===sv.viewId;}); if(cv){ cv.groupBy=sv.groupBy; cv.cardMeta=sv.cardMeta; cv.calendarConfig=sv.calendarConfig; cv.ganttConfig=sv.ganttConfig; cv.chartConfig=sv.chartConfig; cv.dashboardConfig=sv.dashboardConfig; cv.visibleFields=sv.visibleFields; } });
    api.call('saveConfig',{json:JSON.stringify(cfg)}).catch(function(err){ toast('配置保存失败：'+err.message); });
  }

  // ───────── 关于 ─────────
  function showAbout(){
    var a=state.appInfo||{};
    var ver=a.version||BUILD.version||'未知';
    var t=a.dllModified||BUILD.time||'未知';
    var dll=BUILD.dll||'未知';
    alert('多维表分析 '+ver+'\n编译/部署时间：'+t+'\nDLL 路径：'+dll+'\n宿主：'+(a.host||'未知')+'\nExcel 版本：'+(a.excelVersion||'未知'));
  }

  // ───────── Ribbon 回调 ─────────
  window.__mtReload=function(){ if(state.sheet&&state.table) openTable(state.sheet, state.table); };
  window.__mtSave=function(){ saveConfig(); toast('配置已保存'); };

  if(document.readyState==='loading') document.addEventListener('DOMContentLoaded', function(){ waitReady(start); }); else waitReady(start);
})();
