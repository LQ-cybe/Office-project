# -*- coding: utf-8 -*-
import re, io

p = r"D:/编程开发/开源项目/多维表分析软件/MultiTableAddin/src/MultiTableAddin/Html/index.html"
with open(p, 'rb') as f:
    raw = f.read()
data = raw.decode('utf-8')
# 统一换行符便于后续按 \n 匹配（HTML 资源对换行不敏感，构建无影响）
data = data.replace('\r\n', '\n')
orig = data

gear_svg = '<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 1 1-2.83 2.83l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-4 0v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 1 1-2.83-2.83l.06-.06a1.65 1.65 0 0 0 .33-1.82 1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1 0-4h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 1 1 2.83-2.83l.06.06a1.65 1.65 0 0 0 1.82.33H9a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 4 0v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 1 1 2.83 2.83l-.06.06a1.65 1.65 0 0 0-.33 1.82V9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 0 4h-.09a1.65 1.65 0 0 0-1.51 1z"/></svg>'

# 1) 画册 SVG 条目 -> 齿轮 SVG 条目（画册改用看板图标，这里新增 gear 供“设置”按钮使用）
pat = r"    gallery:'<svg[^>]*>.*?</svg>',"
assert re.search(pat, data), "gallery SVG entry not found"
data = re.sub(pat, lambda m: "    gear:" + gear_svg + "',", data, count=1)

# 2) VIEW_ICON 画册 -> kanban（画册改用看板图标）
assert "Gallery:'gallery'" in data, "VIEW_ICON Gallery not found"
data = data.replace("Gallery:'gallery'", "Gallery:'kanban'")

# 3) FIELD_ICON 画册 -> ▥（与看板一致）
assert "Gallery:'▧'" in data, "FIELD_ICON Gallery not found"
data = data.replace("Gallery:'▧'", "Gallery:'▥'")

# 4) 设置按钮 ⚙ -> 齿轮 SVG（⚙ 是字符，不随 currentColor 变色）
old_set = "set.innerHTML='<span class=\"ic\">⚙</span>设置';"
assert old_set in data, "set button not found"
data = data.replace(old_set, "set.innerHTML='<span class=\"ic\">'+icon('gear')+'</span>设置';")

# 5) 窗口左上角品牌标识（之前 .sidebar-brand .dot 规则无对应 DOM 而无效）
old_sb = "  function renderSidebar(){\n    renderSidebarTop();"
new_sb = """  function renderSidebar(){
    // #502 窗口左上角品牌标识：之前 .sidebar-brand .dot 规则无对应 DOM 而无效，这里补一个真实深蓝品牌图标
    var sb=$('sidebar');
    if(sb && !sb.querySelector('.sidebar-brand')){
      var brand=document.createElement('div'); brand.className='sidebar-brand';
      brand.innerHTML='<span class="dot"></span><span>多维表分析</span>';
      sb.insertBefore(brand, sb.firstChild);
    }
    renderSidebarTop();"""
assert old_sb in data, "renderSidebar start not found"
data = data.replace(old_sb, new_sb)

# 6) 拖到末尾空白处归入最后一列（末尾分组）
old_drag = """          for(var i=0;i<els.length;i++){ var c=els[i].closest('.kanban-col'); if(c){ targetCol=c; break; } }
        }
        state._kanbanDragCard=null; state._kanbanDragRow=null; state._kanbanDragGroupBy=null;"""
new_drag = """          for(var i=0;i<els.length;i++){ var c=els[i].closest('.kanban-col'); if(c){ targetCol=c; break; } }
        }
        // #501 拖到末尾空白处（落入分组列容器但未命中具体列）时，归入最后一列（末尾分组）
        if(moved && !targetCol){
          var colsWrap=card.closest('.kanban-cols');
          if(colsWrap){
            var rb=colsWrap.getBoundingClientRect();
            if(em.clientX>=rb.left && em.clientX<=rb.right && em.clientY>=rb.top && em.clientY<=rb.bottom){
              var allCols=colsWrap.querySelectorAll('.kanban-col');
              if(allCols.length) targetCol=allCols[allCols.length-1];
            }
          }
        }
        state._kanbanDragCard=null; state._kanbanDragRow=null; state._kanbanDragGroupBy=null;"""
assert old_drag in data, "drag targetCol block not found"
data = data.replace(old_drag, new_drag)

# 7) renderContent：遗留看板视图按画册渲染
assert "else if(type==='Kanban') renderKanban();" in data, "renderContent Kanban not found"
data = data.replace("else if(type==='Kanban') renderKanban();",
                     "else if(type==='Kanban') renderGallery(); // #503 看板视图已删除，遗留看板视图按画册渲染（画册已吸收看板分组拖拽）")

# 8) 删除 renderKanban 函数（含上方注释）
m = re.search(r"//\s*[─\u2500\-]+\s*看板视图\s*[─\u2500\-]+\s*\n\s*function renderKanban\(\)\{", data)
assert m, "renderKanban comment+signature not found"
start = m.start()
end_marker = "    wrap.appendChild(cols); $('content').innerHTML=''; $('content').appendChild(wrap);\n  }"
assert end_marker in data, "renderKanban end not found"
end = data.index(end_marker) + len(end_marker)
data = data[:start] + "  // #503 看板视图已删除：遗留看板视图由 renderContent 路由到 renderGallery 渲染（画册已吸收看板分组拖拽），renderKanban 不再需要。\n" + data[end:]

# 9) scrollSelectedIntoView：遗留看板视图按画册渲染
assert "else if(type==='Kanban') sel='.kanban-card[data-row]';" in data, "scrollSelected Kanban not found"
data = data.replace("else if(type==='Kanban') sel='.kanban-card[data-row]';",
                    "else if(type==='Kanban') sel='.gallery-card[data-row]'; // #503 遗留看板视图按画册渲染")

with open(p, 'wb') as f:
    f.write(data.encode('utf-8'))
print("index.html changed:", orig != data)
