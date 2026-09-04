# 多维表分析软件 · 拖拽交互逻辑与参数说明

> 本文档对应 Phase C 需求 C3，说明日历视图（日/周/月）与甘特视图中的事件/条形拖拽实现，包括坐标换算、对齐规则、实时保存流程与关键参数。
> 相关代码位于 `MultiTableAddin/src/MultiTableAddin/Html/index.html`。

---

## 1. 公共基础设施

### 1.1 日期格式化
```javascript
function fmtDate(d){ return d? (d.getFullYear()+'-'+pad(d.getMonth()+1)+'-'+pad(d.getDate())) : ''; }
function fmtDateTimeLocal(d, h, m){ return d.getFullYear()+'-'+pad(d.getMonth()+1)+'-'+pad(d.getDate())+'T'+pad(h)+':'+pad(m)+':00'; }
function pad(n){ return n<10?'0'+n:''+n; }
```

- `fmtDate`：生成 `YYYY-MM-DD`（Date 类型字段使用）。
- `fmtDateTimeLocal`：生成 `YYYY-MM-DDTHH:mm:00`（DateTime 类型字段使用）。
- 两者均保证不受时区影响，可直接被 Excel/WPS 解析。

### 1.2 保存通道
所有拖拽在 `mouseup` 后通过统一接口写回 Excel：
```javascript
api.call('updateCell', { rowIndex: r.rowIndex, field: fieldName, value: newValue })
```
- 先更新本地 `r.values[field]`，再异步调用 `updateCell`。
- 成功提示“时间已更新”，失败回滚并重新渲染。

---

## 2. 日历视图拖拽（日 / 周视图）

> 函数：`makeCalEventDraggable(ev, body, r, cfg)`
> 适用：日历日视图、周视图（纵向时间网格）。

### 2.1 时间网格参数
```javascript
var PX_PER_MIN = 40 / 60;   // 每小时 40px，即 2/3 px/分钟
var snap = (cc.snapMinutes && cc.snapMinutes > 0) ? cc.snapMinutes : 15; // 默认 15 分钟对齐
```

| 参数 | 默认值 | 说明 |
|------|--------|------|
| `PX_PER_MIN` | `40/60` | 时间轴纵向分辨率，1 分钟对应 0.667px |
| `snap` | `15` | 拖拽吸附粒度（分钟），来自 `calendarConfig.snapMinutes` |
| `timeToTop(d)` | — | `d.getHours()*40 + d.getMinutes()/60*40`：把日期时间换算成事件 `top` |
| `snapMin(m)` | — | `Math.round(m/snap)*snap`，限制 `[0, 1439]` |

### 2.2 三种拖拽模式

每个事件卡片内部生成两个把手：
- `.cal-event-resize-top`：调整**开始时间**。
- `.cal-event-resize`（仅当 `cfg.endDateField` 存在）：调整**结束时间**。
- 卡片主体：整体移动，保持起止时长不变。

#### 2.2.1 整体移动
```javascript
var dy = em.clientY - startY;
var nsMin = snapMin((origStart.getHours()*60 + origStart.getMinutes()) + dy / PX_PER_MIN);
var ns = new Date(dayBase.getTime() + nsMin * 60000);
var ne = new Date(ns.getTime() + duration);
```
- `dayBase`：原开始日期的 0:00。
- 新开始时间 = 原开始时间 + 纵向偏移分钟数（吸附后）。
- 新结束时间 = 新开始时间 + 原时长 `duration`。

#### 2.2.2 调整开始时间（顶部把手）
```javascript
var nsMin = snapMin(startMin + dy / PX_PER_MIN);
if (nsMin > endMin - snap) nsMin = endMin - snap;   // 开始不得晚于结束
var ns = new Date(dayBase.getTime() + nsMin * 60000);
ev.style.height = Math.max(20, endMin - nsMin) * PX_PER_MIN + 'px';
```
- 仅改变开始时间；结束时间保持不变。
- 限制：开始时间不能跨越结束时间，至少保留一个 `snap` 间隔。
- 高度随开始时间移动而动态变化。

#### 2.2.3 调整结束时间（底部把手）
```javascript
var neMin = snapMin(endMin + dy / PX_PER_MIN);
if (neMin < startMin + snap) neMin = startMin + snap; // 结束不得早于开始
var ne = new Date(dayBase.getTime() + neMin * 60000);
ev.style.height = Math.max(20, neMin - startMin) * PX_PER_MIN + 'px';
```
- 仅改变结束时间；开始时间保持不变。
- 限制：结束时间不能早于开始时间。

### 2.3 保存流程
```javascript
function finishDrag(){ ... updateCalEventTime(r, cfg, ns, ne); }
function updateCalEventTime(r, cfg, newStart, newEnd){
  var updates = [{field: cfg.dateField, value: fmtDateTimeLocal(newStart, h, m)}];
  if(cfg.endDateField && newEnd) updates.push({field: cfg.endDateField, value: fmtDateTimeLocal(newEnd, h, m)});
  renderContent();
  updates.forEach(function(u){ api.call('updateCell', {...}); });
}
```
- 松开鼠标后立即 `renderContent()` 刷新视图。
- 每个字段独立调用 `updateCell`。

---

## 3. 日历视图拖拽（月视图）

> 函数：`makeMonthEventDraggable(ev, r, cfg, cell, cells)`
> 适用：日历月视图（横向按日期单元格拖动）。

### 3.1 参数
```javascript
var startX = e.clientX;
var startIndex = cells.indexOf(cell);
var cellW = cell.getBoundingClientRect().width;
var origDate = toDate(r.values[cfg.dateField]);
var origEnd = cfg.endDateField ? toDate(r.values[cfg.endDateField]) : null;
var duration = origEnd && origDate ? origEnd - origDate : 0;
```

| 参数 | 说明 |
|------|------|
| `startX` | 鼠标按下时的屏幕 X 坐标 |
| `startIndex` | 事件所在的单元格索引 |
| `cellW` | 单个日期单元格宽度 |
| `duration` | 原起止时间差（毫秒），移动时保持不变 |

### 3.2 位置换算
```javascript
var dx = em.clientX - startX;
var delta = Math.round(dx / cellW);   // 横向跨越的单元格数
var newIndex = Math.max(0, Math.min(cells.length - 1, startIndex + delta));
var targetDate = cells[newIndex].date;
```
- 以单元格宽度为步长，计算横向平移天数 `delta`。
- `targetDate` 为新开始日期，时间部分继承原开始时间的小时/分钟。

### 3.3 起止日期更新
```javascript
var h = origDate ? origDate.getHours() : 0, m = origDate ? origDate.getMinutes() : 0;
var newStartStr = (sf && sf.type === 'DateTime') ? fmtDateTimeLocal(targetDate, h, m) : fmtDate(targetDate);
var updates = [{field: cfg.dateField, value: newStartStr}];
if(cfg.endDateField && origEnd){
  var newEnd = new Date(targetDate.getTime() + duration);
  updates.push({field: cfg.endDateField, value: ...});
}
```
- 开始日期替换为 `targetDate`，时间分量保留。
- 结束日期 = `targetDate + duration`。
- 字段类型为 `DateTime` 时输出 `fmtDateTimeLocal`，否则输出 `fmtDate`。

---

## 4. 甘特视图拖拽

> 函数：`makeGanttBarDraggable(bar, r, cfg, dim)`
> 适用：甘特图日/周/月视图。

### 4.1 视图缩放级别
```javascript
var ganttState = { dim: 'month', anchorDate: null, min: null, max: null };
function ganttRange(dim, anchor, dataMin, dataMax){
  var span = {day: 3, week: 21, month: 70}[dim] || 70;
  var center = anchor ? new Date(anchor.getTime())
    : (dataMin && dataMax ? new Date((dataMin.getTime() + dataMax.getTime()) / 2) : new Date());
  var min = new Date(center.getTime() - span / 2 * 86400000);
  var max = new Date(center.getTime() + span / 2 * 86400000);
  if(dataMin && dataMax){
    var inView = !(max < dataMin || min > dataMax);
    if(!inView){    // 数据不在当前窗口内时，以数据中心平移窗口
      var dc = new Date((dataMin.getTime() + dataMax.getTime()) / 2);
      min = new Date(dc.getTime() - span / 2 * 86400000);
      max = new Date(dc.getTime() + span / 2 * 86400000);
    }
  }
  return {min, max};
}
```

| 维度 | 可视天数 `span` | 主刻度 | 说明 |
|------|----------------|--------|------|
| `day` | 3 天 | 6 小时 | 精细调整小时级任务 |
| `week` | 21 天 | 1 天 | 按天调整 |
| `month` | 70 天 | 7 天 | 按周概览 |

- 窗口中心由 `ganttState.anchorDate`（用户点击“今天”或切换维度时记录）决定。
- 若当前窗口内没有任何数据，自动把窗口平移到数据所在区间，避免空白。

### 4.2 条形位置渲染
```javascript
var left = (s - min) / 86400000 / totalDays * 100;
var width = Math.max(1.5, (e - s) / 86400000 / totalDays * 100);
bar.style.left = left + '%';
bar.style.width = width + '%';
```
- 所有位置使用**百分比**，使甘特条随容器自适应。
- `totalDays = (max - min) / 86400000`。

### 4.3 拖拽模式判定
```javascript
var RESIZE_EDGE = 8;  // px
var mode = (e.clientX - rect.left > rect.width - RESIZE_EDGE) ? 'resize' : 'move';
```
- 鼠标落在条形右侧 8px 范围内 → **调整结束日期**。
- 其他区域 → **整体平移**。

### 4.4 实时反馈
```javascript
var trackRect = track.getBoundingClientRect();
var dx = ev.clientX - startX;
var dPct = (dx / trackRect.width) * 100;
if(mode === 'move'){
  bar.style.left = Math.max(0, Math.min(100 - startWidth, startLeft + dPct)) + '%';
} else {
  bar.style.width = Math.max(0.3, startWidth + dPct) + '%';
}
```
- 把鼠标像素位移换算为百分比位移 `dPct`。
- 整体移动时限制 `left ∈ [0, 100 - width]`。
- 调整宽度时最小保留 `0.3%`。

### 4.5 日期反算与对齐
```javascript
var totalMs = ganttState.max - ganttState.min;
var newLeft = parseFloat(bar.style.left) || 0;
var newWidth = parseFloat(bar.style.width) || 0;
var newStart = new Date(ganttState.min.getTime() + (newLeft / 100) * totalMs);
var newEnd   = new Date(ganttState.min.getTime() + ((newLeft + newWidth) / 100) * totalMs);

function alignDate(d){
  var x = new Date(d);
  if(dim === 'day'){ var m = Math.round(x.getMinutes() / 15) * 15; x.setMinutes(m); x.setSeconds(0); }
  else { x.setHours(0, 0, 0, 0); }
  return x;
}
newStart = alignDate(newStart);
newEnd   = alignDate(newEnd);
if(newEnd <= newStart) newEnd = new Date(newStart.getTime() + 86400000);
```

| 维度 | 对齐规则 |
|------|----------|
| `day` | 按 15 分钟吸附（`Math.round(minutes/15)*15`） |
| `week` / `month` | 对齐到当天 0:00 |

- 由百分比反推时间：
  - `newStart = min + left% × totalMs`
  - `newEnd = min + (left% + width%) × totalMs`
- 保证结束不早于或等于开始，最小跨度 1 天。

### 4.6 保存
```javascript
updates.forEach(function(u){
  api.call('updateCell', {rowIndex: r.rowIndex, field: u.field, value: u.value}).then(function(){
    r.values[u.field] = u.value;
    pending--;
    if(pending === 0){ toast('时间已更新'); renderGantt(); }
  }).catch(function(err){ toast('保存失败：' + err.message); renderGantt(); });
});
```
- 每个字段独立保存；全部成功后重新渲染甘特图。

---

## 5. 关键参数速查

| 参数 | 位置 | 默认值 | 含义 |
|------|------|--------|------|
| `PX_PER_MIN` | `makeCalEventDraggable` | `40/60` | 日/周视图 1 分钟对应的像素高度 |
| `snap` | `makeCalEventDraggable` | `15` | 日/周视图时间吸附粒度（分钟） |
| `RESIZE_EDGE` | `makeGanttBarDraggable` | `8` | 甘特条右侧判定为 resize 的边缘宽度（px） |
| `span.day` | `ganttRange` | `3` | 甘特日视图可视天数 |
| `span.week` | `ganttRange` | `21` | 甘特周视图可视天数 |
| `span.month` | `ganttRange` | `70` | 甘特月视图可视天数 |
| 日视图时间对齐 | `alignDate` | `15 min` | 甘特日视图下按 15 分钟取整 |
| 周/月视图日期对齐 | `alignDate` | `00:00` | 甘特周/月视图下对齐到当天 0:00 |

---

## 6. 交互时序图

```
用户 mousedown
    │
    ▼
判定拖拽模式（move / resize-start / resize-end / month-cell-shift）
    │
    ▼
mousemove：实时换算像素位移 → 新时间/新百分比 → 更新 DOM（top/left/width）
    │
    ▼
mouseup：
  ├─ 本地生成新日期值
  ├─ 调用 renderContent() / renderGantt() 刷新视图
  └─ 异步 api.call('updateCell', ...) 写回 Excel/WPS
    │
    ▼
成功：toast('时间已更新')
失败：toast('保存失败：...') 并重新渲染
```

---

## 7. 注意事项

1. **所有日期输出均为本地时间字符串**，不使用 `toISOString()`，避免 Excel/WPS 解析出前一天。
2. **日/周视图为纵向拖动，月视图为横向拖动**，由各自 `mousedown` 的初始坐标方向决定，不需要额外的模式切换。
3. **甘特条使用百分比定位**，保证容器宽度变化时位置不变；最终日期由当前 `ganttState.min/max` 反算，因此切换日/周/月后数值稳定。
4. **吸附规则仅影响最终保存值**，不影响拖动过程中视觉反馈的流畅性。
