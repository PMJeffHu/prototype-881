/**
 * 逾期账单抵减过程演示 Excel 生成脚本
 * 数据源：index.html 内嵌的 PROC_* 数据数组与 runOffsetEngine 引擎（提取重放，保证与页面 100% 一致）
 * 运行：node gen-offset-xlsx.js
 */
var fs = require('fs');
var path = require('path');
var vm = require('vm');
var ExcelJS = require('exceljs');

var html = fs.readFileSync(path.join(__dirname, 'index.html'), 'utf8');
var script = html.match(/<script>([\s\S]*)<\/script>/)[1];

/* 提取四类单据数据 + 单号重写 + 抵减引擎 + PROC_RESULT（截止到页面层代码标记） */
var start = script.indexOf('var PROC_POS');
var end = script.indexOf('/* ===== 页面层代码');
if (start < 0 || end < 0) { console.error('提取失败：未找到 PROC 数据或引擎边界'); process.exit(1); }

var sandbox = { module: { exports: {} } };
vm.createContext(sandbox);
vm.runInContext(script.slice(start, end)
  + '\nmodule.exports = { documents: PROC_RESULT.documents, events: PROC_RESULT.events };', sandbox);
var documents = sandbox.module.exports.documents;
var events = sandbox.module.exports.events;
console.log('提取成功：单据 ' + documents.length + ' 条 / 抵减事件 ' + events.length + ' 条');

/* ---------- 样式常量（Element UI 风格） ---------- */
var HEAD_BG = 'FF409EFF';   /* 表头主色蓝 */
var SUM_BG = 'FFF5F7FA';    /* 合计行淡灰 */
var BLUE = 'FF409EFF';      /* 非零金额蓝字 */
var GRAY = 'FF909399';      /* 零值灰字 */
var MONEY = '#,##0.00;(#,##0.00);"-"';
var RATE = '0.0000';

function byType(t) { return documents.filter(function (d) { return d.type === t; }); }

/* 通用建表：cols=[{h,w,fmt,center}]，rows=二维数组，sumCols=需要 SUM 合计的列号(1起) */
function buildSheet(wb, name, cols, rows, sumLabel, sumCols) {
  var ws = wb.addWorksheet(name, { views: [{ state: 'frozen', ySplit: 1 }] });
  cols.forEach(function (c, i) { ws.getColumn(i + 1).width = c.w || 14; });

  var hr = ws.getRow(1);
  hr.height = 24;
  cols.forEach(function (c, i) {
    var cell = hr.getCell(i + 1);
    cell.value = c.h;
    cell.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11, name: 'Arial' };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: HEAD_BG } };
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
  });

  rows.forEach(function (r, idx) {
    var xr = ws.getRow(idx + 2);
    r.forEach(function (v, i) {
      var cell = xr.getCell(i + 1);
      cell.value = v;
      cell.font = { size: 11, name: 'Arial' };
      if (cols[i].fmt) cell.numFmt = cols[i].fmt;
      if (typeof v === 'number' && cols[i].fmt === MONEY) {
        cell.font = { size: 11, name: 'Arial', color: { argb: v === 0 ? GRAY : BLUE } };
      }
      if (cols[i].center) cell.alignment = { horizontal: 'center' };
    });
  });

  var tr = ws.getRow(rows.length + 2);
  tr.getCell(1).value = sumLabel;
  sumCols.forEach(function (i) {
    var letter = ws.getColumn(i).letter;
    var cell = tr.getCell(i);
    cell.value = { formula: 'SUM(' + letter + '2:' + letter + (rows.length + 1) + ')' };
    cell.numFmt = cols[i - 1].fmt || MONEY;
  });
  for (var i = 1; i <= cols.length; i++) {
    var c = tr.getCell(i);
    c.font = { bold: true, size: 11, name: 'Arial' };
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: SUM_BG } };
    c.border = { top: { style: 'thin', color: { argb: 'FFC0C4CC' } } };
  }
  return ws;
}

var wb = new ExcelJS.Workbook();
wb.creator = 'FPS 应收系统原型';

/* ---------- Sheet 1：逾期正数账单 ---------- */
var posRows = byType('正数账单').map(function (d) {
  return [d.no, d.days, d.cur, d.amtO, d.rate, d.cny, d.offset, d.remain, d.status];
});
buildSheet(wb, '逾期正数账单', [
  { h: '账单号', w: 18 }, { h: '逾期天数', w: 12, center: true }, { h: '账单币种', w: 12, center: true },
  { h: '逾期金额(原币)', w: 16, fmt: MONEY }, { h: '记账汇率', w: 12, fmt: RATE },
  { h: '抵减前金额(CNY)', w: 14, fmt: MONEY }, { h: '被抵减金额(CNY)', w: 16, fmt: MONEY },
  { h: '抵减后余额(CNY)', w: 16, fmt: MONEY }, { h: '抵减状态', w: 12, center: true }
], posRows, '合计', [4, 6, 7, 8]);

/* ---------- Sheet 2：逾期负数账单 ---------- */
var negRows = byType('负数账单').map(function (d) {
  return [d.no, d.days, d.cur, d.amtO, d.rate, d.cny, d.offset, d.remain, d.status];
});
buildSheet(wb, '逾期负数账单', [
  { h: '账单号', w: 18 }, { h: '逾期天数', w: 12, center: true }, { h: '账单币种', w: 12, center: true },
  { h: '逾期金额(原币)', w: 16, fmt: MONEY }, { h: '记账汇率', w: 12, fmt: RATE },
  { h: '抵减前金额(CNY)', w: 14, fmt: MONEY }, { h: '被抵减金额(CNY)', w: 16, fmt: MONEY },
  { h: '抵减后余额(CNY)', w: 16, fmt: MONEY }, { h: '抵减状态', w: 12, center: true }
], negRows, '合计', [4, 6, 7, 8]);

/* ---------- Sheet 3：充值单 ---------- */
var rcRows = byType('充值单').map(function (d) {
  return [d.no, d.cur, d.amtO, d.rate, d.cny, d.offset, d.remain, d.status];
});
buildSheet(wb, '充值单', [
  { h: '充值单号', w: 18 }, { h: '充值币种', w: 12, center: true },
  { h: '未核销金额(原币)', w: 18, fmt: MONEY }, { h: '记账汇率', w: 12, fmt: RATE },
  { h: '抵减前金额(CNY)', w: 14, fmt: MONEY }, { h: '被抵减金额(CNY)', w: 16, fmt: MONEY },
  { h: '抵减后余额(CNY)', w: 16, fmt: MONEY }, { h: '抵减状态', w: 12, center: true }
], rcRows, '合计', [3, 5, 6, 7]);

/* ---------- Sheet 4：宽限单 ---------- */
var grRows = byType('宽限单').map(function (d) {
  return [d.no, d.cny, d.offset, d.remain, d.status];
});
buildSheet(wb, '宽限单', [
  { h: '宽限单号', w: 18 }, { h: '宽限金额(CNY)', w: 16, fmt: MONEY },
  { h: '被抵减金额(CNY)', w: 16, fmt: MONEY }, { h: '抵减后余额(CNY)', w: 16, fmt: MONEY },
  { h: '抵减状态', w: 12, center: true }
], grRows, '合计', [2, 3, 4]);

/* ---------- Sheet 5：抵减过程明细 ---------- */
var evRows = events.map(function (e) {
  return [e.step, e.type, e.srcNo, e.targetNo, e.amt, e.srcRemain, e.targetRemain];
});
buildSheet(wb, '抵减过程明细', [
  { h: '抵减顺序', w: 8, center: true }, { h: '抵减类型', w: 12, center: true },
  { h: '抵减单据编号', w: 18 }, { h: '被抵减单据编号', w: 18 },
  { h: '本次抵减金额(CNY)', w: 18, fmt: MONEY }, { h: '抵减单据剩余(CNY)', w: 16, fmt: MONEY },
  { h: '被抵减单据剩余(CNY)', w: 16, fmt: MONEY }
], evRows, '合计', [5]);

/* ---------- Sheet 6：抵减结果汇总（公式引用各明细表合计行） ---------- */
var ws = wb.addWorksheet('抵减结果汇总');
ws.getColumn(1).width = 16; ws.getColumn(2).width = 12;
ws.getColumn(3).width = 16; ws.getColumn(4).width = 18; ws.getColumn(5).width = 18;

var hr = ws.getRow(1);
hr.height = 24;
['单据类型', '单据数量', '抵减前金额(CNY)', '被抵减金额(CNY)', '抵减后余额(CNY)'].forEach(function (h, i) {
  var c = hr.getCell(i + 1);
  c.value = h;
  c.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11, name: 'Arial' };
  c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: HEAD_BG } };
  c.alignment = { horizontal: 'center', vertical: 'middle' };
});

var sumRefs = [
  ['正数账单', posRows.length, "'逾期正数账单'!F14", "'逾期正数账单'!G14", "'逾期正数账单'!H14"],
  ['负数账单', negRows.length, "'逾期负数账单'!F6", "'逾期负数账单'!G6", "'逾期负数账单'!H6"],
  ['充值单',   rcRows.length, "'充值单'!E6", "'充值单'!F6", "'充值单'!G6"],
  ['宽限单',   grRows.length, "'宽限单'!B4", "'宽限单'!C4", "'宽限单'!D4"]
];
sumRefs.forEach(function (r, idx) {
  var xr = ws.getRow(idx + 2);
  r.forEach(function (v, i) {
    var c = xr.getCell(i + 1);
    if (i >= 2) c.value = { formula: '=' + v };
    else c.value = v;
    c.font = { size: 11, name: 'Arial' };
    if (i >= 2) c.numFmt = MONEY;
    if (i <= 1) c.alignment = { horizontal: i === 0 ? 'left' : 'center' };
  });
});
var tr = ws.getRow(sumRefs.length + 2);
tr.getCell(1).value = '合计';
['B', 'C', 'D', 'E'].forEach(function (L) {
  var c = tr.getCell(L);
  c.value = { formula: '=SUM(' + L + '2:' + L + (sumRefs.length + 1) + ')' };
  c.numFmt = L === 'B' ? '0' : MONEY;
});
for (var i = 1; i <= 5; i++) {
  var c = tr.getCell(i);
  c.font = { bold: true, size: 11, name: 'Arial' };
  c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: SUM_BG } };
  c.border = { top: { style: 'thin', color: { argb: 'FFC0C4CC' } } };
}

/* 补充指标：抵减事件笔数与抵减总金额（引用过程明细表） */
var info = ws.getRow(sumRefs.length + 4);
info.getCell(1).value = '抵减事件笔数';
info.getCell(2).value = { formula: "=COUNT('抵减过程明细'!A2:A" + (evRows.length + 1) + ')' };
var info2 = ws.getRow(sumRefs.length + 5);
info2.getCell(1).value = '抵减总金额(CNY)';
info2.getCell(2).value = { formula: "='抵减过程明细'!E" + (evRows.length + 2) };
info2.getCell(2).numFmt = MONEY;
[info, info2].forEach(function (r) {
  r.getCell(1).font = { bold: true, size: 11, name: 'Arial' };
  r.getCell(2).font = { size: 11, name: 'Arial' };
});

var out = path.join(__dirname, '逾期账单抵减过程演示.xlsx');
wb.xlsx.writeFile(out).then(function () {
  console.log('已生成: ' + out);
});
