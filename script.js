/* =========================================================================
 * 通讯录 Address Book — 主程序
 * -------------------------------------------------------------------------
 * 文件结构：
 *   第 0 部分  小工具（转义、提示、下载）
 *   第 1 部分  Excel 写入：把二维表格装进 .xlsx 文件（zip 打包）
 *   第 2 部分  Excel 读取：解压 .xlsx 并解析里面的 XML
 *   第 3 部分  数据层：联系人的增删改查 + localStorage 存档
 *   第 4 部分  界面渲染：联系人卡片、收藏置顶、搜索
 *   第 5 部分  表单：新建 / 编辑联系人
 *   第 6 部分  导入导出：Excel ↔ 通讯录
 *   第 7 部分  启动
 *
 * 说明：本项目不依赖任何第三方库，Excel 读写全部手写在下面，
 *      所以双击 index.html 就能用，断网也能用。
 * ========================================================================= */

/* =========================================================================
 * 第 0 部分  小工具
 * ========================================================================= */

var $ = function (sel, root) { return (root || document).querySelector(sel); };
var $$ = function (sel, root) { return Array.prototype.slice.call((root || document).querySelectorAll(sel)); };

var enc = new TextEncoder();
var dec = new TextDecoder('utf-8');

function xmlEsc(s) {
  return String(s).replace(/[&<>"']/g, function (ch) {
    return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&apos;' }[ch];
  });
}

function attrEsc(s) { return xmlEsc(s); }

function uid() {
  return 'c' + Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
}

function today() {
  var d = new Date();
  var p = function (n) { return n < 10 ? '0' + n : '' + n; };
  return d.getFullYear() + p(d.getMonth() + 1) + p(d.getDate());
}

var toastTimer = null;
function toast(msg, isErr) {
  var el = $('#toast');
  if (!el) return;
  el.textContent = msg;
  el.className = 'toast show' + (isErr ? ' is-err' : '');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(function () { el.className = 'toast'; }, 2600);
}

function download(bytes, filename, mime) {
  var blob = new Blob([bytes], { type: mime });
  var url = URL.createObjectURL(blob);
  var a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(function () { URL.revokeObjectURL(url); }, 1500);
}

/* CRC32：zip 里每个文件都要带校验码，浏览器没有现成 API，只能自己算 */
var CRC_TABLE = (function () {
  var t = new Int32Array(256);
  for (var n = 0; n < 256; n++) {
    var c = n;
    for (var k = 0; k < 8; k++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
    t[n] = c;
  }
  return t;
})();

function crc32(buf) {
  var c = -1;
  for (var i = 0; i < buf.length; i++) c = CRC_TABLE[(c ^ buf[i]) & 0xFF] ^ (c >>> 8);
  return (c ^ -1) >>> 0;
}

/* =========================================================================
 * 第 1 部分  Excel 写入
 * .xlsx 本质是一个 zip 包，里面是几个 XML 文件。
 * 这里用「不压缩」的方式打包（zip 的 store 模式），Excel 完全认。
 * ========================================================================= */

/* 数字序号 -> 列名：0->A, 1->B, ..., 25->Z, 26->AA */
function colName(index) {
  var s = '';
  var n = index + 1;
  while (n > 0) {
    var m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/* 一个单元格显示多宽：中文算 2 个字符宽度 */
function visualWidth(v) {
  var s = String(v);
  var w = 0;
  for (var i = 0; i < s.length; i++) w += s.charCodeAt(i) > 255 ? 2 : 1;
  return w;
}

function buildSheetXml(rows) {
  var colCount = 1;
  rows.forEach(function (r) { if (r.length > colCount) colCount = r.length; });

  var widths = [];
  for (var c = 0; c < colCount; c++) {
    var w = 8;
    for (var r = 0; r < rows.length; r++) {
      var v = rows[r][c];
      if (v === null || v === undefined || v === '') continue;
      var vw = visualWidth(v);
      if (vw > w) w = vw;
    }
    widths.push(Math.min(Math.max(w + 4, 10), 46));
  }

  var xml = [];
  xml.push('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>');
  xml.push('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">');
  xml.push('<dimension ref="A1:' + colName(colCount - 1) + Math.max(rows.length, 1) + '"/>');
  xml.push('<sheetViews><sheetView workbookViewId="0">' +
    '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>' +
    '<selection pane="bottomLeft" activeCell="A2" sqref="A2"/>' +
    '</sheetView></sheetViews>');
  xml.push('<sheetFormatPr defaultRowHeight="18"/>');

  xml.push('<cols>');
  widths.forEach(function (w, i) {
    xml.push('<col min="' + (i + 1) + '" max="' + (i + 1) + '" width="' + w + '" customWidth="1"/>');
  });
  xml.push('</cols>');

  xml.push('<sheetData>');
  rows.forEach(function (row, ri) {
    var rowNum = ri + 1;
    var cells = '';
    for (var ci = 0; ci < row.length; ci++) {
      var val = row[ci];
      if (val === null || val === undefined || val === '') continue;
      var ref = colName(ci) + rowNum;
      var style = ri === 0 ? ' s="1"' : '';
      if (typeof val === 'number' && isFinite(val)) {
        cells += '<c r="' + ref + '"' + style + '><v>' + val + '</v></c>';
      } else {
        cells += '<c r="' + ref + '"' + style + ' t="inlineStr"><is><t xml:space="preserve">' +
          xmlEsc(val) + '</t></is></c>';
      }
    }
    var rowAttr = ri === 0 ? ' ht="26" customHeight="1"' : '';
    xml.push('<row r="' + rowNum + '"' + rowAttr + '>' + cells + '</row>');
  });
  xml.push('</sheetData>');
  xml.push('<autoFilter ref="A1:' + colName(colCount - 1) + '1"/>');
  xml.push('</worksheet>');
  return xml.join('');
}

var STYLES_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
  '<fonts count="2">' +
    '<font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/></font>' +
    '<font><b/><sz val="11"/><color rgb="FF1F2937"/><name val="Calibri"/><family val="2"/></font>' +
  '</fonts>' +
  '<fills count="3">' +
    '<fill><patternFill patternType="none"/></fill>' +
    '<fill><patternFill patternType="gray125"/></fill>' +
    '<fill><patternFill patternType="solid"><fgColor rgb="FFEEF0FF"/><bgColor indexed="64"/></patternFill></fill>' +
  '</fills>' +
  '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
  '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
  '<cellXfs count="2">' +
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
    '<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>' +
  '</cellXfs>' +
  '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
  '</styleSheet>';

/* 把若干 {name, data} 打成一个 zip（store 模式，不压缩） */
function zipStore(files) {
  var now = new Date();
  var dosTime = ((now.getHours() << 11) | (now.getMinutes() << 5) | (now.getSeconds() >> 1)) & 0xFFFF;
  var dosDate = (((now.getFullYear() - 1980) << 9) | ((now.getMonth() + 1) << 5) | now.getDate()) & 0xFFFF;

  var chunks = [];
  var central = [];
  var offset = 0;

  files.forEach(function (f) {
    var nameBytes = enc.encode(f.name);
    var body = f.data;
    var crc = crc32(body);

    /* 本地文件头 */
    var lh = new Uint8Array(30 + nameBytes.length);
    var lv = new DataView(lh.buffer);
    lv.setUint32(0, 0x04034b50, true);
    lv.setUint16(4, 20, true);
    lv.setUint16(6, 0x0800, true);          /* 文件名用 UTF-8 */
    lv.setUint16(8, 0, true);               /* 0 = 不压缩 */
    lv.setUint16(10, dosTime, true);
    lv.setUint16(12, dosDate, true);
    lv.setUint32(14, crc, true);
    lv.setUint32(18, body.length, true);
    lv.setUint32(22, body.length, true);
    lv.setUint16(26, nameBytes.length, true);
    lv.setUint16(28, 0, true);
    lh.set(nameBytes, 30);

    chunks.push(lh, body);

    /* 中央目录项 */
    var ch = new Uint8Array(46 + nameBytes.length);
    var cv = new DataView(ch.buffer);
    cv.setUint32(0, 0x02014b50, true);
    cv.setUint16(4, 20, true);
    cv.setUint16(6, 20, true);
    cv.setUint16(8, 0x0800, true);
    cv.setUint16(10, 0, true);
    cv.setUint16(12, dosTime, true);
    cv.setUint16(14, dosDate, true);
    cv.setUint32(16, crc, true);
    cv.setUint32(20, body.length, true);
    cv.setUint32(24, body.length, true);
    cv.setUint16(28, nameBytes.length, true);
    cv.setUint16(30, 0, true);
    cv.setUint16(32, 0, true);
    cv.setUint16(34, 0, true);
    cv.setUint16(36, 0, true);
    cv.setUint32(38, 0, true);
    cv.setUint32(42, offset, true);
    ch.set(nameBytes, 46);
    central.push(ch);

    offset += lh.length + body.length;
  });

  var centralSize = central.reduce(function (a, c) { return a + c.length; }, 0);
  var eocd = new Uint8Array(22);
  var ev = new DataView(eocd.buffer);
  ev.setUint32(0, 0x06054b50, true);
  ev.setUint16(4, 0, true);
  ev.setUint16(6, 0, true);
  ev.setUint16(8, files.length, true);
  ev.setUint16(10, files.length, true);
  ev.setUint32(12, centralSize, true);
  ev.setUint32(16, offset, true);
  ev.setUint16(20, 0, true);

  var all = chunks.concat(central, [eocd]);
  var total = all.reduce(function (a, c) { return a + c.length; }, 0);
  var out = new Uint8Array(total);
  var pos = 0;
  all.forEach(function (c) { out.set(c, pos); pos += c.length; });
  return out;
}

/**
 * 把二维数组写成 .xlsx 的字节流
 * @param {Array<Array>} rows 第一行是表头
 * @param {string} sheetName 工作表名字
 * @returns {Uint8Array}
 */
function makeXlsx(rows, sheetName) {
  var safeName = String(sheetName || 'Sheet1').replace(/[\\\/\?\*\[\]:]/g, ' ').slice(0, 28) || 'Sheet1';

  var contentTypes =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
    '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
    '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
    '</Types>';

  var rels =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
    '</Relationships>';

  var workbook =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
    '<sheets><sheet name="' + attrEsc(safeName) + '" sheetId="1" r:id="rId1"/></sheets>' +
    '</workbook>';

  var workbookRels =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
    '</Relationships>';

  return zipStore([
    { name: '[Content_Types].xml', data: enc.encode(contentTypes) },
    { name: '_rels/.rels', data: enc.encode(rels) },
    { name: 'xl/workbook.xml', data: enc.encode(workbook) },
    { name: 'xl/_rels/workbook.xml.rels', data: enc.encode(workbookRels) },
    { name: 'xl/styles.xml', data: enc.encode(STYLES_XML) },
    { name: 'xl/worksheets/sheet1.xml', data: enc.encode(buildSheetXml(rows)) }
  ]);
}

/* =========================================================================
 * 第 2 部分  Excel 读取
 * 要读别人用 Excel 存的文件，就必须支持 zip 的 deflate 压缩，
 * 所以下面手写了一个解压器（inflate）。
 * ========================================================================= */

var LBASE = [3, 4, 5, 6, 7, 8, 9, 10, 11, 13, 15, 17, 19, 23, 27, 31, 35, 43, 51, 59, 67, 83, 99, 115, 131, 163, 195, 227, 258];
var LEXT  = [0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3, 3, 4, 4, 4, 4, 5, 5, 5, 5, 0];
var DBASE = [1, 2, 3, 4, 5, 7, 9, 13, 17, 25, 33, 49, 65, 97, 129, 193, 257, 385, 513, 769, 1025, 1537, 2049, 3073, 4097, 6145, 8193, 12289, 16385, 24577];
var DEXT  = [0, 0, 0, 0, 1, 1, 2, 2, 3, 3, 4, 4, 5, 5, 6, 6, 7, 7, 8, 8, 9, 9, 10, 10, 11, 11, 12, 12, 13, 13];

/* 根据「每个符号的编码长度」建一张查表 */
function buildHuffman(lengths) {
  var maxBits = 0;
  for (var i = 0; i < lengths.length; i++) if (lengths[i] > maxBits) maxBits = lengths[i];

  var counts = new Array(maxBits + 1);
  for (var j = 0; j <= maxBits; j++) counts[j] = 0;
  for (i = 0; i < lengths.length; i++) counts[lengths[i]]++;
  counts[0] = 0;

  var offs = new Array(maxBits + 2);
  for (j = 0; j <= maxBits + 1; j++) offs[j] = 0;
  for (j = 1; j <= maxBits; j++) offs[j + 1] = offs[j] + counts[j];

  var symbols = new Array(lengths.length);
  for (i = 0; i < lengths.length; i++) {
    if (lengths[i]) symbols[offs[lengths[i]]++] = i;
  }
  return { maxBits: maxBits, counts: counts, symbols: symbols };
}

/* inflate：把 deflate 压缩后的数据还原（zip 用的就是它） */
function inflateRaw(data) {
  var out = new Uint8Array(Math.max(1024, data.length * 4));
  var olen = 0;
  var pos = 0;
  var bitbuf = 0;
  var bitcnt = 0;

  function need(n) {
    if (olen + n <= out.length) return;
    var nb = new Uint8Array(Math.max(out.length * 2, olen + n));
    nb.set(out.subarray(0, olen));
    out = nb;
  }

  function bits(n) {
    while (bitcnt < n) {
      if (pos >= data.length) throw new Error('压缩数据不完整');
      bitbuf |= data[pos++] << bitcnt;
      bitcnt += 8;
    }
    if (n === 0) return 0;
    var v = bitbuf & ((1 << n) - 1);
    bitbuf >>>= n;
    bitcnt -= n;
    return v;
  }

  function decodeSymbol(t) {
    var code = 0;
    var first = 0;
    var index = 0;
    for (var len = 1; len <= t.maxBits; len++) {
      code |= bits(1);
      var count = t.counts[len];
      if (code - first < count) return t.symbols[index + code - first];
      index += count;
      first = (first + count) << 1;
      code <<= 1;
    }
    throw new Error('压缩数据里有非法编码');
  }

  var fixedLit = null;
  var fixedDist = null;
  function getFixed() {
    if (fixedLit) return;
    var l = new Array(288);
    var i = 0;
    for (; i < 144; i++) l[i] = 8;
    for (; i < 256; i++) l[i] = 9;
    for (; i < 280; i++) l[i] = 7;
    for (; i < 288; i++) l[i] = 8;
    fixedLit = buildHuffman(l);
    var d = new Array(30);
    for (i = 0; i < 30; i++) d[i] = 5;
    fixedDist = buildHuffman(d);
  }

  function inflateBlock(lit, dist) {
    for (;;) {
      var sym = decodeSymbol(lit);
      if (sym < 256) {
        need(1);
        out[olen++] = sym;
      } else if (sym === 256) {
        return;
      } else {
        var li = sym - 257;
        if (li >= LBASE.length) throw new Error('压缩数据里有非法长度码');
        var len = LBASE[li] + (LEXT[li] ? bits(LEXT[li]) : 0);
        var ds = decodeSymbol(dist);
        if (ds >= DBASE.length) throw new Error('压缩数据里有非法距离码');
        var back = DBASE[ds] + (DEXT[ds] ? bits(DEXT[ds]) : 0);
        need(len);
        for (var k = 0; k < len; k++) {
          out[olen] = out[olen - back];
          olen++;
        }
      }
    }
  }

  var last = 0;
  while (!last) {
    last = bits(1);
    var type = bits(2);

    if (type === 0) {
      /* 不压缩的块 */
      var drop = bitcnt & 7;
      bitbuf >>>= drop;
      bitcnt -= drop;
      var len0 = data[pos] | (data[pos + 1] << 8);
      pos += 4;                       /* 跳过 len 和 nlen */
      need(len0);
      out.set(data.subarray(pos, pos + len0), olen);
      olen += len0;
      pos += len0;
    } else if (type === 1) {
      getFixed();
      inflateBlock(fixedLit, fixedDist);
    } else if (type === 2) {
      var hlit = bits(5) + 257;
      var hdist = bits(5) + 1;
      var hclen = bits(4) + 4;
      var order = [16, 17, 18, 0, 8, 7, 9, 6, 10, 5, 11, 4, 12, 3, 13, 2, 14, 1, 15];
      var clens = new Array(19);
      for (var ci = 0; ci < 19; ci++) clens[ci] = 0;
      for (ci = 0; ci < hclen; ci++) clens[order[ci]] = bits(3);
      var clTable = buildHuffman(clens);

      var lens = [];
      while (lens.length < hlit + hdist) {
        var s = decodeSymbol(clTable);
        if (s < 16) {
          lens.push(s);
        } else if (s === 16) {
          var rep = bits(2) + 3;
          var prev = lens[lens.length - 1] || 0;
          for (var r1 = 0; r1 < rep; r1++) lens.push(prev);
        } else if (s === 17) {
          var rep2 = bits(3) + 3;
          for (var r2 = 0; r2 < rep2; r2++) lens.push(0);
        } else {
          var rep3 = bits(7) + 11;
          for (var r3 = 0; r3 < rep3; r3++) lens.push(0);
        }
      }
      inflateBlock(
        buildHuffman(lens.slice(0, hlit)),
        buildHuffman(lens.slice(hlit, hlit + hdist))
      );
    } else {
      throw new Error('压缩数据损坏（块类型错误）');
    }
  }

  return out.subarray(0, olen);
}

/* 解压整个 zip，返回 { 文件名: 内容字节 } */
function unzip(bytes) {
  var dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  var eocd = -1;
  var minPos = Math.max(0, bytes.length - 22 - 65535);
  for (var i = bytes.length - 22; i >= minPos; i--) {
    if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
  }
  if (eocd < 0) throw new Error('这不是一个有效的 Excel 文件');

  var count = dv.getUint16(eocd + 10, true);
  var cdOffset = dv.getUint32(eocd + 16, true);
  var files = {};
  var p = cdOffset;

  for (var k = 0; k < count; k++) {
    if (dv.getUint32(p, true) !== 0x02014b50) throw new Error('Excel 文件结构损坏');
    var method = dv.getUint16(p + 10, true);
    var csize = dv.getUint32(p + 20, true);
    var nlen = dv.getUint16(p + 28, true);
    var elen = dv.getUint16(p + 30, true);
    var clen = dv.getUint16(p + 32, true);
    var lho = dv.getUint32(p + 42, true);
    var name = dec.decode(bytes.subarray(p + 46, p + 46 + nlen));

    var lnlen = dv.getUint16(lho + 26, true);
    var lelen = dv.getUint16(lho + 28, true);
    var dstart = lho + 30 + lnlen + lelen;
    var raw = bytes.subarray(dstart, dstart + csize);

    if (method === 0) {
      files[name] = raw;
    } else if (method === 8) {
      files[name] = inflateRaw(raw);
    }
    p += 46 + nlen + elen + clen;
  }
  return files;
}

function parseXml(text) {
  var doc = new DOMParser().parseFromString(text, 'application/xml');
  var err = doc.getElementsByTagName('parsererror');
  if (err.length) throw new Error('Excel 内部的 XML 解析失败');
  return doc;
}

function colIndexFromRef(ref) {
  var m = /^([A-Z]+)/.exec(ref || '');
  if (!m) return 0;
  var s = m[1];
  var n = 0;
  for (var i = 0; i < s.length; i++) n = n * 26 + (s.charCodeAt(i) - 64);
  return n - 1;
}

function textOf(el) {
  var ts = el.getElementsByTagName('t');
  var s = '';
  for (var i = 0; i < ts.length; i++) s += ts[i].textContent;
  return s;
}

/**
 * 读一个 .xlsx 文件，返回二维数组（第一行一般是表头）
 * @param {Uint8Array|ArrayBuffer} input
 * @returns {Array<Array<string|number>>}
 */
function readXlsx(input) {
  var bytes = input instanceof Uint8Array ? input : new Uint8Array(input);
  var files = unzip(bytes);

  if (!files['xl/workbook.xml']) throw new Error('文件里找不到工作表，可能不是 .xlsx 格式');

  var wb = parseXml(dec.decode(files['xl/workbook.xml']));
  var sheetEls = wb.getElementsByTagName('sheet');
  if (!sheetEls.length) throw new Error('Excel 文件里没有工作表');

  var RID = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  var rid = sheetEls[0].getAttributeNS(RID, 'id') || sheetEls[0].getAttribute('r:id');

  var target = null;
  if (files['xl/_rels/workbook.xml.rels']) {
    var relsDoc = parseXml(dec.decode(files['xl/_rels/workbook.xml.rels']));
    var rels = relsDoc.getElementsByTagName('Relationship');
    for (var i = 0; i < rels.length; i++) {
      if (rels[i].getAttribute('Id') === rid) {
        target = rels[i].getAttribute('Target');
        break;
      }
    }
  }
  if (!target) target = 'worksheets/sheet1.xml';
  var sheetPath = target.charAt(0) === '/' ? target.slice(1) : 'xl/' + target;
  var sheetBytes = files[sheetPath];
  if (!sheetBytes) throw new Error('Excel 文件里找不到第一个工作表的内容');

  /* 共享字符串表 */
  var shared = [];
  if (files['xl/sharedStrings.xml']) {
    var sst = parseXml(dec.decode(files['xl/sharedStrings.xml']));
    var sis = sst.getElementsByTagName('si');
    for (var s = 0; s < sis.length; s++) shared.push(textOf(sis[s]));
  }

  var sheet = parseXml(dec.decode(sheetBytes));
  var rowEls = sheet.getElementsByTagName('row');
  var rows = [];

  for (var r = 0; r < rowEls.length; r++) {
    var rowEl = rowEls[r];
    var rowNum = parseInt(rowEl.getAttribute('r'), 10);
    var rowIdx = isNaN(rowNum) ? rows.length : rowNum - 1;
    var cells = rowEl.getElementsByTagName('c');
    var line = [];

    for (var c = 0; c < cells.length; c++) {
      var cell = cells[c];
      var ref = cell.getAttribute('r');
      var colIdx = ref ? colIndexFromRef(ref) : line.length;
      var type = cell.getAttribute('t');
      var value = '';

      if (type === 'inlineStr') {
        var is = cell.getElementsByTagName('is')[0];
        value = is ? textOf(is) : '';
      } else {
        var vEl = cell.getElementsByTagName('v')[0];
        var raw = vEl ? vEl.textContent : '';
        if (type === 's') {
          value = shared[parseInt(raw, 10)] || '';
        } else if (type === 'b') {
          value = raw === '1' ? 'TRUE' : 'FALSE';
        } else if (type === 'str' || type === 'e') {
          value = raw;
        } else {
          value = raw === '' ? '' : (isNaN(Number(raw)) ? raw : Number(raw));
        }
      }
      while (line.length < colIdx) line.push('');
      line[colIdx] = value;
    }

    while (rows.length < rowIdx) rows.push([]);
    rows[rowIdx] = line;
  }

  return rows;
}

/* 给测试用的出口（浏览器控制台里也能调 XLSXlite 检查） */
window.XLSXlite = { write: makeXlsx, read: readXlsx, zip: zipStore, unzip: unzip, inflate: inflateRaw };

/* =========================================================================
 * 第 3 部分  数据层
 * ========================================================================= */

var STORAGE_KEY = 'addressbook.contacts.v1';

var state = {
  contacts: [],
  filter: 'all',
  keyword: '',
  editingId: null
};

function blankContact() {
  return {
    id: uid(),
    name: '',
    company: '',
    bookmark: false,
    phones: [],
    emails: [],
    socials: [],     /* [{label:'微信', value:'abc'}] */
    addresses: [],
    notes: '',
    createdAt: Date.now()
  };
}

function normalizeContact(raw) {
  var c = blankContact();
  c.id = raw.id || c.id;
  c.name = String(raw.name || '').trim();
  c.company = String(raw.company || '').trim();
  c.bookmark = !!raw.bookmark;
  c.notes = String(raw.notes || '');
  c.createdAt = raw.createdAt || Date.now();

  var toArr = function (v) {
    if (!v) return [];
    if (Array.isArray(v)) return v.map(function (x) { return String(x || '').trim(); }).filter(Boolean);
    return String(v).split(/\r?\n|;|；/).map(function (x) { return x.trim(); }).filter(Boolean);
  };

  c.phones = toArr(raw.phones);
  c.emails = toArr(raw.emails);
  c.addresses = toArr(raw.addresses);
  c.socials = (Array.isArray(raw.socials) ? raw.socials : []).map(function (s) {
    if (typeof s === 'string') return { label: '社交', value: s.trim() };
    return { label: String(s.label || '社交').trim(), value: String(s.value || '').trim() };
  }).filter(function (s) { return s.value; });

  return c;
}

function loadContacts() {
  try {
    var raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return;
    var list = JSON.parse(raw);
    if (Array.isArray(list)) state.contacts = list.map(normalizeContact);
  } catch (e) {
    console.warn('读取本地数据失败：', e);
  }
}

function saveContacts() {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(state.contacts));
  } catch (e) {
    toast('本地保存失败，可能是浏览器隐私模式限制', true);
  }
}

/* =========================================================================
 * 第 4 部分  界面渲染
 * ========================================================================= */

var ICONS = {
  phone: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M6.6 10.8a15 15 0 0 0 6.6 6.6l2.2-2.2a1 1 0 0 1 1-.23 11.4 11.4 0 0 0 3.6.58 1 1 0 0 1 1 1V20a1 1 0 0 1-1 1A17 17 0 0 1 3 4a1 1 0 0 1 1-1h3.5a1 1 0 0 1 1 1 11.4 11.4 0 0 0 .58 3.6 1 1 0 0 1-.24 1z"/></svg>',
  mail: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="5" width="18" height="14" rx="2"/><path d="M3.5 6.5L12 13l8.5-6.5"/></svg>',
  chat: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M20 12.5c0 3.6-3.6 6.5-8 6.5a9.6 9.6 0 0 1-2.6-.35L5 20.5l1.1-3.2A6.3 6.3 0 0 1 4 12.5C4 8.9 7.6 6 12 6s8 2.9 8 6.5z"/></svg>',
  pin: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M12 21s7-5.6 7-11a7 7 0 1 0-14 0c0 5.4 7 11 7 11z"/><circle cx="12" cy="10" r="2.6"/></svg>',
  star: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linejoin="round"><path d="M12 3.6l2.6 5.3 5.9.86-4.27 4.16 1 5.88L12 17l-5.23 2.8 1-5.88L3.5 9.76l5.9-.86z"/></svg>'
};

function hashHue(text) {
  var h = 0;
  for (var i = 0; i < text.length; i++) h = (h * 31 + text.charCodeAt(i)) % 360;
  return h;
}

function initial(name) {
  var s = String(name || '?').trim();
  return s ? s.charAt(0).toUpperCase() : '?';
}

function visibleContacts() {
  var kw = state.keyword.trim().toLowerCase();
  var list = state.contacts.filter(function (c) {
    if (state.filter === 'bookmark' && !c.bookmark) return false;
    if (!kw) return true;
    var parts = [c.name, c.company, c.notes].concat(c.phones, c.emails, c.addresses);
    c.socials.forEach(function (s) { parts.push(s.label + ' ' + s.value); });
    return parts.join(' ').toLowerCase().indexOf(kw) >= 0;
  });

  return list.sort(function (a, b) {
    if (a.bookmark !== b.bookmark) return a.bookmark ? -1 : 1;   /* 收藏的永远排在最前 */
    var byName = String(a.name).localeCompare(String(b.name), 'zh-Hans-CN');
    return byName !== 0 ? byName : a.createdAt - b.createdAt;
  });
}

function metaRow(icon, values, withTag) {
  if (!values || !values.length) return '';
  var inner = values.map(function (v) {
    if (withTag) {
      return '<span class="val"><span class="tag">' + xmlEsc(v.label) + '</span>' + xmlEsc(v.value) + '</span>';
    }
    return '<span class="val">' + xmlEsc(v) + '</span>';
  }).join('');
  return '<li><span class="ico">' + icon + '</span><div class="vals">' + inner + '</div></li>';
}

function cardHtml(c) {
  var meta = '';
  meta += metaRow(ICONS.phone, c.phones, false);
  meta += metaRow(ICONS.mail, c.emails, false);
  meta += metaRow(ICONS.chat, c.socials, true);
  meta += metaRow(ICONS.pin, c.addresses, false);

  var sub = c.company || (c.bookmark ? '已收藏' : '未填写单位');
  var footNote = c.notes ? '备注：' + c.notes : '';

  return '' +
    '<article class="card' + (c.bookmark ? ' is-mark' : '') + '" data-id="' + c.id + '">' +
      '<button class="star' + (c.bookmark ? ' on' : '') + '" data-act="star" title="' +
        (c.bookmark ? '取消收藏' : '收藏（置顶显示）') + '" aria-label="收藏">' + ICONS.star + '</button>' +
      '<div class="card-head">' +
        '<div class="avatar" style="--h:' + hashHue(c.name || c.id) + '">' + xmlEsc(initial(c.name)) + '</div>' +
        '<div class="who"><h2>' + xmlEsc(c.name || '未命名') + '</h2><p>' + xmlEsc(sub) + '</p></div>' +
      '</div>' +
      (meta ? '<ul class="meta">' + meta + '</ul>' : '<div class="meta-empty"></div>') +
      '<div class="card-foot">' +
        '<span class="when">' + xmlEsc(footNote) + '</span>' +
        '<button class="act" data-act="edit">编辑</button>' +
        '<button class="act del" data-act="del">删除</button>' +
      '</div>' +
    '</article>';
}

function render() {
  var listEl = $('#list');
  if (!listEl) return;

  var list = visibleContacts();
  listEl.innerHTML = list.map(cardHtml).join('');

  $('#countAll').textContent = state.contacts.length;
  $('#countMark').textContent = state.contacts.filter(function (c) { return c.bookmark; }).length;

  var emptyEl = $('#empty');
  var isEmpty = list.length === 0;
  emptyEl.hidden = !isEmpty;
  listEl.hidden = isEmpty;

  if (isEmpty) {
    if (state.contacts.length === 0) {
      $('#emptyTitle').textContent = '还没有联系人';
      $('#emptyText').textContent = '新建第一个联系人，或者先载入几条示例数据看看效果。';
      $('#seedBtn').hidden = false;
    } else {
      $('#emptyTitle').textContent = state.filter === 'bookmark' ? '还没有收藏的联系人' : '没有匹配的联系人';
      $('#emptyText').textContent = state.filter === 'bookmark'
        ? '在联系人卡片右上角点一下星标，它就会出现在这里并置顶显示。'
        : '换个关键词试试，或者切换到「全部」。';
      $('#seedBtn').hidden = true;
    }
  }
}

/* =========================================================================
 * 第 5 部分  表单（新建 / 编辑）
 * ========================================================================= */

var KIND_PLACEHOLDER = {
  phones: '电话号码，例如 138 0000 0000',
  emails: '电子邮箱，例如 wang@example.com',
  addresses: '详细地址，例如 福建省福州市福州大学'
};

function dynList(kind) {
  return $('.group[data-kind="' + kind + '"] .dyn-list');
}

function onRowDelete(kind) {
  var list = dynList(kind);
  var rows = list.children;
  if (rows.length === 0) {
    var hint = document.createElement('div');
    hint.className = 'dyn-empty';
    hint.textContent = '还没有添加，点右上角「+ 添加」';
    list.appendChild(hint);
  }
  linkGroupEvents(kind);
}

function linkGroupEvents(kind) {
  var list = dynList(kind);
  $$('.row-del', list).forEach(function (btn) {
    btn.onclick = function () {
      btn.parentNode.remove();
      onRowDelete(kind);
    };
  });
}

function clearDynamic() {
  ['phones', 'emails', 'socials', 'addresses'].forEach(function (kind) {
    dynList(kind).innerHTML = '';
  });
}

function addFieldRow(kind, value, label) {
  var list = dynList(kind);
  var hint = $('.dyn-empty', list);
  if (hint) hint.remove();

  var row = document.createElement('div');
  row.className = 'dyn-row' + (kind === 'socials' ? ' two' : '');

  var delBtn = '<button type="button" class="row-del" title="删除这一行">' +
    '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><path d="M6 6l12 12M18 6L6 18"/></svg>' +
    '</button>';

  if (kind === 'socials') {
    row.innerHTML =
      '<input class="label-input" maxlength="12" placeholder="平台" value="' + attrEsc(label || '') + '">' +
      '<input class="value-input" maxlength="120" placeholder="微信号 / 账号 / 链接" value="' + attrEsc(value || '') + '">' +
      delBtn;
  } else {
    row.innerHTML =
      '<input class="value-input" maxlength="120" placeholder="' + KIND_PLACEHOLDER[kind] + '" value="' +
      attrEsc(value || '') + '">' + delBtn;
  }

  list.appendChild(row);
  linkGroupEvents(kind);
  return row;
}

function ensureEmptyHints() {
  ['phones', 'emails', 'socials', 'addresses'].forEach(function (kind) {
    var list = dynList(kind);
    if (!list.children.length) {
      var hint = document.createElement('div');
      hint.className = 'dyn-empty';
      hint.textContent = '还没有添加，点右上角「+ 添加」';
      list.appendChild(hint);
    }
  });
}

function openEditor(id) {
  var c = id ? state.contacts.filter(function (x) { return x.id === id; })[0] : null;
  state.editingId = c ? c.id : null;

  $('#modalTitle').textContent = c ? '编辑联系人' : '新建联系人';
  $('#fName').value = c ? c.name : '';
  $('#fCompany').value = c ? c.company : '';
  $('#fBookmark').checked = c ? c.bookmark : false;
  $('#fNotes').value = c ? c.notes : '';
  $('#formErr').textContent = '';

  clearDynamic();
  if (c) {
    c.phones.forEach(function (v) { addFieldRow('phones', v); });
    c.emails.forEach(function (v) { addFieldRow('emails', v); });
    c.socials.forEach(function (s) { addFieldRow('socials', s.value, s.label); });
    c.addresses.forEach(function (v) { addFieldRow('addresses', v); });
  }
  ensureEmptyHints();

  $('#overlay').hidden = false;
  document.body.classList.add('is-locked');
  setTimeout(function () { $('#fName').focus(); }, 30);
}

function closeEditor() {
  $('#overlay').hidden = true;
  if ($('#dialogOverlay').hidden) document.body.classList.remove('is-locked');
}

function collectValues(kind) {
  return $$('.dyn-row', dynList(kind)).map(function (row) {
    var val = $('.value-input', row).value.trim();
    if (!val) return null;
    if (kind === 'socials') {
      var label = $('.label-input', row).value.trim() || '社交';
      return { label: label, value: val };
    }
    return val;
  }).filter(Boolean);
}

function saveContact() {
  var name = $('#fName').value.trim();
  if (!name) {
    $('#formErr').textContent = '请先填写姓名';
    $('#fName').focus();
    return;
  }

  var data = {
    name: name,
    company: $('#fCompany').value.trim(),
    bookmark: $('#fBookmark').checked,
    phones: collectValues('phones'),
    emails: collectValues('emails'),
    socials: collectValues('socials'),
    addresses: collectValues('addresses'),
    notes: $('#fNotes').value.trim()
  };

  if (state.editingId) {
    state.contacts = state.contacts.map(function (c) {
      if (c.id !== state.editingId) return c;
      return Object.assign({}, c, data);
    });
    toast('已保存修改');
  } else {
    var c = blankContact();
    Object.assign(c, data);
    state.contacts.push(c);
    toast('已添加联系人');
  }

  saveContacts();
  render();
  closeEditor();
}

/* =========================================================================
 * 第 6 部分  导入 / 导出 Excel
 * ========================================================================= */

/* 出一张二维表：表头 + 每个联系人一行 */
function buildExportRows() {
  var maxPhones = 1, maxEmails = 1, maxSocials = 1, maxAddresses = 1;
  state.contacts.forEach(function (c) {
    maxPhones = Math.max(maxPhones, c.phones.length);
    maxEmails = Math.max(maxEmails, c.emails.length);
    maxSocials = Math.max(maxSocials, c.socials.length);
    maxAddresses = Math.max(maxAddresses, c.addresses.length);
  });

  var header = ['姓名', '收藏', '公司/院系'];
  for (var i = 1; i <= maxPhones; i++) header.push('电话' + i);
  for (i = 1; i <= maxEmails; i++) header.push('邮箱' + i);
  for (i = 1; i <= maxSocials; i++) header.push('社交' + i);
  for (i = 1; i <= maxAddresses; i++) header.push('地址' + i);
  header.push('备注');

  var rows = [header];
  state.contacts.forEach(function (c) {
    var row = [c.name, c.bookmark ? '是' : '否', c.company];
    for (var p = 0; p < maxPhones; p++) row.push(c.phones[p] || '');
    for (var e = 0; e < maxEmails; e++) row.push(c.emails[e] || '');
    for (var s = 0; s < maxSocials; s++) {
      var social = c.socials[s];
      row.push(social ? (social.label + ': ' + social.value) : '');
    }
    for (var a = 0; a < maxAddresses; a++) row.push(c.addresses[a] || '');
    row.push(c.notes);
    rows.push(row);
  });
  return rows;
}

function exportExcel() {
  if (!state.contacts.length) {
    toast('通讯录是空的，先加几个联系人吧', true);
    return;
  }
  try {
    var rows = buildExportRows();
    var bytes = makeXlsx(rows, 'Contacts');
    download(bytes, '通讯录-' + today() + '.xlsx',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    toast('已导出 ' + state.contacts.length + ' 位联系人');
  } catch (e) {
    console.error(e);
    toast('导出失败：' + e.message, true);
  }
}

/* 把表头文字归一化，好做模糊匹配 */
function headerKey(raw) {
  return String(raw || '')
    .replace(/[\s\-_\/\\.、,，()（）\[\]【】:：]/g, '')
    .replace(/\d+$/, '')
    .toLowerCase();
}

function matchHeader(key) {
  if (/^(姓名|名字|联系人|名称|name|fullname)$/.test(key)) return 'name';
  if (/^(收藏|书签|收藏夹|bookmark|favorite|favourite|star)$/.test(key)) return 'bookmark';
  if (/^(公司|单位|院系|部门|学校|company|org|organization|department)/.test(key)) return 'company';
  if (/^(电话|手机|手机号|电话号码|联系电话|联系方式|phone|mobile|tel|telephone)$/.test(key)) return 'phone';
  if (/^(邮箱|电子邮箱|邮件|电子邮件|email|mail|e-mail)$/.test(key)) return 'email';
  if (/^(社交|社交账号|社交平台|微信|qq|wechat|social|im|weibo|微博)$/.test(key)) return 'social';
  if (/^(地址|住址|通讯地址|家庭住址|address)$/.test(key)) return 'address';
  if (/^(备注|说明|note|notes|remark|comment)$/.test(key)) return 'notes';
  return null;
}

function parseBookmark(v) {
  var s = String(v).trim().toLowerCase();
  return ['是', '已收藏', 'y', 'yes', 'true', '1', '★', '√', 'v', 'x'].indexOf(s) >= 0;
}

/* 二维表 -> 联系人数组 */
function rowsToContacts(rows) {
  var clean = rows.filter(function (r) {
    return r && r.some(function (v) { return v !== '' && v !== null && v !== undefined; });
  });
  if (!clean.length) return { contacts: [], skipped: 0 };

  var headerRow = clean[0].map(headerKey);
  var mapped = headerRow.map(matchHeader);

  /* 第一行不像表头（没有姓名列）时，当作没有表头，按默认顺序解析 */
  var hasHeader = mapped.indexOf('name') >= 0;
  var body;
  if (hasHeader) {
    body = clean.slice(1);
  } else {
    mapped = ['name', 'bookmark', 'company'];
    body = clean;
  }

  var contacts = [];
  var skipped = 0;

  body.forEach(function (row) {
    var c = { name: '', bookmark: false, company: '', phones: [], emails: [], socials: [], addresses: [], notes: '' };

    mapped.forEach(function (field, idx) {
      var val = row[idx];
      if (val === null || val === undefined || val === '') return;
      val = String(val).trim();
      if (!val) return;

      if (field === 'name') c.name = val;
      else if (field === 'bookmark') c.bookmark = parseBookmark(val);
      else if (field === 'company') c.company = val;
      else if (field === 'phone') c.phones.push(val);
      else if (field === 'email') c.emails.push(val);
      else if (field === 'address') c.addresses.push(val);
      else if (field === 'notes') c.notes = c.notes ? c.notes + ' ' + val : val;
      else if (field === 'social') {
        var m = /^([^:：]{1,12})[:：]\s*(.+)$/.exec(val);
        if (m) c.socials.push({ label: m[1].trim(), value: m[2].trim() });
        else c.socials.push({ label: '社交', value: val });
      }
    });

    if (!c.name) { skipped++; return; }
    contacts.push(c);
  });

  return { contacts: contacts, skipped: skipped };
}

function importFromFile(file) {
  var reader = new FileReader();

  reader.onload = function () {
    try {
      var bytes = new Uint8Array(reader.result);

      /* 老的 .xls（二进制格式）不支持，提前给个明确提示 */
      var isZip = bytes.length > 4 && bytes[0] === 0x50 && bytes[1] === 0x4B;
      var isOle = bytes.length > 8 && bytes[0] === 0xD0 && bytes[1] === 0xCF;
      if (isOle) throw new Error('这是旧版 .xls 格式，请在 Excel 里另存为 .xlsx 再导入');
      if (!isZip) throw new Error('文件不是 .xlsx 格式，请用 Excel 另存为 .xlsx');

      var rows = readXlsx(bytes);
      var result = rowsToContacts(rows);

      if (!result.contacts.length) {
        toast('文件里没有读到联系人（也许少了「姓名」这一列）', true);
        return;
      }

      askImportMode(result.contacts, result.skipped);
    } catch (e) {
      console.error(e);
      toast('导入失败：' + e.message, true);
    }
  };

  reader.onerror = function () { toast('文件读取失败', true); };
  reader.readAsArrayBuffer(file);
}

function askImportMode(contacts, skipped) {
  var text = '从文件里读到 ' + contacts.length + ' 位联系人。\n' +
    '「追加」= 加到现有通讯录后面；「覆盖」= 清空现有数据后再导入。' +
    (skipped ? '\n（另有 ' + skipped + ' 行没有姓名，已跳过）' : '');

  dialog({
    title: '导入方式',
    text: text,
    okText: '追加导入',
    altText: '覆盖全部',
    onOk: function () {
      var added = contacts.map(function (raw) {
        var c = normalizeContact(raw);
        c.id = uid();
        c.createdAt = Date.now();
        return c;
      });
      state.contacts = state.contacts.concat(added);
      saveContacts();
      render();
      toast('已追加 ' + added.length + ' 位联系人');
    },
    onAlt: function () {
      state.contacts = contacts.map(function (raw) {
        var c = normalizeContact(raw);
        c.id = uid();
        c.createdAt = Date.now();
        return c;
      });
      saveContacts();
      render();
      toast('已导入 ' + state.contacts.length + ' 位联系人');
    }
  });
}

/* 通用弹窗：确定 / 可选第二按钮 / 取消 */
function dialog(opts) {
  var overlay = $('#dialogOverlay');
  $('#dialogTitle').textContent = opts.title || '提示';
  $('#dialogText').textContent = opts.text || '';

  var okBtn = $('#dialogOk');
  var altBtn = $('#dialogAlt');
  var cancelBtn = $('#dialogCancel');

  okBtn.textContent = opts.okText || '确定';
  if (opts.altText) {
    altBtn.hidden = false;
    altBtn.textContent = opts.altText;
  } else {
    altBtn.hidden = true;
  }
  cancelBtn.hidden = !!opts.hideCancel;

  var close = function () {
    overlay.hidden = true;
    if ($('#overlay').hidden) document.body.classList.remove('is-locked');
    okBtn.onclick = altBtn.onclick = cancelBtn.onclick = null;
  };

  okBtn.onclick = function () { close(); if (opts.onOk) opts.onOk(); };
  altBtn.onclick = function () { close(); if (opts.onAlt) opts.onAlt(); };
  cancelBtn.onclick = function () { close(); if (opts.onCancel) opts.onCancel(); };

  overlay.hidden = false;
  document.body.classList.add('is-locked');
  setTimeout(function () { okBtn.focus(); }, 30);
}

/* =========================================================================
 * 第 7 部分  启动与事件绑定
 * ========================================================================= */

var SAMPLE = [
  { name: '王小明', company: '福州大学 · 计算机学院', bookmark: true,
    phones: ['138 0000 1234', '0591-2286 6789'],
    emails: ['wangxiaoming@example.com'],
    socials: [{ label: '微信', value: 'wxm_2001' }],
    addresses: ['福建省福州市闽侯县大学城 2 号'], notes: '项目组长' },
  { name: '李思远', company: '软件工程 2 班', bookmark: true,
    phones: ['139 1111 2222'],
    emails: ['lisiyuan@example.com', 'lsy_backup@example.com'],
    socials: [{ label: 'GitHub', value: 'lisiyuan-dev' }, { label: 'QQ', value: '12345678' }],
    addresses: [], notes: '负责前端页面' },
  { name: '陈雨薇', company: '合作企业 · 产品部', bookmark: false,
    phones: ['137 6666 8888'],
    emails: ['chenyuwei@company.cn'],
    socials: [{ label: '钉钉', value: 'chenyuwei' }],
    addresses: ['福建省厦门市思明区软件园二期'], notes: '需求对接人' }
];

function seedSamples() {
  SAMPLE.forEach(function (raw) {
    var c = normalizeContact(raw);
    c.id = uid();
    c.createdAt = Date.now() + Math.floor(Math.random() * 1000);
    state.contacts.push(c);
  });
  saveContacts();
  render();
  toast('已载入 ' + SAMPLE.length + ' 条示例数据');
}

function init() {
  if (!$('#list')) return;      /* 不是主页面（例如测试页）就不初始化界面 */

  loadContacts();
  render();

  $('#addBtn').onclick = function () { openEditor(null); };
  $('#emptyAddBtn').onclick = function () { openEditor(null); };
  $('#seedBtn').onclick = seedSamples;
  $('#closeBtn').onclick = closeEditor;
  $('#cancelBtn').onclick = closeEditor;
  $('#saveBtn').onclick = saveContact;
  $('#exportBtn').onclick = exportExcel;

  $('#importBtn').onclick = function () { $('#fileInput').value = ''; $('#fileInput').click(); };
  $('#fileInput').onchange = function (e) {
    var file = e.target.files && e.target.files[0];
    if (file) importFromFile(file);
  };

  $('#searchInput').oninput = function (e) {
    state.keyword = e.target.value;
    render();
  };

  $('#tabs').onclick = function (e) {
    var btn = e.target.closest('.tab');
    if (!btn) return;
    state.filter = btn.dataset.filter;
    $$('.tab').forEach(function (t) { t.classList.toggle('is-active', t === btn); });
    render();
  };

  /* 卡片上的操作：收藏 / 编辑 / 删除（事件委托） */
  $('#list').onclick = function (e) {
    var btn = e.target.closest('[data-act]');
    if (!btn) return;
    var card = btn.closest('.card');
    var id = card && card.dataset.id;
    if (!id) return;

    var act = btn.dataset.act;

    if (act === 'star') {
      state.contacts = state.contacts.map(function (c) {
        if (c.id !== id) return c;
        return Object.assign({}, c, { bookmark: !c.bookmark });
      });
      saveContacts();
      render();
      var now = state.contacts.filter(function (c) { return c.id === id; })[0];
      toast(now.bookmark ? '已收藏，排在列表最前' : '已取消收藏');
    } else if (act === 'edit') {
      openEditor(id);
    } else if (act === 'del') {
      var target = state.contacts.filter(function (c) { return c.id === id; })[0];
      if (!target) return;
      dialog({
        title: '删除联系人',
        text: '确定要删除「' + target.name + '」吗？删除后只能重新添加。',
        okText: '删除',
        onOk: function () {
          state.contacts = state.contacts.filter(function (c) { return c.id !== id; });
          saveContacts();
          render();
          toast('已删除');
        }
      });
    }
  };

  /* 各组「+ 添加」按钮 */
  $$('.group').forEach(function (group) {
    var kind = group.dataset.kind;
    $('.add-btn', group).onclick = function () {
      var row = addFieldRow(kind, '');
      $('input', row).focus();
    };
  });

  /* Esc 关闭弹窗 */
  document.addEventListener('keydown', function (e) {
    if (e.key !== 'Escape') return;
    if (!$('#dialogOverlay').hidden) { $('#dialogCancel').click(); return; }
    if (!$('#overlay').hidden) closeEditor();
  });

  /* 点遮罩关闭 */
  $('#overlay').addEventListener('mousedown', function (e) {
    if (e.target === $('#overlay')) closeEditor();
  });
  $('#dialogOverlay').addEventListener('mousedown', function (e) {
    if (e.target === $('#dialogOverlay')) $('#dialogCancel').click();
  });
}

if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', init);
} else {
  init();
}
