"""Small OOXML helpers: inspect, surgical text-cell updates, audit workbook.

Existing result workbooks are never round-tripped through an object model.
Text updates touch named cells only. Explicit slot extensions append blank
style-only columns and their header structure; existing cell payloads and other
ZIP parts are retained. No formula recalculation is done.
"""
from __future__ import annotations

import hashlib
import os
import posixpath
import re
import shutil
import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape, quoteattr

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
Q = "{" + NS + "}"
CELL = re.compile(r"([A-Z]+)([1-9][0-9]*)\Z")


def digest(path):
    h = hashlib.sha256()
    with Path(path).open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            h.update(block)
    return h.hexdigest()


def clean_text(value):
    value = "" if value is None else str(value)
    if any(ord(c) < 32 and c not in "\t\n\r" for c in value):
        raise ValueError("文本包含 XML 不允许的控制字符")
    if len(value) > 32767:
        raise ValueError("单元格超出 Excel 32767 字符上限；完整内容应保存在 JSON")
    return value


def col_number(name):
    result = 0
    for char in name:
        result = result * 26 + ord(char) - 64
    return result


def col_name(number):
    result = ""
    while number:
        number, r = divmod(number - 1, 26)
        result = chr(r + 65) + result
    return result


def read_workbook(path):
    with zipfile.ZipFile(path) as z:
        wb = ET.fromstring(z.read("xl/workbook.xml"))
        if wb.tag != Q + "workbook":
            raise ValueError("仅支持标准 transitional XLSX；其它格式须使用可保留原格式的工具")
        rels = {r.attrib["Id"]: r.attrib["Target"] for r in
                ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))}
        strings = []
        if "xl/sharedStrings.xml" in z.namelist():
            strings = ["".join(t.text or "" for t in item.iter(Q + "t"))
                       for item in ET.fromstring(z.read("xl/sharedStrings.xml"))]
        result = {}
        for sheet in wb.find(Q + "sheets"):
            target = rels[sheet.attrib["{" + REL + "}id"]]
            part = (target.lstrip("/") if target.startswith("/") else
                    posixpath.normpath(posixpath.join("xl", target)))
            root = ET.fromstring(z.read(part))
            cells = {}
            for c in root.iter(Q + "c"):
                typ = c.get("t")
                v = c.find(Q + "v")
                raw = v.text if v is not None else None
                if typ == "s":
                    value = strings[int(raw)]
                elif typ == "inlineStr":
                    value = "".join(t.text or "" for t in c.iter(Q + "t"))
                else:
                    value = raw
                    if typ in (None, "n") and raw is not None:
                        try:
                            value = format(float(raw), ".15g")
                        except ValueError:
                            pass
                cells[c.get("r")] = {"value": value, "raw": raw,
                                      "formula": c.findtext(Q + "f"), "style": c.get("s")}
            result[sheet.get("name")] = {"part": part, "cells": cells,
                "protected": root.find(Q + "sheetProtection") is not None,
                "merged": [x.get("ref") for x in root.iter(Q + "mergeCell")]}
        return result


def merged_non_anchor(address, ranges):
    m = CELL.fullmatch(address)
    col, row = col_number(m[1]), int(m[2])
    for region in ranges:
        start, _, end = region.partition(":")
        a, b = CELL.fullmatch(start), CELL.fullmatch(end or start)
        if (col_number(a[1]) <= col <= col_number(b[1]) and
                int(a[2]) <= row <= int(b[2]) and address != start):
            return True
    return False


def patch_sheet(xml_bytes, updates):
    text = xml_bytes.decode("utf-8")
    # Keep the original namespace declarations and every unrelated byte.
    # Prefix-aware matching supports both default and prefixed SpreadsheetML.
    root_match = re.search(r"<((?:[\w]+:)?worksheet)\b", text)
    if not root_match:
        raise ValueError("无法定位 worksheet XML")
    prefix = root_match[1][:-len("worksheet")]
    ctag, rtag = prefix + "c", prefix + "row"
    for address, value in updates.items():
        value = clean_text(value)
        match = CELL.fullmatch(address)
        if not match:
            raise ValueError("非法单元格地址: " + address)
        pattern = re.compile(r"<" + re.escape(ctag) + r"\b(?=[^>]*\br=" +
            quoteattr(address) + r")(?:[^>]*?/>|[^>]*?>.*?</" + re.escape(ctag) + r">)", re.S)
        existing = pattern.search(text)
        attrs = ' r="' + address + '"'
        if existing:
            opening = re.match(r"<[^\s>]+(.*?)(?:/?>)", existing[0], re.S)[1]
            attrs = re.sub(r'\s+t="[^"]*"', "", opening).rstrip("/ ")
        element = ("<" + ctag + attrs + ' t="inlineStr"><' + prefix +
                   'is><' + prefix + 't xml:space="preserve">' + escape(value) +
                   "</" + prefix + "t></" + prefix + "is></" + ctag + ">")
        if existing:
            text = text[:existing.start()] + element + text[existing.end():]
        else:
            row_pattern = re.compile(r"<" + re.escape(rtag) + r"\b(?=[^>]*\br=\"" +
                match[2] + r"\")[^>]*>.*?</" + re.escape(rtag) + ">", re.S)
            row_match = row_pattern.search(text)
            if not row_match:
                raise ValueError("目标行不存在，不能推测表结构: " + address)
            row_text = row_match[0]
            insert = row_text.rfind("</" + rtag + ">")
            for c in re.finditer(r"<" + re.escape(ctag) + r'\b[^>]*\br="([A-Z]+)[0-9]+"', row_text):
                if col_number(c[1]) > col_number(match[1]):
                    insert = c.start()
                    break
            row_text = row_text[:insert] + element + row_text[insert:]
            text = text[:row_match.start()] + row_text + text[row_match.end():]
    ET.fromstring(text)  # structural check before saving
    return text.encode("utf-8")


def extend_sheet(xml_bytes, source_columns, target_columns):
    """Append an empty paired object slot, copying styles/widths, not results."""
    text = xml_bytes.decode("utf-8")
    root = ET.fromstring(text)
    source = [col_number(c) for c in source_columns]
    target = [col_number(c) for c in target_columns]
    occupied = [col_number(CELL.fullmatch(c.get("r"))[1]) for c in root.iter(Q + "c")]
    occupied += [col_number(CELL.fullmatch(m.get("ref").split(":")[-1])[1]) for m in root.iter(Q + "mergeCell")]
    if len(source) != 2 or source[1] != source[0] + 1 or len(target) != 2 or target[1] != target[0] + 1:
        raise ValueError("对象列组必须是相邻两列")
    if target[0] <= max(occupied, default=0) or target[-1] > 16384:
        raise ValueError("只能在现有列组右侧空白处扩展，不能覆盖已有数据或超出Excel列数")
    if root.find(Q + "tableParts") is not None:
        raise ValueError("原生Excel Table布局须使用可保留该功能的工具扩展，不直接改列")
    if root.find(Q + "sheetProtection") is not None:
        raise ValueError("不能扩展受保护工作表")
    prefix = re.search(r"<((?:[\w]+:)?worksheet)\b", text)[1][:-len("worksheet")]
    ctag, rtag = prefix + "c", prefix + "row"
    def append_cells(match):
        row = match[0]
        row_number = re.search(r'\br="([0-9]+)"', row)[1]
        cells = []
        for s, t in zip(source_columns, target_columns):
            old = re.search(r"<" + re.escape(ctag) + r'\b(?=[^>]*\br="' + s + row_number + r'")[^>]*?/?>', row)
            if old:
                style = re.search(r'\bs="([0-9]+)"', old[0])
                cells.append('<' + ctag + ' r="' + t + row_number + '"' + (' s="' + style[1] + '"' if style else '') + '/>')
            else:
                cells.append('<' + ctag + ' r="' + t + row_number + '"/>')
        insert = row.rfind('</' + rtag + '>')
        ext = row.find('<' + prefix + 'extLst')
        if ext >= 0:
            insert = min(insert, ext)
        row = row[:insert] + ''.join(cells) + row[insert:]
        def spans(m):
            return 'spans="' + m[1] + ':' + str(max(int(m[2]), target[-1])) + '"'
        return re.sub(r'spans="([0-9]+):([0-9]+)"', spans, row, count=1)
    text = re.sub(r"<" + re.escape(rtag) + r"\b[^>]*>.*?</" + re.escape(rtag) + ">", append_cells, text, flags=re.S)
    # Column definitions may cover a range, so split only ranges intersecting
    # the new pair rather than introducing overlapping <col> definitions.
    columns = root.find(Q + "cols")
    definitions = [dict(c.attrib) for c in columns] if columns is not None else []
    original_cols = re.search(r"<" + prefix + r"cols\b(?:[^>]*?/>|[^>]*?>.*?</" + prefix + "cols>)", text, re.S)
    col_elements = []
    for attrs in definitions:
        lo, hi = int(attrs["min"]), int(attrs["max"])
        intervals = [(lo, hi)] if hi < target[0] or lo > target[-1] else [(lo, target[0] - 1), (target[-1] + 1, hi)]
        for a, b in intervals:
            if a <= b:
                col_elements.append({**attrs, "min": str(a), "max": str(b)})
    for s, t in zip(source, target):
        attrs = next((d for d in definitions if int(d["min"]) <= s <= int(d["max"])), {})
        col_elements.append({**attrs, "min": str(t), "max": str(t)})
    col_elements.sort(key=lambda a: int(a["min"]))
    cols_xml = '<' + prefix + 'cols>' + ''.join('<' + prefix + 'col ' + ' '.join(k + '=' + quoteattr(v) for k, v in attrs.items()) + '/>' for attrs in col_elements) + '</' + prefix + 'cols>'
    if original_cols:
        text = text[:original_cols.start()] + cols_xml + text[original_cols.end():]
    else:
        data_pos = text.index('<' + prefix + 'sheetData')
        text = text[:data_pos] + cols_xml + text[data_pos:]
    # Clone only merges wholly contained in the source pair. A broad page
    # title crossing existing groups is unrelated and stays exactly as-is.
    new_merges = []
    offset = target[0] - source[0]
    for merge in root.iter(Q + "mergeCell"):
        a, _, b = merge.get("ref").partition(":")
        left, right = CELL.fullmatch(a), CELL.fullmatch(b or a)
        if source[0] <= col_number(left[1]) <= col_number(right[1]) <= source[-1]:
            new_merges.append(col_name(col_number(left[1]) + offset) + left[2] + ':' + col_name(col_number(right[1]) + offset) + right[2])
    if new_merges:
        xml = ''.join('<' + prefix + 'mergeCell ref=' + quoteattr(ref) + '/>' for ref in new_merges)
        merges = re.search('<' + prefix + r'mergeCells\b[^>]*>.*?</' + prefix + 'mergeCells>', text, re.S)
        if merges:
            updated = merges[0].replace('</' + prefix + 'mergeCells>', xml + '</' + prefix + 'mergeCells>')
            count = len(list(root.iter(Q + "mergeCell"))) + len(new_merges)
            updated = re.sub(r'\bcount="[0-9]+"', 'count="' + str(count) + '"', updated, count=1)
            text = text[:merges.start()] + updated + text[merges.end():]
        else:
            # These elements follow mergeCells in the worksheet content order.
            later = r'phoneticPr|conditionalFormatting|dataValidations|hyperlinks|printOptions|pageMargins|pageSetup|headerFooter|rowBreaks|colBreaks|customProperties|cellWatches|ignoredErrors|smartTags|drawing|legacyDrawing|legacyDrawingHF|picture|oleObjects|controls|webPublishItems|tableParts|extLst'
            after_data = text.index('</' + prefix + 'sheetData>') + len('</' + prefix + 'sheetData>')
            found = re.search('<' + prefix + '(?:' + later + r')\b', text[after_data:])
            pos = after_data + found.start() if found else text.rfind('</' + prefix + 'worksheet>')
            text = text[:pos] + '<' + prefix + 'mergeCells count="' + str(len(new_merges)) + '">' + xml + '</' + prefix + 'mergeCells>' + text[pos:]
    def dimension(m):
        ref = m[1]
        a, sep, b = ref.partition(':')
        end = CELL.fullmatch(b or a)
        return m[0].replace(ref, a + ':' + col_name(max(target[-1], col_number(end[1]))) + end[2])
    text = re.sub('<' + prefix + r'dimension\b[^>]*\bref="([A-Z]+[0-9]+(?::[A-Z]+[0-9]+)?)"[^>]*/>', dimension, text, count=1)
    ET.fromstring(text)
    return text.encode("utf-8")


def update_workbook(path, changes, backup, extensions=None):
    """changes: [{sheet, cell, expected, value}]. Reject stale/formula edits."""
    path, backup = Path(path), Path(backup)
    before = digest(path)
    sheets = read_workbook(path)
    by_part = {}
    seen = set()
    for change in changes:
        key = (change["sheet"], change["cell"])
        if key in seen:
            raise ValueError("重复目标单元格: " + str(key))
        seen.add(key)
        sheet = sheets[change["sheet"]]
        if sheet["protected"]:
            raise ValueError("目标工作表有保护，不能绕过保护写入")
        cell = change["cell"]
        if not CELL.fullmatch(cell) or merged_non_anchor(cell, sheet["merged"]):
            raise ValueError("非法地址或合并区域非左上角: " + cell)
        old = sheet["cells"].get(cell, {}).get("value")
        if (old or "") != (change["expected"] or ""):
            raise ValueError("Excel 已变化，需重新读取: " + cell)
        if sheet["cells"].get(cell, {}).get("formula") is not None:
            raise ValueError("不能覆盖公式单元格: " + cell)
        if (old or "") != clean_text(change["value"]):
            by_part.setdefault(sheet["part"], {})[cell] = change["value"]
    for extension in extensions or []:
        sheet = sheets[extension["sheet"]]
        if sheet["protected"]:
            raise ValueError("不能扩展受保护工作表")
        by_part.setdefault(sheet["part"], {})
    if not by_part:
        return {"status": "unchanged", "sha256": before, "cells": []}
    # Validate every part before touching the workbook.
    with zipfile.ZipFile(path) as z:
        payloads = {part: z.read(part) for part in by_part}
        for extension in extensions or []:
            part = sheets[extension["sheet"]]["part"]
            payloads[part] = extend_sheet(payloads[part], extension["source_columns"], extension["target_columns"])
        payloads = {part: patch_sheet(payloads[part], values) for part, values in by_part.items()}
        backup.parent.mkdir(parents=True, exist_ok=True)
        if backup.exists():
            raise FileExistsError("备份已存在，不覆盖: " + str(backup))
        shutil.copy2(path, backup)
        if digest(backup) != before:
            raise ValueError("备份校验失败")
        fd, tmp = tempfile.mkstemp(prefix=".delivery-", suffix=".xlsx", dir=path.parent)
        os.close(fd)
        try:
            with zipfile.ZipFile(tmp, "w") as dest:
                dest.comment = z.comment
                for info in z.infolist():
                    dest.writestr(info, payloads.get(info.filename, z.read(info.filename)))
            checked = read_workbook(tmp)
            for change in changes:
                if (checked[change["sheet"]]["cells"].get(change["cell"], {}).get("value") or "") != (change["value"] or ""):
                    raise ValueError("写入后单元格核验失败")
            if digest(path) != before:
                raise ValueError("Excel 同时被其他进程修改，取消更新")
        except Exception:
            Path(tmp).unlink(missing_ok=True)
            raise
    try:
        os.replace(tmp, path)  # closed ZIP needed on Windows
    finally:
        Path(tmp).unlink(missing_ok=True)
    return {"status": "updated", "backup": str(backup), "sha256": digest(path),
            "cells": [dict(c) for c in changes]}


def write_audit(path, tables):
    """Create a simple text-only XLSX; every string is inline, never a formula."""
    types = ['<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
             '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>']
    sheets, relationships = [], []
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("_rels/.rels", '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="' + REL + '/officeDocument" Target="xl/workbook.xml"/></Relationships>')
        z.writestr("xl/styles.xml", '<styleSheet xmlns="' + NS + '"><fonts count="2"><font><sz val="11"/><name val="Microsoft YaHei"/></font><font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Microsoft YaHei"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF24476B"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf/></cellStyleXfs><cellXfs count="2"><xf fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf><xf fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>')
        for i, (name, rows) in enumerate(tables.items(), 1):
            if len(name) > 31 or re.search(r'[\\/?*\[\]:]', name):
                raise ValueError("非法工作表名")
            types.append('<Override PartName="/xl/worksheets/sheet%d.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' % i)
            sheets.append('<sheet name=' + quoteattr(name) + ' sheetId="%d" r:id="rId%d"/>' % (i, i))
            relationships.append('<Relationship Id="rId%d" Type="%s/worksheet" Target="worksheets/sheet%d.xml"/>' % (i, REL, i))
            row_xml = []
            for r, row in enumerate(rows, 1):
                cells = []
                for c, value in enumerate(row, 1):
                    value = "" if value is None else str(value)
                    if len(value) > 32767:
                        value = value[:32700] + "\n[完整内容见同目录审计 JSON]"
                    cells.append('<c r="%s%d" t="inlineStr" s="%d"><is><t xml:space="preserve">%s</t></is></c>' % (col_name(c), r, int(r == 1), escape(clean_text(value))))
                row_xml.append('<row r="%d">%s</row>' % (r, "".join(cells)))
            count = max((len(row) for row in rows), default=1)
            z.writestr("xl/worksheets/sheet%d.xml" % i, '<worksheet xmlns="' + NS + '"><sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><cols>' + ''.join('<col min="%d" max="%d" width="%d" customWidth="1"/>' % (c, c, 18 if c <= 2 else 38) for c in range(1, count + 1)) + '</cols><sheetData>' + ''.join(row_xml) + '</sheetData><autoFilter ref="A1:%s%d"/></worksheet>' % (col_name(count), max(1, len(rows))))
        relationships.append('<Relationship Id="styles" Type="' + REL + '/styles" Target="styles.xml"/>')
        z.writestr("xl/workbook.xml", '<workbook xmlns="' + NS + '" xmlns:r="' + REL + '"><sheets>' + ''.join(sheets) + '</sheets></workbook>')
        z.writestr("xl/_rels/workbook.xml.rels", '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + ''.join(relationships) + '</Relationships>')
        z.writestr("[Content_Types].xml", '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' + ''.join(types) + '</Types>')
