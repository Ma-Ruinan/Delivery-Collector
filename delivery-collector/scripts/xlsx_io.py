"""Small OOXML helpers: inspect, surgical text-cell updates, audit workbook.

Existing result workbooks are never round-tripped through an object model.
Only named cell elements in named sheet parts change; other ZIP payloads and
all unrelated worksheet bytes stay intact. No formula recalculation is done.
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


def update_workbook(path, changes, backup):
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
    if not by_part:
        return {"status": "unchanged", "sha256": before, "cells": []}
    # Validate every part before touching the workbook.
    with zipfile.ZipFile(path) as z:
        payloads = {part: patch_sheet(z.read(part), values) for part, values in by_part.items()}
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
