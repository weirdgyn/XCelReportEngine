import argparse
import json
import re
import zipfile
from collections import Counter
from pathlib import Path
from xml.etree import ElementTree as ET


NS = {
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "rel": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "pkgrel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
}


def local_name(tag):
    return tag.rsplit("}", 1)[-1]


def xml_root(archive, name):
    try:
        return ET.fromstring(archive.read(name))
    except KeyError:
        return None


def relationship_map(archive, name):
    root = xml_root(archive, name)
    if root is None:
        return {}
    return {
        item.attrib.get("Id"): {
            "target": item.attrib.get("Target"),
            "type": item.attrib.get("Type", "").rsplit("/", 1)[-1],
            "mode": item.attrib.get("TargetMode"),
        }
        for item in root
    }


def normalized_xl_target(target):
    target = target.replace("\\", "/")
    if target.startswith("/"):
        return target.lstrip("/")
    while target.startswith("../"):
        target = target[3:]
    return "xl/" + target.lstrip("/")


def worksheet_rels_path(sheet_path):
    path = Path(sheet_path)
    return str(path.parent / "_rels" / (path.name + ".rels")).replace("\\", "/")


def drawing_summary(archive, drawing_path):
    root = xml_root(archive, drawing_path)
    if root is None:
        return None
    anchors = Counter(local_name(node.tag) for node in root if local_name(node.tag).endswith("Anchor"))
    charts = sum(1 for node in root.iter() if local_name(node.tag) == "chart")
    pictures = sum(1 for node in root.iter() if local_name(node.tag) == "pic")
    shapes = sum(1 for node in root.iter() if local_name(node.tag) == "sp")
    return {"path": drawing_path, "anchors": dict(anchors), "pictures": pictures, "charts": charts, "shapes": shapes}


def worksheet_summary(archive, sheet_path, name, state):
    root = xml_root(archive, sheet_path)
    if root is None:
        return {"name": name, "state": state, "path": sheet_path, "error": "missing worksheet part"}

    cells = root.findall(".//main:c", NS)
    formulas = root.findall(".//main:f", NS)
    formula_types = Counter(item.attrib.get("t", "normal") for item in formulas)
    formulas_without_cache = 0
    formula_samples = []
    for cell in cells:
        formula = cell.find("main:f", NS)
        if formula is not None:
            value = cell.find("main:v", NS)
            if value is None or value.text is None:
                formulas_without_cache += 1
            if len(formula_samples) < 8:
                formula_samples.append({"cell": cell.attrib.get("r"), "formula": formula.text or "", "cached": None if value is None else value.text})

    merges = root.find("main:mergeCells", NS)
    protection = root.find("main:sheetProtection", NS)
    dimension = root.find("main:dimension", NS)
    cols = root.findall(".//main:cols/main:col", NS)
    custom_rows = sum(1 for row in root.findall(".//main:sheetData/main:row", NS) if row.attrib.get("customHeight") == "1")
    rels = relationship_map(archive, worksheet_rels_path(sheet_path))
    drawings = []
    for rel in rels.values():
        if rel["type"] == "drawing" and rel["target"]:
            target = rel["target"].replace("../", "xl/") if rel["target"].startswith("../") else "xl/worksheets/" + rel["target"]
            target = str(Path(target)).replace("\\", "/")
            drawings.append(drawing_summary(archive, target))

    return {
        "name": name,
        "state": state,
        "path": sheet_path,
        "dimension": None if dimension is None else dimension.attrib.get("ref"),
        "cells": len(cells),
        "styled_cells": sum(1 for cell in cells if "s" in cell.attrib),
        "formulas": len(formulas),
        "formula_types": dict(formula_types),
        "formulas_without_cached_value": formulas_without_cache,
        "formula_samples": formula_samples,
        "merged_ranges": 0 if merges is None else int(merges.attrib.get("count", len(list(merges)))),
        "column_definitions": len(cols),
        "custom_height_rows": custom_rows,
        "sheet_protection": None if protection is None else dict(protection.attrib),
        "conditional_format_blocks": len(root.findall("main:conditionalFormatting", NS)),
        "data_validation_blocks": len(root.findall(".//main:dataValidations/main:dataValidation", NS)),
        "hyperlinks": len(root.findall(".//main:hyperlinks/main:hyperlink", NS)),
        "auto_filter": root.find("main:autoFilter", NS) is not None,
        "page_setup": root.find("main:pageSetup", NS) is not None,
        "page_margins": root.find("main:pageMargins", NS) is not None,
        "print_options": root.find("main:printOptions", NS) is not None,
        "drawing_parts": [item for item in drawings if item],
        "table_relationships": sum(1 for rel in rels.values() if rel["type"] == "table"),
        "external_relationships": sum(1 for rel in rels.values() if rel["mode"] == "External"),
    }


def styles_summary(archive):
    root = xml_root(archive, "xl/styles.xml")
    if root is None:
        return None
    result = {}
    for name in ("numFmts", "fonts", "fills", "borders", "cellStyleXfs", "cellXfs", "cellStyles", "dxfs"):
        node = root.find("main:" + name, NS)
        result[name] = 0 if node is None else int(node.attrib.get("count", len(list(node))))
    return result


def workbook_summary(path):
    with zipfile.ZipFile(path) as archive:
        names = archive.namelist()
        workbook = xml_root(archive, "xl/workbook.xml")
        rels = relationship_map(archive, "xl/_rels/workbook.xml.rels")
        sheets = []
        if workbook is not None:
            for sheet in workbook.findall(".//main:sheets/main:sheet", NS):
                rel_id = sheet.attrib.get("{%s}id" % NS["rel"])
                rel = rels.get(rel_id, {})
                target = rel.get("target")
                sheet_path = normalized_xl_target(target) if target else ""
                sheets.append(worksheet_summary(archive, sheet_path, sheet.attrib.get("name"), sheet.attrib.get("state", "visible")))

        defined_names = []
        if workbook is not None:
            for item in workbook.findall(".//main:definedNames/main:definedName", NS):
                defined_names.append({"name": item.attrib.get("name"), "localSheetId": item.attrib.get("localSheetId"), "hidden": item.attrib.get("hidden"), "value": item.text})

        calc = None if workbook is None else workbook.find("main:calcPr", NS)
        content_types = xml_root(archive, "[Content_Types].xml")
        workbook_content_type = None
        if content_types is not None:
            for override in content_types.findall("ct:Override", NS):
                if override.attrib.get("PartName") == "/xl/workbook.xml":
                    workbook_content_type = override.attrib.get("ContentType")

        media = [name for name in names if name.startswith("xl/media/") and not name.endswith("/")]
        media_types = Counter(Path(name).suffix.lower() for name in media)
        external_links = [name for name in names if re.fullmatch(r"xl/externalLinks/externalLink\d+\.xml", name)]
        charts = [name for name in names if re.fullmatch(r"xl/charts/chart\d+\.xml", name)]
        tables = [name for name in names if re.fullmatch(r"xl/tables/table\d+\.xml", name)]
        pivots = [name for name in names if name.startswith("xl/pivot") and name.endswith(".xml")]
        comments = [name for name in names if re.fullmatch(r"xl/comments\d+\.xml", name)]
        threaded_comments = [name for name in names if "threadedComments" in name]
        shared_strings_root = xml_root(archive, "xl/sharedStrings.xml")

        return {
            "path": str(path),
            "size_bytes": Path(path).stat().st_size,
            "workbook_content_type": workbook_content_type,
            "zip_parts": len(names),
            "sheets": sheets,
            "defined_names": defined_names,
            "calculation_properties": None if calc is None else dict(calc.attrib),
            "styles": styles_summary(archive),
            "shared_strings": None if shared_strings_root is None else {
                "count": int(shared_strings_root.attrib.get("count", 0)),
                "unique_count": int(shared_strings_root.attrib.get("uniqueCount", len(list(shared_strings_root)))),
            },
            "media_files": len(media),
            "media_bytes_uncompressed": sum(archive.getinfo(name).file_size for name in media),
            "media_types": dict(media_types),
            "chart_parts": len(charts),
            "table_parts": len(tables),
            "external_link_parts": len(external_links),
            "pivot_xml_parts": len(pivots),
            "comments_parts": len(comments),
            "threaded_comments_parts": len(threaded_comments),
            "has_connections": "xl/connections.xml" in names,
            "has_vba": any(name.lower().endswith("vbaproject.bin") for name in names),
            "has_calc_chain": "xl/calcChain.xml" in names,
            "has_custom_xml": any(name.startswith("customXml/") for name in names),
            "has_printer_settings": any(name.startswith("xl/printerSettings/") for name in names),
            "integrity": "valid_zip",
        }


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("files", nargs="+")
    parser.add_argument("--output")
    args = parser.parse_args()
    result = [workbook_summary(Path(item)) for item in args.files]
    text = json.dumps(result, indent=2, ensure_ascii=False)
    if args.output:
        Path(args.output).write_text(text + "\n", encoding="utf-8")
    else:
        print(text)


if __name__ == "__main__":
    main()
