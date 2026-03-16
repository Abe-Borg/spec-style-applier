from __future__ import annotations

import base64
import re
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Any, Dict, List, Tuple

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"


def _iter_hf_entries(registry: Dict[str, Any]) -> List[Tuple[str, Dict[str, Any]]]:
    hf_data = registry.get("headers_footers", {}) if isinstance(registry, dict) else {}
    headers = hf_data.get("headers", []) if isinstance(hf_data, dict) else []
    footers = hf_data.get("footers", []) if isinstance(hf_data, dict) else []
    entries: List[Tuple[str, Dict[str, Any]]] = []
    entries.extend(("header", e) for e in headers if isinstance(e, dict))
    entries.extend(("footer", e) for e in footers if isinstance(e, dict))
    return entries


def _resolve_media_items(entry: Dict[str, Any]) -> List[Dict[str, Any]]:
    media = entry.get("media") or entry.get("media_files") or []
    return [m for m in media if isinstance(m, dict)]


def _resolve_media_filename(media_item: Dict[str, Any]) -> str | None:
    for key in ("path", "target", "name", "filename"):
        val = media_item.get(key)
        if isinstance(val, str) and val.strip():
            path = val.strip()
            return path.split("/")[-1]
    return None


def _resolve_media_bytes(media_item: Dict[str, Any]) -> bytes | None:
    for key in ("content_base64", "base64", "data", "data_base64"):
        val = media_item.get(key)
        if isinstance(val, str) and val:
            return base64.b64decode(val)
    return None


def _remove_existing_hf_files(target_extract_dir: Path, log: List[str]) -> None:
    word_dir = target_extract_dir / "word"
    rels_dir = word_dir / "_rels"

    for pattern in ("header*.xml", "footer*.xml"):
        for path in word_dir.glob(pattern):
            path.unlink(missing_ok=True)
            log.append(f"Removed old part: word/{path.name}")

    for pattern in ("header*.xml.rels", "footer*.xml.rels"):
        for path in rels_dir.glob(pattern):
            path.unlink(missing_ok=True)
            log.append(f"Removed old rels: word/_rels/{path.name}")


def _write_hf_parts(target_extract_dir: Path, entries: List[Tuple[str, Dict[str, Any]]], log: List[str]) -> Dict[str, str]:
    part_to_type: Dict[str, str] = {}
    media_out = target_extract_dir / "word" / "media"
    written_media: set[str] = set()

    for kind, entry in entries:
        part_name = entry.get("part_name")
        xml_content = entry.get("xml") or entry.get("part_xml")
        if not isinstance(part_name, str) or not isinstance(xml_content, str):
            log.append(f"WARNING: Skipping malformed {kind} entry (missing part_name/xml)")
            continue

        part_path = target_extract_dir / part_name
        part_path.parent.mkdir(parents=True, exist_ok=True)
        part_path.write_text(xml_content, encoding="utf-8")
        part_to_type[part_name] = kind
        log.append(f"Wrote {kind} part: {part_name}")

        rels_xml = entry.get("rels_xml") or entry.get("relationships_xml")
        rels_name = entry.get("rels_part_name")
        if isinstance(rels_xml, str):
            if not isinstance(rels_name, str) or not rels_name:
                rels_name = f"word/_rels/{Path(part_name).name}.rels"
            rels_path = target_extract_dir / rels_name
            rels_path.parent.mkdir(parents=True, exist_ok=True)
            rels_path.write_text(rels_xml, encoding="utf-8")
            log.append(f"Wrote rels part: {rels_name}")

        for media_item in _resolve_media_items(entry):
            filename = _resolve_media_filename(media_item)
            payload = _resolve_media_bytes(media_item)
            if not filename or payload is None:
                continue
            media_out.mkdir(parents=True, exist_ok=True)
            if filename in written_media:
                continue
            (media_out / filename).write_bytes(payload)
            written_media.add(filename)
            log.append(f"Wrote media asset: word/media/{filename}")

    return part_to_type


def _next_rid(rels_root: ET.Element) -> int:
    rids = []
    for rel in rels_root.findall(f"{{{PKG_REL_NS}}}Relationship"):
        rid = rel.attrib.get("Id", "")
        m = re.fullmatch(r"rId(\d+)", rid)
        if m:
            rids.append(int(m.group(1)))
    return (max(rids) if rids else 0) + 1


def _rebuild_document_rels(target_extract_dir: Path, part_to_type: Dict[str, str], log: List[str]) -> Dict[str, str]:
    rels_path = target_extract_dir / "word" / "_rels" / "document.xml.rels"
    if not rels_path.exists():
        raise FileNotFoundError(f"Missing required file: {rels_path}")

    root = ET.fromstring(rels_path.read_bytes())
    for rel in list(root.findall(f"{{{PKG_REL_NS}}}Relationship")):
        rel_type = rel.attrib.get("Type", "")
        if rel_type.endswith("/header") or rel_type.endswith("/footer"):
            root.remove(rel)

    part_to_rid: Dict[str, str] = {}
    rid_num = _next_rid(root)
    for part_name, kind in sorted(part_to_type.items()):
        rid = f"rId{rid_num}"
        rid_num += 1
        target = Path(part_name).name
        rel_type = f"http://schemas.openxmlformats.org/officeDocument/2006/relationships/{kind}"
        ET.SubElement(root, f"{{{PKG_REL_NS}}}Relationship", {"Id": rid, "Type": rel_type, "Target": target})
        part_to_rid[part_name] = rid

    rels_path.write_bytes(ET.tostring(root, encoding="utf-8", xml_declaration=True))
    log.append(f"Rebuilt document.xml.rels header/footer relationships ({len(part_to_rid)} entries)")
    return part_to_rid


def _extract_arch_hf_refs(page_layout_section: Dict[str, Any]) -> Tuple[Dict[str, str], Dict[str, str]]:
    headers: Dict[str, str] = {}
    footers: Dict[str, str] = {}

    for key in ("header_refs", "headers"):
        val = page_layout_section.get(key)
        if isinstance(val, dict):
            headers = {k: v for k, v in val.items() if isinstance(k, str) and isinstance(v, str)}
            break
    for key in ("footer_refs", "footers"):
        val = page_layout_section.get(key)
        if isinstance(val, dict):
            footers = {k: v for k, v in val.items() if isinstance(k, str) and isinstance(v, str)}
            break

    return headers, footers


def _build_arch_rid_to_part(entries: List[Tuple[str, Dict[str, Any]]]) -> Dict[str, str]:
    out: Dict[str, str] = {}
    for _kind, entry in entries:
        part_name = entry.get("part_name")
        if not isinstance(part_name, str):
            continue
        for key in ("rid", "rId", "relationship_id"):
            rid = entry.get(key)
            if isinstance(rid, str):
                out[rid] = part_name
    return out


def _rewire_document_sectpr(target_extract_dir: Path, registry: Dict[str, Any], entries: List[Tuple[str, Dict[str, Any]]], part_to_rid: Dict[str, str], log: List[str]) -> None:
    doc_path = target_extract_dir / "word" / "document.xml"
    if not doc_path.exists():
        return

    page_layout = registry.get("page_layout", {}) if isinstance(registry, dict) else {}
    section_chain = page_layout.get("section_chain", []) if isinstance(page_layout, dict) else []
    section_chain = [s for s in section_chain if isinstance(s, dict)]

    rid_to_part = _build_arch_rid_to_part(entries)

    root = ET.fromstring(doc_path.read_bytes())
    sectprs = root.findall(f".//{{{W_NS}}}sectPr")
    if not sectprs:
        log.append("No sectPr blocks found; skipped header/footer rewiring")
        return

    use_one_to_one = len(section_chain) == len(sectprs) and len(section_chain) > 0
    fallback_section = section_chain[-1] if section_chain else {}

    for idx, sectpr in enumerate(sectprs):
        for tag in ("headerReference", "footerReference"):
            for node in list(sectpr.findall(f"{{{W_NS}}}{tag}")):
                sectpr.remove(node)

        source = section_chain[idx] if use_one_to_one else fallback_section
        headers, footers = _extract_arch_hf_refs(source)

        insert_nodes: List[ET.Element] = []
        has_first = False
        for ref_type, old_rid in headers.items():
            part_name = rid_to_part.get(old_rid)
            new_rid = part_to_rid.get(part_name) if part_name else None
            if not new_rid:
                continue
            node = ET.Element(f"{{{W_NS}}}headerReference")
            node.attrib[f"{{{W_NS}}}type"] = ref_type
            node.attrib[f"{{{R_NS}}}id"] = new_rid
            insert_nodes.append(node)
            has_first = has_first or ref_type == "first"

        for ref_type, old_rid in footers.items():
            part_name = rid_to_part.get(old_rid)
            new_rid = part_to_rid.get(part_name) if part_name else None
            if not new_rid:
                continue
            node = ET.Element(f"{{{W_NS}}}footerReference")
            node.attrib[f"{{{W_NS}}}type"] = ref_type
            node.attrib[f"{{{R_NS}}}id"] = new_rid
            insert_nodes.append(node)
            has_first = has_first or ref_type == "first"

        for node in reversed(insert_nodes):
            sectpr.insert(0, node)

        if has_first and sectpr.find(f"{{{W_NS}}}titlePg") is None:
            anchor = len(insert_nodes)
            sectpr.insert(anchor, ET.Element(f"{{{W_NS}}}titlePg"))

    doc_path.write_bytes(ET.tostring(root, encoding="utf-8", xml_declaration=True))
    log.append(f"Rewired sectPr header/footer references in {len(sectprs)} sections")


def _ensure_content_types(target_extract_dir: Path, part_to_type: Dict[str, str], entries: List[Tuple[str, Dict[str, Any]]], log: List[str]) -> None:
    ct_path = target_extract_dir / "[Content_Types].xml"
    if not ct_path.exists():
        return

    root = ET.fromstring(ct_path.read_bytes())
    existing_overrides = {
        node.attrib.get("PartName")
        for node in root.findall(f"{{{CT_NS}}}Override")
    }
    existing_defaults = {
        (node.attrib.get("Extension", "").lower(), node.attrib.get("ContentType", ""))
        for node in root.findall(f"{{{CT_NS}}}Default")
    }

    for part_name, kind in sorted(part_to_type.items()):
        part_uri = f"/{part_name}"
        if part_uri in existing_overrides:
            continue
        content_type = (
            "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"
            if kind == "header"
            else "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"
        )
        ET.SubElement(root, f"{{{CT_NS}}}Override", {"PartName": part_uri, "ContentType": content_type})

    ext_to_type = {
        "png": "image/png",
        "jpg": "image/jpeg",
        "jpeg": "image/jpeg",
        "gif": "image/gif",
        "bmp": "image/bmp",
        "tif": "image/tiff",
        "tiff": "image/tiff",
        "wmf": "image/x-wmf",
        "emf": "image/x-emf",
    }
    for _kind, entry in entries:
        for media in _resolve_media_items(entry):
            filename = _resolve_media_filename(media)
            if not filename or "." not in filename:
                continue
            ext = filename.rsplit(".", 1)[-1].lower()
            content_type = ext_to_type.get(ext)
            if not content_type:
                continue
            marker = (ext, content_type)
            if marker in existing_defaults:
                continue
            ET.SubElement(root, f"{{{CT_NS}}}Default", {"Extension": ext, "ContentType": content_type})
            existing_defaults.add(marker)

    ct_path.write_bytes(ET.tostring(root, encoding="utf-8", xml_declaration=True))
    log.append("Updated [Content_Types].xml for header/footer parts and media")


def import_headers_footers(target_extract_dir: Path, registry: Dict[str, Any], log: List[str]) -> None:
    entries = _iter_hf_entries(registry)
    if not entries:
        log.append("No architect headers/footers in registry; skipping import")
        return

    _remove_existing_hf_files(target_extract_dir, log)
    part_to_type = _write_hf_parts(target_extract_dir, entries, log)
    part_to_rid = _rebuild_document_rels(target_extract_dir, part_to_type, log)
    _rewire_document_sectpr(target_extract_dir, registry, entries, part_to_rid, log)
    _ensure_content_types(target_extract_dir, part_to_type, entries, log)
