from __future__ import annotations

import base64
import hashlib
import re
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, List, Tuple

from core.ooxml_namespaces import (
    CT_NS,
    PKG_REL_NS,
    serialize_content_types,
    serialize_package_relationships,
)
from core.section_mapping import choose_section_sources
from core.sectpr_tools import (
    canonical_sectpr_order_index,
    child_tag_name,
    extract_all_sectpr_blocks,
    extract_sectpr_children,
    replace_nth_sectpr_block,
    strip_tag_block,
)


@dataclass
class HeaderFooterImportResult:
    part_names: set[str] = field(default_factory=set)
    rels_names: set[str] = field(default_factory=set)
    media_names: set[str] = field(default_factory=set)


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


def _allocate_unique_media_name(part_name: str, index: int, original_name: str, payload: bytes, used: set[str]) -> str:
    stem = Path(part_name).stem
    suffix = hashlib.sha1(payload).hexdigest()[:8]
    ext = Path(original_name).suffix.lower() or ".bin"
    candidate = f"hf_{stem}_{index:02d}_{suffix}{ext}"
    n = 1
    while candidate in used:
        candidate = f"hf_{stem}_{index:02d}_{suffix}_{n}{ext}"
        n += 1
    used.add(candidate)
    return candidate


def _normalize_rel_target(target: str) -> str:
    return target.strip().lstrip("./")


def _write_hf_parts(
    target_extract_dir: Path,
    entries: List[Tuple[str, Dict[str, Any]]],
    log: List[str],
    result: HeaderFooterImportResult,
) -> Dict[str, str]:
    part_to_type: Dict[str, str] = {}
    media_out = target_extract_dir / "word" / "media"
    media_out.mkdir(parents=True, exist_ok=True)
    written_media: set[str] = {p.name for p in media_out.iterdir() if p.is_file()} if media_out.exists() else set()

    for kind, entry in entries:
        part_name = entry.get("part_name")
        xml_content = entry.get("xml") or entry.get("part_xml")
        if not isinstance(part_name, str) or not isinstance(xml_content, str):
            log.append(f"WARNING: Skipping malformed {kind} entry (missing part_name/xml)")
            continue

        part_path = target_extract_dir / part_name
        part_path.parent.mkdir(parents=True, exist_ok=True)
        part_path.write_text(xml_content, encoding="utf-8")
        result.part_names.add(part_name)
        part_to_type[part_name] = kind
        log.append(f"Wrote {kind} part: {part_name}")

        rels_xml = entry.get("rels_xml") or entry.get("relationships_xml")
        rels_name = entry.get("rels_part_name")
        target_map: Dict[str, str] = {}

        media_items = _resolve_media_items(entry)
        for idx, media_item in enumerate(media_items, start=1):
            filename = _resolve_media_filename(media_item)
            payload = _resolve_media_bytes(media_item)
            if not filename or payload is None:
                continue
            new_name = _allocate_unique_media_name(part_name, idx, filename, payload, written_media)
            out_rel = f"media/{new_name}"
            target_map[_normalize_rel_target(f"media/{filename}")] = out_rel
            target_map[_normalize_rel_target(f"./media/{filename}")] = out_rel
            (media_out / new_name).write_bytes(payload)
            result.media_names.add(f"word/media/{new_name}")
            log.append(f"Wrote media asset: word/media/{new_name}")

        if isinstance(rels_xml, str):
            if not isinstance(rels_name, str) or not rels_name:
                rels_name = f"word/_rels/{Path(part_name).name}.rels"
            if target_map:
                rels_root = ET.fromstring(rels_xml.encode("utf-8"))
                for rel in rels_root.findall(f"{{{PKG_REL_NS}}}Relationship"):
                    old = _normalize_rel_target(rel.attrib.get("Target", ""))
                    if old in target_map:
                        rel.set("Target", target_map[old])
                rels_xml = serialize_package_relationships(rels_root).decode("utf-8")
            rels_path = target_extract_dir / rels_name
            rels_path.parent.mkdir(parents=True, exist_ok=True)
            rels_path.write_text(rels_xml, encoding="utf-8")
            result.rels_names.add(rels_name)
            log.append(f"Wrote rels part: {rels_name}")

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

    rels_path.write_bytes(serialize_package_relationships(root))
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


def _raw_ref(kind: str, ref_type: str, rid: str) -> str:
    return f'<w:{kind}Reference w:type="{ref_type}" r:id="{rid}"/>'


def _rewire_document_sectpr(target_extract_dir: Path, registry: Dict[str, Any], entries: List[Tuple[str, Dict[str, Any]]], part_to_rid: Dict[str, str], log: List[str]) -> None:
    doc_path = target_extract_dir / "word" / "document.xml"
    if not doc_path.exists():
        return

    page_layout = registry.get("page_layout", {}) if isinstance(registry, dict) else {}
    rid_to_part = _build_arch_rid_to_part(entries)

    doc_original = doc_path.read_text(encoding="utf-8")
    sectprs = extract_all_sectpr_blocks(doc_original)
    if not sectprs:
        log.append("No sectPr blocks found; skipped header/footer rewiring")
        return

    section_sources = choose_section_sources(len(sectprs), page_layout, require_default=True, log=log)
    updated_xml = doc_original
    order_index = canonical_sectpr_order_index()

    for idx, (target_sectpr, source) in enumerate(zip(sectprs, section_sources)):
        headers, footers = _extract_arch_hf_refs(source)

        open_tag_m = re.match(r'(<w:sectPr\b[^>]*>)', target_sectpr)
        close_tag = "</w:sectPr>"
        if not open_tag_m or not target_sectpr.endswith(close_tag):
            continue
        open_tag = open_tag_m.group(1)
        inner = target_sectpr[len(open_tag):-len(close_tag)]

        for tag in ("headerReference", "footerReference", "titlePg"):
            inner = strip_tag_block(inner, tag)
        children = extract_sectpr_children(inner)

        insert_nodes: List[str] = []
        has_first = False
        for ref_type, old_rid in headers.items():
            part_name = rid_to_part.get(old_rid)
            new_rid = part_to_rid.get(part_name) if part_name else None
            if not new_rid:
                continue
            insert_nodes.append(_raw_ref("header", ref_type, new_rid))
            has_first = has_first or ref_type == "first"

        for ref_type, old_rid in footers.items():
            part_name = rid_to_part.get(old_rid)
            new_rid = part_to_rid.get(part_name) if part_name else None
            if not new_rid:
                continue
            insert_nodes.append(_raw_ref("footer", ref_type, new_rid))
            has_first = has_first or ref_type == "first"

        if has_first:
            insert_nodes.append("<w:titlePg/>")

        for node in insert_nodes:
            node_tag = child_tag_name(node)
            node_order = order_index.get(node_tag or "", 10_000)
            insert_at = len(children)
            for i, child in enumerate(children):
                ctag = child_tag_name(child)
                if order_index.get(ctag or "", 10_000) > node_order:
                    insert_at = i
                    break
            children.insert(insert_at, node)

        updated_sectpr = f"{open_tag}{''.join(children)}{close_tag}"
        updated_xml = replace_nth_sectpr_block(updated_xml, idx, updated_sectpr)

    ET.fromstring(updated_xml.encode("utf-8"))
    doc_path.write_text(updated_xml, encoding="utf-8")
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

    ct_path.write_bytes(serialize_content_types(root))
    log.append("Updated [Content_Types].xml for header/footer parts and media")


def import_headers_footers(target_extract_dir: Path, registry: Dict[str, Any], log: List[str]) -> HeaderFooterImportResult:
    result = HeaderFooterImportResult()
    entries = _iter_hf_entries(registry)
    if not entries:
        log.append("No architect headers/footers in registry; skipping import")
        return result

    _remove_existing_hf_files(target_extract_dir, log)
    part_to_type = _write_hf_parts(target_extract_dir, entries, log, result)
    part_to_rid = _rebuild_document_rels(target_extract_dir, part_to_type, log)
    _rewire_document_sectpr(target_extract_dir, registry, entries, part_to_rid, log)
    _ensure_content_types(target_extract_dir, part_to_type, entries, log)
    return result
