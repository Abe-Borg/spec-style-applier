# docx_patch.py
from __future__ import annotations

from pathlib import Path
import zipfile
import xml.etree.ElementTree as ET
from typing import Dict, List, Union

BytesOrStr = Union[bytes, str]


def validate_xml_wellformedness(replacements: Dict[str, bytes]) -> List[str]:
    """Parse each replacement part with ElementTree to verify XML well-formedness.

    Returns a list of error strings (empty means all parts are valid).
    """
    errors: List[str] = []
    for name, content in replacements.items():
        try:
            ET.fromstring(content)
        except ET.ParseError as exc:
            errors.append(f"{name}: XML parse error: {exc}")
    return errors


def patch_docx(
    src_docx: Path,
    out_docx: Path,
    replacements: Dict[str, BytesOrStr],
    sync_mode: str = "body_only",
) -> None:
    """
    Create out_docx by copying every ZIP entry from src_docx unchanged,
    except for entries whose internal paths match keys in `replacements`.

    This is NOT a "rebuild from extracted folder".
    It's a surgical patch: swap specific parts, preserve everything else.
    """
    src_docx = Path(src_docx)
    out_docx = Path(out_docx)

    rep_bytes: Dict[str, bytes] = {}
    for k, v in replacements.items():
        if isinstance(v, str):
            rep_bytes[k] = v.encode("utf-8")
        else:
            rep_bytes[k] = v

    # Phase 2 hard invariants — enforce at patch boundary
    ALLOWED_PATCHES_BODY_ONLY = {
        "word/document.xml",
        "word/styles.xml",
        "word/theme/theme1.xml",
        "word/numbering.xml",
        "word/settings.xml",
        "word/fontTable.xml",
        "[Content_Types].xml",
        "word/_rels/document.xml.rels",
    }

    FORBIDDEN_PREFIXES_BODY_ONLY = (
        "word/header",
        "word/footer",
    )

    if sync_mode == "template_sync":
        allowed = set(ALLOWED_PATCHES_BODY_ONLY)
        forbidden_prefixes: tuple = ()
    else:
        allowed = ALLOWED_PATCHES_BODY_ONLY
        forbidden_prefixes = FORBIDDEN_PREFIXES_BODY_ONLY

    for name in rep_bytes:
        if forbidden_prefixes and name.startswith(forbidden_prefixes):
            raise RuntimeError(f"Forbidden patch target in {sync_mode} mode: {name}")

        # In template_sync mode, allow header/footer patches through prefix check
        if name not in allowed:
            if sync_mode == "template_sync" and (
                name.startswith("word/header") or name.startswith("word/footer")
            ):
                pass  # allowed in template_sync
            else:
                raise RuntimeError(
                    f"Illegal patch target: {name}\n"
                    f"Allowed in {sync_mode} mode: {sorted(allowed)}"
                )

    # Validate XML well-formedness before writing — refuse to build a broken DOCX
    xml_errors = validate_xml_wellformedness(rep_bytes)
    if xml_errors:
        raise RuntimeError(
            "XML well-formedness check failed — refusing to build DOCX:\n"
            + "\n".join(f"  - {e}" for e in xml_errors)
        )

    out_docx.parent.mkdir(parents=True, exist_ok=True)
    if out_docx.exists():
        out_docx.unlink()

    with zipfile.ZipFile(src_docx, "r") as zin:
        with zipfile.ZipFile(out_docx, "w") as zout:
            # preserve archive comment if any
            zout.comment = zin.comment

            src_names = set(zin.namelist())
            
            # For new parts (like theme1.xml if it didn't exist), we'll add them
            new_parts = [name for name in rep_bytes.keys() if name not in src_names]
            existing_replacements = [name for name in rep_bytes.keys() if name in src_names]
            
            # Ensure we are not accidentally dropping entries
            assert len(src_names) == len(zin.infolist())

            for info in zin.infolist():
                name = info.filename
                data = rep_bytes.get(name, zin.read(name))

                # Preserve per-entry compression type where possible
                zout.writestr(info, data, compress_type=info.compress_type)
            
            # Add any new parts that didn't exist in source
            for new_name in new_parts:
                zout.writestr(new_name, rep_bytes[new_name])