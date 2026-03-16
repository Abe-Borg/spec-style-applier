# docx_patch.py
from __future__ import annotations

from pathlib import Path
import re
import zipfile
import xml.etree.ElementTree as ET
from typing import Dict, List, Set, Union

BytesOrStr = Union[bytes, str]


def validate_xml_wellformedness(replacements: Dict[str, bytes]) -> List[str]:
    """Parse each XML replacement part with ElementTree to verify well-formedness."""
    errors: List[str] = []
    for name, content in replacements.items():
        if not (name.endswith(".xml") or name.endswith(".rels") or name == "[Content_Types].xml"):
            continue
        try:
            ET.fromstring(content)
        except ET.ParseError as exc:
            errors.append(f"{name}: XML parse error: {exc}")
    return errors


def _is_allowed_patch(name: str, allowed_patches: Set[str]) -> bool:
    if name in allowed_patches:
        return True
    patterns = (
        r"word/header\d+\.xml$",
        r"word/footer\d+\.xml$",
        r"word/_rels/header\d+\.xml\.rels$",
        r"word/_rels/footer\d+\.xml\.rels$",
    )
    if any(re.match(pattern, name) for pattern in patterns):
        return True
    if name.startswith("word/media/"):
        return True
    return False


def patch_docx(
    src_docx: Path,
    out_docx: Path,
    replacements: Dict[str, BytesOrStr],
    exclude_parts: Set[str] | None = None,
) -> None:
    """
    Create out_docx by copying every ZIP entry from src_docx unchanged,
    except for entries whose internal paths match keys in `replacements`.

    This is NOT a "rebuild from extracted folder".
    It's a surgical patch: swap specific parts, preserve everything else.
    """
    src_docx = Path(src_docx)
    out_docx = Path(out_docx)
    exclude_parts = set(exclude_parts or set())

    rep_bytes: Dict[str, bytes] = {}
    for k, v in replacements.items():
        if isinstance(v, str):
            rep_bytes[k] = v.encode("utf-8")
        else:
            rep_bytes[k] = v

    FORBIDDEN_EXACT = set()

    ALLOWED_PATCHES = {
        "word/document.xml",
        "word/styles.xml",
        "word/theme/theme1.xml",
        "word/numbering.xml",
        "word/settings.xml",
        "word/fontTable.xml",
        "[Content_Types].xml",
        "word/_rels/document.xml.rels",
    }

    for name in rep_bytes:
        if name in FORBIDDEN_EXACT:
            raise RuntimeError(f"Forbidden patch target: {name}")

        if not _is_allowed_patch(name, ALLOWED_PATCHES):
            raise RuntimeError(
                f"Illegal patch target: {name}\n"
                f"Allowed base set: {sorted(ALLOWED_PATCHES)} plus header/footer/media patterns"
            )

    for name in exclude_parts:
        if not _is_allowed_patch(name, ALLOWED_PATCHES):
            raise RuntimeError(f"Illegal excluded part: {name}")

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

            # Ensure we are not accidentally dropping entries
            assert len(src_names) == len(zin.infolist())

            for info in zin.infolist():
                name = info.filename
                if name in exclude_parts:
                    continue
                data = rep_bytes.get(name, zin.read(name))

                # Preserve per-entry compression type where possible
                zout.writestr(info, data, compress_type=info.compress_type)

            # Add any new parts that didn't exist in source
            for new_name in new_parts:
                zout.writestr(new_name, rep_bytes[new_name])
