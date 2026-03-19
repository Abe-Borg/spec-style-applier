"""
Word Document Decomposer

Extracts DOCX files into their constituent XML parts for processing
by the Phase 2 MEP Specification Styling Engine.
"""

import zipfile
import shutil
from pathlib import Path


class DocxDecomposer:
    def __init__(self, docx_path):
        """
        Initialize the decomposer with a path to a .docx file.

        Args:
            docx_path: Path to the input .docx file
        """
        self.docx_path = Path(docx_path)
        self.extract_dir = None

    def extract(self, output_dir=None):
        """
        Extract the .docx file to a directory.

        Args:
            output_dir: Directory to extract to. If None, creates a directory
                    based on the docx filename.

        Returns:
            Path to the extraction directory
        """
        if output_dir is None:
            base_name = self.docx_path.stem
            output_dir = Path(f"{base_name}_extracted")
        else:
            output_dir = Path(output_dir)

        # Remove existing directory if it exists (OneDrive-safe)
        if output_dir.exists():
            import time
            import uuid
            max_retries = 3
            for attempt in range(max_retries):
                try:
                    shutil.rmtree(output_dir)
                    break
                except PermissionError as e:
                    if attempt < max_retries - 1:
                        print(f"Folder locked (OneDrive?), retrying in 2s... ({attempt + 1}/{max_retries})")
                        time.sleep(2)
                    else:
                        # Last resort: rename instead of delete
                        backup = output_dir.with_name(f"{output_dir.name}_old_{uuid.uuid4().hex[:8]}")
                        print(f"Cannot delete {output_dir}, renaming to {backup}")
                        output_dir.rename(backup)

        output_dir.mkdir(parents=True, exist_ok=True)

        # Extract the ZIP archive
        print(f"Extracting {self.docx_path} to {output_dir}...")
        with zipfile.ZipFile(self.docx_path, 'r') as zip_ref:
            zip_ref.extractall(output_dir)

        self.extract_dir = output_dir
        print(f"Extraction complete: {len(list(output_dir.rglob('*')))} items extracted")
        return output_dir
