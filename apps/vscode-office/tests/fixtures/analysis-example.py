"""Repository entrypoint for the packaged standard-library workbook example."""
from pathlib import Path
import runpy

if __name__ == "__main__":
    runpy.run_path(str(Path(__file__).resolve().parents[2] / "skills" / "genoffice-workbook" / "scripts" / "analyze.py"), run_name="__main__")
