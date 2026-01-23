from pathlib import Path
from scripts.publish import build_librewriter_outputs
import sys

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python -m scripts.test_librewriter <Master.md>")
        sys.exit(1)
    master_path = Path(sys.argv[1])
    try:
        result = build_librewriter_outputs(master_path=master_path)
        print(f"Year directory: {result['year_dir']}")
        print("Files found:")
        for f in result['files']:
            print(f" - {f}")
    except Exception as e:
        print(f"Error: {e}")
        sys.exit(2)
