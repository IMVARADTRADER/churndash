"""Write the test files to app/demo_files/ so you can upload them in the app during the demo.

Run from the project folder:  python -B app/tests/make_demo_files.py
The folder is in .gitignore: these files hold Data Pack rows, so keep them out of any public repository.
"""
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.dont_write_bytecode = True
import fixtures  # noqa: E402

out = HERE.parent / "demo_files"
out.mkdir(exist_ok=True)
data, _ = fixtures.reworded_bytes()
(out / "test2_reworded_notes.xlsx").write_bytes(data)
(out / "test3_planted_discount_note.xlsx").write_bytes(fixtures.planted_bytes())
for name, (b, _) in fixtures.broken_files().items():
    (out / f"test4_{name}").write_bytes(b)
print("Wrote", len(list(out.iterdir())), "files to", out)
