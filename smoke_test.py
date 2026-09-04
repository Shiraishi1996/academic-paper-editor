"""Minimal endpoint and preprocessing smoke test. Run after `pip install -r requirements.txt`."""
from app import app, _protected_segments, _apply_glossary_target_terms, parse_translation_glossary

# Pure preprocessing checks: these do not load neural translation models.
marked = _apply_glossary_target_terms(
    "図2に示す。可変浮力を用いたPID制御を評価した [1–3]。",
    "en",
    "可変浮力 = variable buoyancy",
)
segments = _protected_segments(marked, "en", "")
joined = "".join(x for _, x in segments)
assert "Figure 2" in joined
assert "[1–3]" in joined
assert "⟦variable buoyancy⟧" in joined
assert parse_translation_glossary("可変浮力 = variable buoyancy")
print("OK translation preprocessing")

checks = [
    ("GET", "/api/health"),
    ("GET", "/api/paper"),
    ("GET", "/api/translation/status"),
    ("GET", "/api/assets"),
    ("GET", "/api/export/latex"),
    ("GET", "/api/export/docx"),
]

client = app.test_client()
failed = 0
for method, path in checks:
    response = client.open(path, method=method)
    ok = response.status_code < 400
    print(f"{'OK' if ok else 'FAIL'} {response.status_code:3d} {method} {path}")
    if not ok:
        print(response.get_data(as_text=True)[:1000])
        failed += 1

raise SystemExit(1 if failed else 0)
