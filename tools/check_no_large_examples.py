from pathlib import Path
MAX_BYTES = 300_000
root = Path(__file__).resolve().parents[1]
bad = []
for p in root.rglob('*'):
    if p.is_file() and '.git' not in p.parts and p.stat().st_size > MAX_BYTES:
        bad.append((p.relative_to(root), p.stat().st_size))
if bad:
    for path, size in bad:
        print(f'{path}: {size} bytes')
    raise SystemExit(1)
print('OK: no large example files')
