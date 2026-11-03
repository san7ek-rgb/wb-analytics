"""Собрать проверяемый выпуск; затем закоммитить и отправить изменения в main."""
from hashlib import sha256
from pathlib import Path
import json

root = Path(__file__).resolve().parents[1]
source = (root / 'src' / 'Code.gs').read_bytes()
digest = sha256(source).hexdigest()
path = f'releases/{digest}.gs'
(root / 'releases').mkdir(exist_ok=True)
target = root / path
if target.exists() and target.read_bytes() != source:
    raise SystemExit('Уже существующий выпуск отличается от исходника.')
target.write_bytes(source)
(root / 'release.json').write_text(
    json.dumps({'schema': 1, 'path': path, 'sha256': digest}, indent=2) + '\n',
    encoding='utf-8',
)
print('Выпуск:', digest[:12])
