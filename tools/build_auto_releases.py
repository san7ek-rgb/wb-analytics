"""Build immutable receiver and server releases. Contains public code only."""
from hashlib import sha256
import json
from pathlib import Path

root = Path(__file__).resolve().parents[1]
server_files = ('wb_funnel_collector.py', 'wb-funnel.service', 'wb-funnel.timer',
                'wb-comparison.service', 'wb-comparison.timer')


def immutable(path, data):
    path.parent.mkdir(parents=True, exist_ok=True)
    if path.exists() and path.read_bytes() != data:
        raise SystemExit('Existing release content changed: ' + str(path))
    path.write_bytes(data)


receiver = (root / 'CostReceiver.gs').read_bytes()
digest = sha256(receiver).hexdigest()
receiver_path = 'receiver-releases/' + digest + '.gs'
immutable(root / receiver_path, receiver)
(root / 'receiver-release.json').write_text(json.dumps({'schema': 1, 'sha256': digest, 'path': receiver_path}, indent=2) + '\n')
contents = {name: (root / 'server' / name).read_bytes() for name in server_files}
metadata = {name: {'sha256': sha256(data).hexdigest(), 'size': len(data)} for name, data in contents.items()}
identity = {'schema': 1, 'files': metadata}
digest = sha256(json.dumps(identity, sort_keys=True, separators=(',', ':')).encode()).hexdigest()
for name, data in contents.items():
    immutable(root / 'server-releases' / digest / name, data)
(root / 'server-release.json').write_text(json.dumps({**identity, 'sha256': digest}, indent=2) + '\n')
print('Receiver and server immutable releases built.')
