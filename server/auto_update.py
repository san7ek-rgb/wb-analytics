#!/usr/bin/env python3
"""Fixed-scope systemd updater. Never reads .env, browser profiles or WB data."""
import argparse
import fcntl
import hashlib
import json
import os
from pathlib import Path
import pwd
import re
import shutil
import subprocess
import tempfile
import time
from urllib.request import urlopen
from urllib.error import URLError

BASE = 'https://raw.githubusercontent.com/san7ek-rgb/wb-analytics/main/'
APP = Path('/home/reportdeploy/wb-automation')
ROOT = Path('/var/lib/wb-autoupdate')
UNITS = Path('/etc/systemd/system')
FILES = ('wb_funnel_collector.py', 'wb-funnel.service', 'wb-funnel.timer',
         'wb-comparison.service', 'wb-comparison.timer')
TIMERS = ('wb-funnel.timer', 'wb-comparison.timer')
SERVICES = ('wb-funnel.service', 'wb-comparison.service')
HEX = re.compile(r'^[a-f0-9]{64}$')


class Busy(Exception):
    pass


def command(*args):
    result = subprocess.run(args, capture_output=True, text=True, timeout=60)
    if result.returncode:
        raise RuntimeError('Command failed: ' + args[0] + ' ' + args[1])
    return result.stdout.strip()


def fetch(path, limit):
    with urlopen(BASE + path, timeout=45) as response:
        data = response.read(limit + 1)
    if len(data) > limit:
        raise ValueError('Release file too large')
    return data


def canonical(value):
    return json.dumps(value, sort_keys=True, separators=(',', ':')).encode()


def validate_manifest(manifest):
    if not isinstance(manifest, dict) or manifest.get('schema') != 1:
        raise ValueError('Unsupported release schema')
    release = manifest.get('sha256', '')
    files = manifest.get('files')
    if not isinstance(release, str) or not HEX.fullmatch(release) or not isinstance(files, dict) or set(files) != set(FILES):
        raise ValueError('Invalid release file list')
    for meta in files.values():
        if not isinstance(meta, dict) or not isinstance(meta.get('sha256'), str) or not HEX.fullmatch(meta['sha256']) or type(meta.get('size')) is not int or not 0 < meta['size'] <= 512000:
            raise ValueError('Invalid release checksum or size')
    expected = hashlib.sha256(canonical({'schema': 1, 'files': files})).hexdigest()
    if release != expected:
        raise ValueError('Release manifest checksum mismatch')
    return manifest


def download_release(manifest, stage):
    validate_manifest(manifest)
    for name, meta in manifest['files'].items():
        data = fetch('server-releases/' + manifest['sha256'] + '/' + name, meta['size'])
        if len(data) != meta['size'] or hashlib.sha256(data).hexdigest() != meta['sha256']:
            raise ValueError('Checksum mismatch: ' + name)
        (stage / name).write_bytes(data)
    # Compiling does not execute downloaded code.
    compile((stage / 'wb_funnel_collector.py').read_bytes(), 'collector', 'exec')


def target(name):
    if name not in FILES:
        raise ValueError('Unknown target')
    return APP / name if name.endswith('.py') else UNITS / name


def install_bytes(path, data, collector=False):
    fd, temporary = tempfile.mkstemp(prefix='.' + path.name + '.', dir=path.parent)
    try:
        with os.fdopen(fd, 'wb') as out:
            out.write(data)
            out.flush()
            os.fsync(out.fileno())
        os.chmod(temporary, 0o600 if collector else 0o644)
        if collector:
            user = pwd.getpwnam('reportdeploy')
            os.chown(temporary, user.pw_uid, user.pw_gid)
        os.replace(temporary, path)
    finally:
        if os.path.exists(temporary):
            os.unlink(temporary)


def save_state(state):
    install_bytes(ROOT / 'state.json', json.dumps(state).encode())


def read_state():
    path = ROOT / 'state.json'
    return json.loads(path.read_text()) if path.exists() else {}


def active(unit):
    result = subprocess.run(['systemctl', 'is-active', unit], capture_output=True, text=True, timeout=30)
    value = result.stdout.strip()
    if value in ('active', 'activating', 'deactivating', 'reloading'):
        return True
    if value in ('inactive', 'failed', 'unknown') or result.returncode == 4:
        return False
    raise RuntimeError('Cannot read service state: ' + unit)


def enabled(unit):
    result = subprocess.run(['systemctl', 'is-enabled', unit], capture_output=True, text=True, timeout=30)
    return result.returncode == 0 and result.stdout.strip() in ('enabled', 'enabled-runtime')


def finish_bootstrap():
    marker = ROOT / 'bootstrap-pending'
    if marker.exists():
        command('systemctl', 'enable', '--now', *TIMERS)
        marker.unlink()


def snapshot(state):
    folder = Path(tempfile.mkdtemp(prefix='backup-', dir=ROOT))
    present = []
    for name in FILES:
        path = target(name)
        if path.is_symlink():
            raise ValueError('Refusing symlink target: ' + name)
        if path.exists():
            shutil.copyfile(path, folder / name)
            present.append(name)
    meta = {'state': state, 'present': present, 'active_timers': [t for t in TIMERS if active(t)],
            'enabled_timers': [t for t in TIMERS if enabled(t)]}
    (folder / 'backup.json').write_text(json.dumps(meta))
    return folder


def restore(folder):
    folder = Path(folder)
    if folder.parent != ROOT or not folder.name.startswith('backup-'):
        raise ValueError('Invalid backup path')
    meta = json.loads((folder / 'backup.json').read_text())
    for timer in TIMERS:
        if active(timer):
            command('systemctl', 'stop', timer)
        if timer not in meta['present'] and target(timer).exists():
            command('systemctl', 'disable', timer)
    for name in FILES:
        if name in meta['present']:
            install_bytes(target(name), (folder / name).read_bytes(), name.endswith('.py'))
        else:
            target(name).unlink(missing_ok=True)
    command('systemctl', 'daemon-reload')
    for timer in TIMERS:
        if timer in meta['present']:
            command('systemctl', 'enable' if timer in meta.get('enabled_timers', []) else 'disable', timer)
    for timer in meta['active_timers']:
        command('systemctl', 'restart', timer)
    return meta['state']


def verify_runtime(stage):
    command('systemd-analyze', 'verify', *(str(stage / n) for n in FILES if not n.endswith('.py')))
    # Run imports as the unprivileged collector user, without loading .env.
    candidate = APP / '.update-candidate.py'
    try:
        install_bytes(candidate, (stage / 'wb_funnel_collector.py').read_bytes(), True)
        command('runuser', '-u', 'reportdeploy', '--', str(APP / 'venv/bin/python'), str(candidate), '--self-check')
    finally:
        candidate.unlink(missing_ok=True)


def apply_release(manifest, stage, state):
    verify_runtime(stage)
    backup = snapshot(state)
    meta = json.loads((backup / 'backup.json').read_text())
    try:
        for timer in meta['active_timers']:
            command('systemctl', 'stop', timer)
        if any(active(service) for service in SERVICES):
            raise Busy('Collection started during staging; update deferred')
        for name in FILES:
            install_bytes(target(name), (stage / name).read_bytes(), name.endswith('.py'))
        command('systemctl', 'daemon-reload')
        for timer in meta['active_timers']:
            command('systemctl', 'restart', timer)
        save_state({'current': manifest['sha256'], 'previous_backup': str(backup), 'updated': int(time.time())})
    except Busy:
        for timer in meta['active_timers']:
            command('systemctl', 'start', timer)
        raise
    except Exception:
        restored = restore(backup)
        restored['rejected'] = manifest['sha256']
        save_state(restored)
        raise RuntimeError('Installation failed; previous files restored')


def update(rollback=False):
    if APP.is_symlink() or APP.resolve() != APP or not (APP / 'venv/bin/python').is_file():
        raise ValueError('Existing collector installation missing or redirected')
    ROOT.mkdir(mode=0o700, parents=True, exist_ok=True)
    os.chmod(ROOT, 0o700)
    with open(ROOT / 'update.lock', 'a') as updater_lock:
        try:
            fcntl.flock(updater_lock, fcntl.LOCK_EX | fcntl.LOCK_NB)
        except BlockingIOError:
            raise Busy('Another updater is running')
        state = read_state()
        if rollback:
            manifest = None
        else:
            manifest = validate_manifest(json.loads(fetch('server-release.json?v=' + str(time.time_ns()), 16000)))
            if manifest['sha256'] == state.get('current'):
                finish_bootstrap()
                return
            if manifest['sha256'] == state.get('rejected'):
                print('Rejected release remains blocked; waiting for a newer release.', flush=True)
                return
        collector_path = APP / '.collector.lock'
        if collector_path.is_symlink():
            raise ValueError('Collector lock must not be a symlink')
        with open(collector_path, 'a') as collector_lock:
            user = pwd.getpwnam('reportdeploy')
            os.fchown(collector_lock.fileno(), user.pw_uid, user.pw_gid)
            os.fchmod(collector_lock.fileno(), 0o600)
            try:
                fcntl.flock(collector_lock, fcntl.LOCK_EX | fcntl.LOCK_NB)
            except BlockingIOError:
                raise Busy('Collection running; update deferred')
            if any(active(service) for service in SERVICES):
                raise Busy('Collector service active; update deferred')
            if rollback:
                if not state.get('previous_backup'):
                    raise ValueError('No previous release backup')
                restored = restore(state['previous_backup'])
                restored['rejected'] = state.get('current')
                save_state(restored)
                print('Previous server files restored; reverted release blocked.', flush=True)
                return
            with tempfile.TemporaryDirectory(prefix='release-', dir=ROOT) as directory:
                stage = Path(directory)
                try:
                    download_release(manifest, stage)
                    apply_release(manifest, stage, state)
                except (URLError, TimeoutError, Busy):
                    # A busy server or network outage must retry the SAME release.
                    raise
                except Exception:
                    failed = read_state()
                    failed['rejected'] = manifest['sha256']
                    save_state(failed)
                    raise
            finish_bootstrap()
            print('Server release applied: ' + manifest['sha256'][:12], flush=True)


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--rollback', action='store_true')
    args = parser.parse_args()
    if os.geteuid() != 0:
        raise SystemExit('Run the updater as root.')
    try:
        update(args.rollback)
    except Busy as error:
        print(str(error), flush=True)
    except Exception as error:
        # Errors contain fixed status descriptions and public filenames, not .env.
        print('Auto-update failed: ' + str(error), flush=True)
        raise SystemExit(1)
