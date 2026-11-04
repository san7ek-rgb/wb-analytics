"""One publishing step for the original loader, receiver loader and server."""
from pathlib import Path
import subprocess
import sys

tools = Path(__file__).resolve().parent
for name in ('build_comparison_receiver.py', 'build_release.py', 'build_auto_releases.py'):
    subprocess.run([sys.executable, str(tools / name)], check=True)
