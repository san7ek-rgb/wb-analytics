"""Build a single Apps Script receiver file, including the comparison sidebar."""
import json
from pathlib import Path
root = Path(__file__).resolve().parents[1]
target = root / 'CostReceiver.gs'
marker = '\n// GENERATED COMPARISON UI — edit src/Comparison.gs and src/ComparisonSidebar.html\n'
base = target.read_text().split(marker)[0]
exports = ['doPost', 'wbInstallComparison', 'wbComparisonOnOpen', 'wbOpenComparison',
           'wbComparisonInit', 'wbQueueComparison', 'wbCancelComparison', 'wbComparisonStatus', 'wbHistoryStatus','wbStartHistoryImport','wbPauseHistoryImport','wbResumeHistoryImport']
target.write_text(base + marker + (root / 'src/Comparison.gs').read_text() + '\nconst WB_COMPARISON_HTML = ' + json.dumps((root / 'src/ComparisonSidebar.html').read_text(), ensure_ascii=False) + ';\n' +
                  'const WB_RECEIVER_EXPORTS = Object.freeze({' + ','.join(exports) + '});\n')
