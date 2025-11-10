#!/usr/bin/env python3
import subprocess, json, tempfile, os, argparse

parser = argparse.ArgumentParser(description='Extract PdfBox geometry for baseline and candidate PDFs')
parser.add_argument('--base', required=True)
parser.add_argument('--candidate', required=True)
parser.add_argument('--pages', default='1')
parser.add_argument('--pdfbox-dir', default='PdfVisualDiff/tools/pdfbox')
parser.add_argument('--output', default='out/diff-spacing.json')
args = parser.parse_args()

java_cp = os.pathsep.join([
    os.path.join(args.pdfbox_dir, 'geometry-extractor.jar'),
    os.path.join(args.pdfbox_dir, 'pdfbox-app-3.0.2.jar'),
])

def extract(pdf, out_path):
    subprocess.run([
        'java', '-cp', java_cp, 'GeometryExtractor',
        '--pdf', pdf,
        '--pages', args.pages,
        '--output', out_path
    ], check=True)

with tempfile.NamedTemporaryFile(delete=False) as base_tmp:
    base_path = base_tmp.name
with tempfile.NamedTemporaryFile(delete=False) as cand_tmp:
    cand_path = cand_tmp.name

try:
    extract(args.base, base_path)
    extract(args.candidate, cand_path)
    with open(base_path) as f:
        base = json.load(f)
    with open(cand_path) as f:
        candidate = json.load(f)
    report = {'base': base, 'candidate': candidate}
    os.makedirs(os.path.dirname(args.output), exist_ok=True)
    with open(args.output, 'w') as f:
        json.dump(report, f, indent=2)
    print(f'Report saved to {args.output}')
finally:
    os.unlink(base_path)
    os.unlink(cand_path)
