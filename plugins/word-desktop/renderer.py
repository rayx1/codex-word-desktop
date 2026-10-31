"""Separate Poppler PDF -> PNG step. Standard library only; never opens Word."""
import argparse
import json
import os
from pathlib import Path
import re
import shutil
import subprocess
import sys
import uuid


def executable(name):
    found = shutil.which(name)
    if found:
        return found
    root = Path(os.environ.get('USERPROFILE', '')) / '.cache/codex-runtimes/codex-primary-runtime/dependencies/native/poppler/Library/bin'
    candidate = root / (name + '.exe')
    if candidate.is_file():
        return str(candidate)
    raise RuntimeError(f'{name} missing: install Poppler and add its bin directory to PATH')


def run(args):
    p = subprocess.run(args, capture_output=True, text=True, encoding='utf-8', errors='replace', timeout=120)
    if p.returncode:
        raise RuntimeError(p.stderr.strip() or f'Exit {p.returncode}')
    return p.stdout


def diagnose():
    """Use the same binary resolution as rendering; no documents are opened."""
    checks = []
    checks.append(dict(name='python', ok=sys.version_info >= (3, 10),
                       path=sys.executable, version=sys.version.split()[0],
                       error=None if sys.version_info >= (3, 10) else 'Python is older than 3.10.',
                       action='Use Python 3.10 or newer.'))
    for name in ('pdfinfo', 'pdftoppm'):
        try:
            path = executable(name)
            p = subprocess.run([path, '-v'], capture_output=True, text=True,
                               encoding='utf-8', errors='replace', timeout=10)
            if p.returncode:
                raise RuntimeError(p.stderr.strip() or f'Exit {p.returncode}')
            lines = (p.stdout + p.stderr).strip().splitlines()
            checks.append(dict(name=name, ok=True, path=path,
                               version=lines[0] if lines else 'version not reported'))
        except Exception as exc:
            checks.append(dict(name=name, ok=False, error=str(exc),
                               action='Install a trusted Poppler build and put its bin directory on PATH.'))
    return dict(ok=all(c['ok'] for c in checks), checks=checks)


def render(pdf, work, pages=None, dpi=144):
    result = dict(step='pdf_to_images', page_count=None, images=[], errors=[], dpi=dpi)
    try:
        if not 72 <= dpi <= 300:
            raise ValueError('dpi must be between 72 and 300')
        pdf = Path(pdf).resolve(strict=True)
        info = run([executable('pdfinfo'), str(pdf)])
        count = int(re.search(r'^Pages:\s+(\d+)', info, re.M)[1])
        result['page_count'] = count
        chosen = list(range(1, count + 1)) if pages is None else sorted(set(pages))
        if not chosen or any(type(n) is not int or n < 1 or n > count for n in chosen):
            raise ValueError(f'pages must contain integers from 1 to {count}')
        out = Path(work).resolve() / ('render-' + uuid.uuid4().hex)
        out.mkdir(parents=True)
        result['directory'] = str(out)
        for page in chosen:
            prefix = out / f'page-{page:04d}'
            try:
                run([executable('pdftoppm'), '-f', str(page), '-l', str(page), '-r', str(dpi), '-singlefile', '-png', str(pdf), str(prefix)])
                image = prefix.with_suffix('.png')
                if not image.is_file() or image.stat().st_size == 0:
                    raise RuntimeError('Poppler returned no image')
                result['images'].append(dict(page=page, path=str(image)))
            except Exception as exc:
                result['errors'].append(dict(page=page, message=str(exc)))
    except Exception as exc:
        result['errors'].append(dict(message=str(exc)))
    result['ok'] = not result['errors']
    return result


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('pdf', nargs='?')
    parser.add_argument('--diagnose', action='store_true', help='Check Python and Poppler without opening a PDF.')
    parser.add_argument('--work', default=str(Path(__file__).parent / 'work'))
    parser.add_argument('--pages', type=int, nargs='+')
    parser.add_argument('--dpi', type=int, default=144)
    a = parser.parse_args()
    if a.diagnose:
        result = diagnose()
        print(json.dumps(result, ensure_ascii=False))
        sys.exit(0 if result['ok'] else 1)
    if not a.pdf:
        parser.error('a PDF path is required unless --diagnose is used')
    print(json.dumps(render(a.pdf, a.work, a.pages, a.dpi), ensure_ascii=False))
