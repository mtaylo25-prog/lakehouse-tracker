#!/usr/bin/env python3
"""Compare strings/en.json and strings/ko.json: same keys, same array lengths, token sets per string, no dashes."""
import json, re, sys
en = json.load(open('build/strings/en.json')); ko = json.load(open('build/strings/ko.json'))
TOK = re.compile(r'\{([a-zA-Z0-9_]+)\}')
EXTRA_OK = {'s07': {'cat_8k', 'cat_4k'}, 's19': {'chip_d90', 'chip_m12', 'result_labels'}, 's21': {'compare'}, 'common': {'points'}}
probs = []
def walk(a, b, path):
    if isinstance(a, dict):
        if not isinstance(b, dict): probs.append(f'{path}: type'); return
        for k in a:
            if k not in b: probs.append(f'{path}.{k}: missing in ko')
        for k in b:
            if k not in a and k not in EXTRA_OK.get(path.lstrip('.').split('.')[0], set()): probs.append(f'{path}.{k}: extra in ko')
        for k in a:
            if k in b: walk(a[k], b[k], f'{path}.{k}')
    elif isinstance(a, list):
        if not isinstance(b, list) or len(a) != len(b): probs.append(f'{path}: list length {len(a)} vs {len(b) if isinstance(b, list) else type(b)}'); return
        for i, (x, y) in enumerate(zip(a, b)): walk(x, y, f'{path}[{i}]')
    elif isinstance(a, str):
        if not isinstance(b, str): probs.append(f'{path}: type'); return
        ta, tb = set(TOK.findall(a)), set(TOK.findall(b))
        if ta != tb and not (path.endswith('.note') and path.startswith('.s23')): probs.append(f'{path}: tokens en {sorted(ta)} ko {sorted(tb)}')
        if re.search(r'[-–—−]', b) and not re.search(r'SS-4|W-8BEN|599-64-00496|1120-F', b): probs.append(f'{path}: dash in ko: {b[:80]}')
        if path.endswith('.icon') and a != b: probs.append(f'{path}: icon changed')
walk(en, ko, '')
print('\n'.join(probs) if probs else 'strings OK: keys, lengths, tokens, no dashes')
sys.exit(1 if probs else 0)
