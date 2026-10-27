# -*- coding: utf-8 -*-
import sys
sys.path.insert(0, '.')
from hu_niche import _hu_detect_niche, _opener_matches_niche

# Opener validation with REAL Hungarian labels
val_cases = [
    ('Az a WPC terasz Debrecenben jol nez ki.', 'nyilaszaro', False),
    ('Az a Schuco antracit ablak Gyorben jol nez ki.', 'nyilaszaro', True),
    ('Az a Daikin split Pecsett jol nez ki.', 'futesszerelo', False),
    ('Az a Vaillant kazan Debrecenben jol nez ki.', 'futesszerelo', True),
    ('Az a Vaillant kazan Debrecenben jol nez ki.', 'klima', False),
    ('Az a Daikin egyseg Szombathelyen jol nez ki.', 'klima', True),
    ('Az a napelem teto Egerben jol nez ki.', 'nyilaszaro', False),
    ('Ratalatam az oldalatokra Gyorben.', 'nyilaszaro', True),
]

# Map to real labels using _hu_detect_niche_from_name
from hu_niche import _hu_detect_niche_from_name
label_map = {
    'nyilaszaro': _hu_detect_niche_from_name('Ablak Bt'),
    'klima': _hu_detect_niche_from_name('Klima Kft'),
    'futesszerelo': _hu_detect_niche_from_name('Gazmester Bt'),
}
print('Label map:', label_map)

p = 0; f = 0
for opener, niche_key, expected in val_cases:
    real_niche = label_map.get(niche_key, niche_key)
    result = _opener_matches_niche(opener, real_niche)
    ok = result == expected
    marker = "OK  " if ok else "FAIL"
    if not ok:
        f += 1
        print(f"{marker} niche={real_niche!r}, expected={expected}: {opener[:60]!r}")
    else:
        p += 1
        print(f"{marker} niche={real_niche!r}: {result}")
print(f"
Opener validation: {p}/{p+f} passed")
