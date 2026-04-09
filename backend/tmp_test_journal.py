"""
Run the actual processor and trace what happens for entity-to-partner and partner-to-entity.
"""
import sys
sys.path.insert(0, '.')
import logging
logging.basicConfig(level=logging.DEBUG)

from app.ic_processor import (
    read_journal_report, normalize_entity_code, extract_6digit_from_icp
)

# Read Parent journal
primary, fallback = read_journal_report(r'uploads\reports\31\inputs\Parent report.xlsx')

print("\n=== PRIMARY LOOKUP KEYS ===")
for k in sorted(primary.keys()):
    entries = primary[k]
    total = sum(e['debit'] - e['credit'] for e in entries)
    print(f"  {k} -> {len(entries)} entries, net raw={total:.2f}")

print(f"\nTotal primary keys: {len(primary)}")

print("\n=== FALLBACK LOOKUP KEYS ===")
for k in sorted(fallback.keys()):
    entries = fallback[k]
    total = sum(e['debit'] - e['credit'] for e in entries)
    print(f"  {k} -> {len(entries)} entries, net raw={total:.2f}")

print(f"\nTotal fallback keys: {len(fallback)}")

# Check specific keys  
test_keys = [
    ('001001', 'ICP_E101000', '120030'),
    ('001001', 'ICP_E101000', '121015'),
    ('101000', 'ICP_001001', '120030'),
    ('101000', 'ICP_001001', '121015'),
]

print("\n=== SPECIFIC KEY CHECKS ===")
for k in test_keys:
    in_primary = k in primary
    in_fallback = k in fallback
    print(f"  {k} -> primary={in_primary}, fallback={in_fallback}")
    if in_primary:
        for e in primary[k]:
            print(f"    debit={e['debit']}, credit={e['credit']}, entity={e['entity_code']}, acct={e['account_code']}")
