import sys
from pathlib import Path

def clean(p):
    p = Path(p)
    if not p.exists():
        print('MISSING', p)
        return 2
    b = p.read_bytes()
    n = b.count(b"\x00")
    if n == 0:
        print(f'No null bytes in {p}');
        return 0
    bak = p.with_suffix(p.suffix + '.bak')
    p.rename(bak)
    nb = b.replace(b"\x00", b"")
    p.write_bytes(nb)
    print(f'Replaced {n} null bytes in {p}; backup at {bak}')
    # quick syntax check
    try:
        import py_compile
        py_compile.compile(str(p), doraise=True)
        print('Syntax OK')
        return 0
    except Exception as e:
        print('Syntax check failed:', e)
        return 3

if __name__ == '__main__':
    if len(sys.argv) < 2:
        print('Usage: clean_nulls.py <file>')
        sys.exit(1)
    sys.exit(clean(sys.argv[1]))
