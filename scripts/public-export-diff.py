import re, subprocess, sys
def exported(rev):
    src = subprocess.run(['git','show',f'{rev}:src/core/index.ts'],
                         capture_output=True, text=True).stdout
    src = re.sub(r'/\*.*?\*/', '', src, flags=re.S)
    src = re.sub(r'//[^\n]*', '', src)
    names = set()
    # `export {` and `export type {` alike
    for block in re.findall(r'export\s+(?:type\s+)?\{(.*?)\}', src, re.S):
        for item in block.split(','):
            item = re.sub(r'^type\s+', '', item.strip())
            item = re.split(r'\s+as\s+', item)[-1].strip()
            if re.fullmatch(r'[A-Za-z_][A-Za-z0-9_]*', item):
                names.add(item)
    return names
b, a = exported(sys.argv[1] + '^'), exported(sys.argv[1])
print(f"之前 {len(b)} 个 → 之后 {len(a)} 个；移除 {len(b-a)}，新增 {len(a-b)}")
if len(b-a) <= 12 and b-a:
    print("  移除:", ", ".join(sorted(b-a)))
