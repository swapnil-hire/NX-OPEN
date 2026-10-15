"""Source consistency only. Does NOT compile VB.NET or verify NX/Excel execution."""
from pathlib import Path
import re
source = Path(__file__).resolve().parents[1] / 'NX2306_Drawing_Checker.vb'
s = source.read_text()
clean = []
for n, line in enumerate(s.splitlines(), 1):
    out = ''; j = 0; quoted = False
    while j < len(line):
        char = line[j]
        if char == '"':
            if quoted and j + 1 < len(line) and line[j+1] == '"': j += 2; continue
            quoted = not quoted; j += 1; continue
        if not quoted:
            if char == "'": break
            out += char
        j += 1
    assert not quoted, f'Unterminated VB string at line {n}'
    clean.append(out)
stack = []
for n, line in enumerate(clean, 1):
    for char in line:
        if char in '({': stack.append((char, n))
        if char in ')}':
            assert stack and stack[-1][0] == {')': '(', '}': '{'}[char], (n, char, stack[-3:])
            stack.pop()
assert not stack, stack
for kind, start, end in [
    ('Sub', r'^\s*(?:(?:Public|Private|Friend|Protected|Shared|Overrides)\s+)*Sub\s+', r'^\s*End Sub\s*$'),
    ('Function', r'^\s*(?:(?:Public|Private|Friend|Protected|Shared)\s+)*Function\s+', r'^\s*End Function\s*$'),
    ('Class', r'^\s*(?:Public |Private |Friend )?Class\s+', r'^\s*End Class\s*$'),
    ('Module', r'^\s*Module\s+', r'^\s*End Module\s*$'),
    ('Using', r'^\s*Using\s+', r'^\s*End Using\s*$'),
    ('Try', r'^\s*Try\s*$', r'^\s*End Try\s*$'),
    ('For', r'^\s*For\s+', r'^\s*Next(?:\s+\w+)?\s*$'),
    ('While', r'^\s*While\s+', r'^\s*End While\s*$')]:
    a = sum(bool(re.search(start,x,re.I)) for x in clean)
    b = sum(bool(re.search(end,x,re.I)) for x in clean)
    # WinForms event multiline anonymous Sub closures have End Sub too.
    if kind == 'Sub': a += sum(bool(re.search(r'AddHandler .*\bSub\([^)]*\)\s*$',x,re.I)) for x in clean)
    assert a == b, (kind,a,b)
# Read-only source contract. Image export Commit is the only allowed Commit call.
assert len(re.findall(r'\bbuilder\.Commit\(',s)) == 1
for forbidden in [r'\.Save\(', r'\.SetAttribute\(', r'\.SetUserAttribute\(', r'\.SetDimensionText\(', r'\.Update\(', r'(?:annotation|dimension|view|fcf|datum)\.Color\s*=']:
    assert not re.search(forbidden,'\n'.join(clean),re.I), forbidden
assert 'builder.Destroy()' in s and 'app.Quit()' in s and 'book.Close(False)' in s
assert 'com.ReleaseAll()' in s
assert 'AskViewDependentStatus(annotation.Tag' in s
assert 'universal cross-part identity' in s
assert 'Case "FAIL", "MISSING", "ERROR"' in s
assert 'INCOMPLETE - NO OVERALL PASS' in s
print('PASS: VB lexical/delimiter/block-count consistency and source preservation checks.')
print('Not a compiler result; NX and Excel runtime validation remains required.')
