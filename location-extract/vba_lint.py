"""Structural lint for a VBA .bas file: block balance, line continuations,
GoTo labels, undeclared identifiers under Option Explicit, duplicate procs."""
import re, sys, collections

path = sys.argv[1]
raw = open(path, encoding='utf-8', errors='replace').read().split('\n')

def strip_line(line):
    """Remove a trailing comment and mask string literals (keeps positions)."""
    out, i, n = [], 0, len(line)
    while i < n:
        c = line[i]
        if c == '"':
            j = i + 1
            while j < n:
                if line[j] == '"':
                    if j + 1 < n and line[j+1] == '"':
                        j += 2; continue
                    break
                j += 1
            out.append('S' * (min(j, n-1) + 1 - i)); i = j + 1; continue
        if c == "'":
            break
        out.append(c); i += 1
    return ''.join(out)

# join line continuations into logical lines
logical, buf, start = [], '', 0
for idx, line in enumerate(raw, 1):
    code = strip_line(line).rstrip()
    cont = code.endswith(' _')
    body = code[:-2].rstrip() if cont else code
    if not buf:
        start = idx
    buf = (buf + ' ' + body.strip()).strip() if buf else body.strip()
    if cont:
        continue
    if buf:
        logical.append((start, buf))
    buf = ''
if buf:
    logical.append((start, buf))
    print("!! UNTERMINATED line continuation at EOF")

errors = []
BLOCK_OPEN = {
    'sub': 'end sub', 'function': 'end function', 'property get': 'end property',
    'property let': 'end property', 'property set': 'end property',
}
stack = []          # (kind, line)
labels = collections.defaultdict(set)   # proc-name -> labels
gotos  = collections.defaultdict(list)  # proc-name -> (line, label)
declared = collections.defaultdict(set) # proc-name -> identifiers
used     = collections.defaultdict(list)
procs    = []
current  = None
proc_line = {}

for lineno, code in logical:
    low = code.lower().strip()
    m = re.match(r'^(?:public |private |friend |static )*?(sub|function|property (?:get|let|set))\s+([a-z0-9_]+)', low)
    if m and not low.startswith('end '):
        name = m.group(2)
        if current:
            errors.append(f"L{proc_line[current]}: proc '{current}' not closed before '{name}'")
        current = name; proc_line[name] = lineno; procs.append(name)
        # params + local decls come from following lines; collect signature text
        declared[name.lower()] |= set(re.findall(r'[a-z_][a-z0-9_]*', code.split('(',1)[-1]))
        continue
    if low.startswith('end sub') or low.startswith('end function') or low.startswith('end property'):
        if not current:
            errors.append(f"L{lineno}: '{low}' outside any procedure")
        else:
            stack = [s for s in stack]  # keep
            current = None
        continue

    # labels (e.g. "Failed:") only valid inside a proc
    lab = re.match(r'^([a-z_][a-z0-9_]*)\s*:\s*$', low)
    if lab and current:
        labels[current].add(lab.group(1))
        continue
    for g in re.findall(r'\bgoto\s+([a-z_][a-z0-9_]*)', low):
        if current: gotos[current].append((lineno, g))

    # declarations
    dm = re.match(r'^\s*(dim|redim|static|const|public|private|friend)\b(.*)$', low)
    if dm and not current:
        for part in re.split(r',(?![^()]*\))', dm.group(2)):
            nm = re.match(r'\s*([a-z_][a-z0-9_]*)\s*(\(|\bas\b|\=|$)', part.strip())
            if nm: declared['<MODULE>'].add(nm.group(1))
        continue
    if dm and current:
        decl = dm.group(2)
        for part in re.split(r',(?![^()]*\))', decl):
            nm = re.match(r'\s*([a-z_][a-z0-9_]*)\s*(\(|\bas\b|\=|$)', part.strip())
            if nm: declared[current].add(nm.group(1))
        continue

    # block balance
    if re.match(r'^if\b.*\bthen\s*$', low) and not low.rstrip().endswith('_'):
        stack.append(('if', lineno))
    elif re.match(r'^(select case|for |for\b|while |with |do\b|do$)', low):
        kind = low.split()[0]
        stack.append((kind if kind in ('select','for','while','with','do') else kind, lineno))
    if re.match(r'^end if\b', low):
        if not stack or stack[-1][0] != 'if': errors.append(f"L{lineno}: 'End If' with no open block If")
        else: stack.pop()
    elif re.match(r'^end select\b', low):
        if not stack or stack[-1][0] != 'select': errors.append(f"L{lineno}: 'End Select' mismatch")
        else: stack.pop()
    elif re.match(r'^next\b', low):
        if not stack or stack[-1][0] != 'for': errors.append(f"L{lineno}: 'Next' with no open For")
        else: stack.pop()
    elif re.match(r'^wend\b', low):
        if not stack or stack[-1][0] != 'while': errors.append(f"L{lineno}: 'Wend' with no open While")
        else: stack.pop()
    elif re.match(r'^end with\b', low):
        if not stack or stack[-1][0] != 'with': errors.append(f"L{lineno}: 'End With' with no open With")
        else: stack.pop()
    elif re.match(r'^loop\b', low):
        if not stack or stack[-1][0] != 'do': errors.append(f"L{lineno}: 'Loop' with no open Do")
        else: stack.pop()

    if current:
        for tok in re.findall(r'\b([a-z_][a-z0-9_]*)\b', low):
            used[current].append((lineno, tok))

if stack:
    for kind, ln in stack: errors.append(f"L{ln}: unclosed block '{kind}'")

VBA_KEYWORDS = set("""as byref byval to step then else elseif end if and or not xor mod is nothing
true false new set let get dim redim public private friend static const optional paramarray
sub function property exit for next while wend do loop until select case with each in goto on error
resume next integer long double string boolean variant date object single currency byte decimal
err vbnullstring vbcrlf vbcr vblf vbyes vbno vbcancel vbinformation vbexclamation vbquestion
vbyesno vbyesnocancel me mid left right len val cstr cdbl clng cint trim format instr instrrev
split join replace ucase lcase strcomp array ubound lbound isnumeric createobject thisworkbook
application debug print msgbox inputbox now date dateserial year month day timeserial timediff
chr chrw str fix int abs sqr cos sin tan exp log rnd cdate cvar typename vartype appactivate
sendkeys base comparemode exists items keys add remove removeall textcompare binarycompare
abort open send setrequestheader settimeouts status statustext responsetext readystate
value2 numberformat hasformula clearcontents cells range rows count columns parent name index
listobjects listrows listcolumns headerrowrange row global ignorecase pattern test execute
submatches followhyperlink statusbar ontime earliesttime procedure schedule wait
worksheets activesheet screenupdating enableevents displayalerts caption visible
getactivewindow settext putinclipboard getfromclipboard clipboard""".split())

undecl = []
for proc, toks in used.items():
    have = declared.get(proc, set()) | declared.get(proc.lower(), set()) | declared.get('<MODULE>', set())
    seen = set()
    for lineno, tok in toks:
        if tok in seen: continue
        seen.add(tok)
        if set(tok) <= {'s'}: continue
        if tok in VBA_KEYWORDS or tok in have or tok == proc: continue
        if tok in {p.lower() for p in procs}: continue
        undecl.append((proc, lineno, tok))

print(f"procedures: {len(procs)}   logical lines: {len(logical)}")
dups = [p for p,c in collections.Counter(procs).items() if c>1]
print("duplicate procs:", dups or "none")
bad_gotos = [(p,l,g) for p,gs in gotos.items() for l,g in gs if g not in labels.get(p,set())]
print("GoTo without label:", bad_gotos or "none")
print("block/structure errors:", len(errors))
for e in errors[:25]: print("  ", e)
print("possibly-undeclared identifiers:", len(undecl))
for u in sorted(set((p,t) for p,_,t in undecl))[:40]: print("  ", u)
