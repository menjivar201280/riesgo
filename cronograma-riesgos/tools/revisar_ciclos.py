"""Busca referencias circulares ESTÁTICAS (como las ve Excel) a nivel de columna en una hoja Cronograma."""
import sys, re
from openpyxl import load_workbook
from openpyxl.formula.tokenizer import Tokenizer
from openpyxl.utils import column_index_from_string, get_column_letter
wb = load_workbook(sys.argv[1]); ws = wb[sys.argv[2]]; row = int(sys.argv[3]) if len(sys.argv) > 3 else 6
graph = {}
for c in ws[row]:
    v = c.value
    if not (isinstance(v, str) and v.startswith("=")): continue
    deps = set()
    for t in Tokenizer(v).items:
        if t.type == "OPERAND" and t.subtype == "RANGE" and "!" not in t.value:
            for part in t.value.split(":"):
                m = re.match(r"\$?([A-Z]{1,3})\$?(\d+)?$", part)
                if m: deps.add(m.group(1))
    graph[c.column_letter] = deps
# columnas fuente (vacías o sin fórmula) no tienen aristas
seen, stack, cyc = {}, [], []
def dfs(u):
    seen[u] = 1; stack.append(u)
    for w in graph.get(u, ()):
        if w == u and u in graph:   # auto-referencia de la misma columna (rango)
            cyc.append([u]); continue
        if seen.get(w) == 1: cyc.append(stack[stack.index(w):] + [w])
        elif not seen.get(w): dfs(w)
    stack.pop(); seen[u] = 2
for u in list(graph): 
    if not seen.get(u): dfs(u)
print("CICLOS:", cyc[:10] if cyc else "ninguno")
