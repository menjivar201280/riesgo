import sys, datetime as dt
sys.path.insert(0, "/home/user/riesgo/cronograma-riesgos")
import build_cronograma as B
D = dt.date
RO = B.PUESTOS[0]
def row(name, s, sub=False, ini=D(2026,10,5), fin=D(2026,10,30), peso=None, obs=None):
    s = tuple(s) + (None,) * (4 - len(s))
    d = dict(area=RO, ini=ini, fin=fin, peso=peso, s1=s[0], s2=s[1], s3=s[2], s4=s[3], obs=obs, resp=None)
    d["sub" if sub else "macro"] = name
    return d
rows = [
    row("Matriz de riesgos operativos", ()),
    row("Levantar información", (), sub=True, peso=0.30),
    row("Validar con áreas", (), sub=True, peso=0.20),
    row("Elaborar matriz", (), sub=True),
    row("Presentar a comité", (), sub=True),
    row("Informe de eventos de pérdida", (0.3, 0.6)),
]
B.build(sys.argv[1], test_rows=rows)
