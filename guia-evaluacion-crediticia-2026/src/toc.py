"""Rellena el índice (campo TOC) con entradas y números de página calculados
a partir de una conversión a PDF con LibreOffice. Word puede actualizarlo luego."""
import re, sys, subprocess, zipfile, shutil, os, tempfile
from xml.sax.saxutils import escape

DOCX = sys.argv[1]
SOFFICE = sys.argv[2]

def read_doc(path):
    with zipfile.ZipFile(path) as z:
        return z.read("word/document.xml").decode("utf8")

def write_doc(src, dst, xml):
    with zipfile.ZipFile(src) as zi, zipfile.ZipFile(dst + ".tmp", "w", zipfile.ZIP_DEFLATED) as zo:
        for it in zi.infolist():
            data = zi.read(it.filename)
            if it.filename == "word/document.xml":
                data = xml.encode("utf8")
            zo.writestr(it, data)
    os.replace(dst + ".tmp", dst)

def headings(xml):
    out = []
    for m in re.finditer(r'<w:p>(?:(?!</w:p>).)*?<w:pStyle w:val="Heading([12])"/>(.*?)</w:p>', xml, re.S):
        txt = "".join(re.findall(r"<w:t[^>]*>(.*?)</w:t>", m.group(2)))
        out.append((int(m.group(1)), txt))
    return out

def entries(hs, pages):
    ps = []
    for (lvl, t), pg in zip(hs, pages):
        ps.append(f'<w:p><w:pPr><w:pStyle w:val="TOC{lvl}"/></w:pPr><w:r><w:t xml:space="preserve">{t}</w:t></w:r>'
                  f'<w:r><w:tab/></w:r><w:r><w:t>{pg}</w:t></w:r></w:p>')
    return "".join(ps)

def inject(xml, body):
    a = xml.index('<w:fldChar w:fldCharType="separate"/></w:r></w:p>') + len('<w:fldChar w:fldCharType="separate"/></w:r></w:p>')
    b = xml.index('<w:p><w:r><w:fldChar w:fldCharType="end"/>', a)
    return xml[:a] + body + xml[b:]

def unesc(t):
    return t.replace("&amp;", "&").replace("&lt;", "<").replace("&gt;", ">").replace("&quot;", '"')

def page_numbers(path, hs):
    d = tempfile.mkdtemp()
    shutil.copy(path, os.path.join(d, "t.docx"))
    subprocess.run([sys.executable, SOFFICE, "--headless", "--convert-to", "pdf", "--outdir", d, os.path.join(d, "t.docx")], capture_output=True)
    pdf = os.path.join(d, "t.pdf")
    n = int(re.search(r"Pages:\s+(\d+)", subprocess.run(["pdfinfo", pdf], capture_output=True, text=True).stdout).group(1))
    texts = []
    for i in range(1, n + 1):
        t = subprocess.run(["pdftotext", "-f", str(i), "-l", str(i), "-layout", pdf, "-"], capture_output=True, text=True).stdout
        texts.append(re.sub(r"\s+", " ", t))
    # saltar portada e índice: el cuerpo empieza en la página con el primer párrafo de la Presentación
    start = next(i for i, t in enumerate(texts) if "Esta guía explica, de forma ordenada" in t)
    res = []
    for _, t in hs:
        key = re.sub(r"\s+", " ", unesc(t))[:45]
        found = None
        for i in range(start, n):
            if key in texts[i]:
                found = i; break
        if found is None:
            res.append("—"); continue
        res.append(str(found + 1)); start = found
    return res, n

xml0 = read_doc(DOCX)
hs = headings(xml0)
xml1 = inject(xml0, entries(hs, ["00"] * len(hs)))
write_doc(DOCX, DOCX, xml1)
pages, n = page_numbers(DOCX, hs)
write_doc(DOCX, DOCX, inject(xml0, entries(hs, pages)))
pages2, n2 = page_numbers(DOCX, hs)
print(pages2[:6], pages2[-4:]); print("headings", len(hs), "pages", n2, "stable", pages == pages2, "missing", pages2.count("—"))
