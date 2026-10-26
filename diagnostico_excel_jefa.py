# -*- coding: utf-8 -*-
"""Diagnóstico SOLO LECTURA del xlsm de la jefa. Streaming (zipfile + iterparse)."""
import argparse, os, re, subprocess, tempfile, zipfile, hashlib
import xml.etree.ElementTree as ET

NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
RNS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
VBA_PATTERNS = [
    r"7100", r"Filial", r"FPA", r"\.Clear\b", r"\.ClearContents", r"\.ClearFormats",
    r"Cells\.Clear", r"PasteSpecial", r"\.Formula\s*=", r"\.Formula2", r"FormulaLocal",
    r"FormulaR1C1", r"\.Delete\b", r"UsedRange", r"\.Style\b", r"NumberFormat",
    r"Destination\s*:?=", r"ScreenUpdating", r"Calculation", r"\.Copy\b", r"\.Paste\b",
    r"Names\.Add", r"XLOOKUP|MAXIFS|MINIFS|\bIFS\(|TEXTJOIN|CONCAT\(|FILTER\(|UNIQUE\(|LET\(",
]


def sha256(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for c in iter(lambda: f.read(1 << 20), b""):
            h.update(c)
    return h.hexdigest()


def col_to_num(ref):
    m = re.match(r"([A-Z]+)(\d+)", ref or "")
    if not m:
        return 0, 0
    n = 0
    for ch in m.group(1):
        n = n * 26 + ord(ch) - 64
    return int(m.group(2)), n


def num_to_col(n):
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def extract_vba(z, out_dir, olevba_exe, log):
    names = [n for n in z.namelist() if n.lower().endswith("vbaproject.bin")]
    if not names:
        log.append("!! No hay vbaProject.bin")
        return None
    tmp = os.path.join(tempfile.gettempdir(), "diag_vbaProject.bin")
    with z.open(names[0]) as src, open(tmp, "wb") as dst:
        for c in iter(lambda: src.read(1 << 20), b""):
            dst.write(c)
    code = subprocess.run([olevba_exe, "-m", "oletools.olevba", "-c", tmp], capture_output=True, text=True,
                          encoding="utf-8", errors="replace").stdout
    full = subprocess.run([olevba_exe, "-m", "oletools.olevba", tmp], capture_output=True, text=True,
                          encoding="utf-8", errors="replace").stdout
    out = os.path.join(out_dir, "vba_extracted.bas")
    open(out, "w", encoding="utf-8").write(code)
    open(os.path.join(out_dir, "olevba_report.txt"), "w", encoding="utf-8").write(full)
    log.append(f"VBA extraído: {out} ({len(code)} chars)")
    return out


def hunt_vba(bas, log):
    if not bas:
        return
    lines = open(bas, encoding="utf-8", errors="replace").read().splitlines()
    log.append("\n## Búsqueda en VBA")
    for pat in VBA_PATTERNS:
        rx = re.compile(pat, re.I)
        hits, module = [], "?"
        for i, ln in enumerate(lines):
            if ln.startswith("VBA MACRO") or ln.startswith("Attribute VB_Name"):
                module = ln.strip()
            if rx.search(ln) and not ln.strip().startswith("'"):
                hits.append((i + 1, module, ln.strip()))
        log.append(f"\n### `{pat}` -> {len(hits)} coincidencias")
        for h in hits[:40]:
            log.append(f"  L{h[0]} [{h[1]}]: {h[2][:160]}")


def audit_workbook_xml(z, log):
    log.append("\n## workbook.xml / definedNames")
    root = ET.fromstring(z.read("xl/workbook.xml"))
    sheets = [(s.get("name"), s.get(RNS + "id"), s.get("state")) for s in root.iter(NS + "sheet")]
    log.append(f"Hojas: {sheets}")
    dn = list(root.iter(NS + "definedName"))
    log.append(f"definedNames: {len(dn)}")
    bad = 0
    for d in dn:
        t = d.text or ""
        if re.search(r"#REF!|#NAME\?|#¡REF|[A-Za-z]:\\|\[\d+\]|\.xl[sx]", t) or "_xlfn" in t:
            bad += 1
            if bad <= 60:
                log.append(f"  ROTO/EXTERNO: {d.get('name')} (hidden={d.get('hidden')}, "
                           f"scope={d.get('localSheetId')}) = {t[:140]}")
    log.append(f"definedNames problemáticos: {bad}")
    ext = [n for n in z.namelist() if n.startswith("xl/externalLinks/") and n.endswith(".xml")]
    log.append(f"externalLinks: {len(ext)}")
    for n in ext:
        rel = n.replace("externalLinks/", "externalLinks/_rels/") + ".rels"
        if rel in z.namelist():
            tg = re.findall(rb'Target="([^"]+)"', z.read(rel))[:3]
            log.append(f"  {n}: {tg}")
    return sheets


def audit_styles(z, log):
    log.append("\n## styles.xml / calcChain")
    for n in ("xl/calcChain.xml", "xl/styles.xml"):
        if n in z.namelist():
            i = z.getinfo(n)
            log.append(f"{n}: {i.file_size/1e6:.2f} MB (comprimido {i.compress_size/1e6:.2f} MB)")
        else:
            log.append(f"{n}: NO EXISTE")
    if "xl/styles.xml" in z.namelist():
        counts = {}
        with z.open("xl/styles.xml") as f:
            for ev, el in ET.iterparse(f, events=("end",)):
                tag = el.tag.replace(NS, "")
                if tag in ("cellXfs", "cellStyleXfs", "cellStyles", "dxfs", "numFmts", "fonts", "fills", "borders"):
                    counts[tag] = f"{el.get('count')} (real {len(list(el))})"
        log.append(f"Conteos: {counts}")


def audit_sheets(z, sheets_info, log):
    log.append("\n## Hojas (streaming)")
    rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    rid2t = {r.get("Id"): r.get("Target") for r in rels}
    for name, rid, state in sheets_info:
        t = rid2t.get(rid, "")
        part = t.lstrip("/") if t.startswith("/") else "xl/" + t
        if part not in z.namelist():
            log.append(f"- {name}: parte no encontrada ({t})")
            continue
        zi = z.getinfo(part)
        dim = None; maxr = maxc = 0; rows = cells = empty_styled = 0
        nf = cf_rules = merges = dv = fx_total = shared = 0
        fx, samples = {}, []
        with z.open(part) as f:
            for ev, el in ET.iterparse(f, events=("end",)):
                tag = el.tag.replace(NS, "")
                if tag == "dimension":
                    dim = el.get("ref")
                elif tag == "c":
                    cells += 1
                    fel = el.find(NS + "f")
                    if fel is not None or el.find(NS + "v") is not None or el.find(NS + "is") is not None:
                        r, c = col_to_num(el.get("r"))
                        maxr = max(maxr, r); maxc = max(maxc, c)
                    else:
                        empty_styled += 1
                    if fel is not None:
                        fx_total += 1
                        txt = fel.text or ""
                        if fel.get("t") == "shared":
                            shared += 1
                        for m in re.findall(r"_x(?:lfn|lws|lpm)\.[A-Za-z0-9_.]+", txt):
                            fx[m] = fx.get(m, 0) + 1
                        if re.search(r"#REF!|#NAME\?|\[\d+\]|[A-Za-z]:\\", txt) and len(samples) < 5:
                            samples.append((el.get("r"), txt[:120]))
                    el.clear()
                elif tag == "row":
                    rows += 1; el.clear()
                elif tag == "conditionalFormatting":
                    nf += 1
                elif tag == "cfRule":
                    cf_rules += 1
                elif tag == "mergeCell":
                    merges += 1
                elif tag == "dataValidation":
                    dv += 1
        last = f"{num_to_col(maxc)}{maxr}" if maxc else "vacía"
        log.append(f"- **{name}** [{state or 'visible'}] {part} {zi.file_size/1e6:.2f} MB | dimension={dim} | "
                   f"último real={last} | filas={rows} celdas={cells} (vacías con estilo={empty_styled}) | "
                   f"fórmulas={fx_total} (shared={shared}) | CF bloques={nf} reglas={cf_rules} | merges={merges} dv={dv}")
        if fx:
            log.append(f"    funciones _xlfn/_xlws: {sorted(fx.items(), key=lambda x: -x[1])[:15]}")
        for s in samples:
            log.append(f"    FÓRMULA ROTA {s}")


def audit_parts(z, log):
    log.append("\n## Top 15 partes del ZIP por tamaño comprimido")
    for i in sorted(z.infolist(), key=lambda x: -x.compress_size)[:15]:
        log.append(f"  {i.filename}: {i.compress_size/1e6:.2f} MB (descomp. {i.file_size/1e6:.2f} MB)")


def run(path, out_dir, olevba_exe, full=True):
    log = [f"# Auditoría: {os.path.basename(path)}",
           f"Tamaño: {os.path.getsize(path)/1e6:.2f} MB | SHA256: {sha256(path)}"]
    with zipfile.ZipFile(path) as z:
        if full:
            hunt_vba(extract_vba(z, out_dir, olevba_exe, log), log)
        sheets = audit_workbook_xml(z, log)
        audit_styles(z, log)
        audit_sheets(z, sheets, log)
        audit_parts(z, log)
    return "\n".join(log)


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--xlsm", required=True)
    ap.add_argument("--out", default="diagnostico_excel")
    ap.add_argument("--olevba", required=True)
    ap.add_argument("--compare", nargs="*", default=[])
    a = ap.parse_args()
    os.makedirs(a.out, exist_ok=True)
    h0 = sha256(a.xlsm)
    open(os.path.join(a.out, "diagnostico_principal.txt"), "w", encoding="utf-8").write(
        run(a.xlsm, a.out, a.olevba))
    for p in a.compare:
        try:
            r = run(p, a.out, a.olevba, full=False)
        except Exception as e:
            r = f"ERROR {p}: {e}"
        open(os.path.join(a.out, "diag_" + os.path.basename(p)[:40] + ".txt"), "w", encoding="utf-8").write(r)
    assert sha256(a.xlsm) == h0, "El archivo cambió"
    try:
        import psutil
        print("Pico RAM (MB):", psutil.Process().memory_info().peak_wset / 1e6)
    except Exception:
        pass
    print("OK. SHA256 intacto:", h0)
