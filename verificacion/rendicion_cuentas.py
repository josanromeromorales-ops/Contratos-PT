# -*- coding: utf-8 -*-
"""Descarga el listado de contratos de Pinto publicado en rendiciondecuentas.es
(Tribunal de Cuentas) para varios ejercicios y lo guarda en CSV.
Uso: python verificacion/rendicion_cuentas.py 2020 2025 salida.csv"""
import csv, html, re, sys, time, urllib.parse, urllib.request
from http.cookiejar import CookieJar

BASE = "https://www.rendiciondecuentas.es/es/consultadeentidadesycuentas/contratos/consultarEntidadContratos.html"
PARAMS = dict(idEntidadPpal=4383, idEntidad=4383, idTipoEntidad="A", tipoEntidad="Ayuntamiento",
              nombreEntidad="Pinto", nifEntidad="P2811300I")
opener = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(CookieJar()))
opener.addheaders = [("User-Agent", "Mozilla/5.0 (verificacion Ayto Pinto)")]


def get(url):
    for i in range(5):
        try:
            with opener.open(url, timeout=60) as r:
                return r.read().decode(r.headers.get_content_charset() or "utf-8", "replace")
        except Exception as e:
            print("  reintento", i + 1, e, flush=True)
            time.sleep(5 * (i + 1))
    raise RuntimeError(url)


def limpia(s):
    return re.sub(r"\s+", " ", html.unescape(re.sub(r"<[^>]+>", " ", s))).strip()


def filas(pagina):
    out = []
    for tr in re.findall(r"<tr[^>]*>(.*?)</tr>", pagina, re.S | re.I):
        tds = re.findall(r"<td[^>]*>(.*?)</td>", tr, re.S | re.I)
        if len(tds) < 4:
            continue
        m = re.search(r'href="([^"]*idEvl=[^"]*)"', tds[0])
        enlace = html.unescape(m.group(1)) if m else ""
        idevl = urllib.parse.unquote(re.search(r"idEvl=([^&]+)", enlace).group(1)) if enlace else ""
        out.append([limpia(t) for t in tds[:4]] + [idevl, enlace])
    return out


def main(desde, hasta, salida):
    with open(salida, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["ejercicio", "expediente", "tipo_contratacion", "objeto", "importe_adjudicacion", "idEvl", "enlace"])
        for anio in range(desde, hasta + 1):
            p = dict(PARAMS, ejercicio=anio)
            primera = get(BASE + "?" + urllib.parse.urlencode(p))
            m = re.search(r"Se han encontrado\s+([\d.]+)\s+elementos", primera)
            total = int(m.group(1).replace(".", "")) if m else 0
            ids = re.findall(r"d-(\d+)-p=", primera)
            tabla = ids[0] if ids else "2677838"
            paginas = -(-total // 10)
            print(f"{anio}: {total} contratos, {paginas} páginas", flush=True)
            n = 0
            for pg in range(1, paginas + 1):
                txt = primera if pg == 1 else get(BASE + "?" + urllib.parse.urlencode(dict(p, **{f"d-{tabla}-p": pg})))
                fs = filas(txt)
                for fl in fs:
                    w.writerow([anio] + fl)
                n += len(fs)
                time.sleep(0.5)
            print(f"   leídas {n} filas", flush=True)


if __name__ == "__main__":
    main(int(sys.argv[1]), int(sys.argv[2]), sys.argv[3])
