# -*- coding: utf-8 -*-
"""
pinto_contratos.py
==================
Descarga los datos abiertos de la Plataforma de Contratación del Sector Público
(Ministerio de Hacienda), se queda con los expedientes de los órganos de
contratación del Ayuntamiento de Pinto y genera / actualiza un Excel.

Fuentes (formato ATOM/XML CODICE, actualización diaria):
  - Licitaciones de perfiles alojados en la Plataforma (sin menores)  -> sindicacion_643
  - Contratos menores                                                 -> sindicacion_1143

USO (desde una ventana de comandos, en la carpeta del programa):

  1) Carga inicial del histórico (una sola vez; puede tardar bastante):
       python pinto_contratos.py --historico 2018

  2) Actualización diaria (lo que se programa en el Programador de tareas):
       python pinto_contratos.py

  3) Si se han descargado los ZIP a mano (por ejemplo, con el navegador):
       python pinto_contratos.py --zip C:\\ruta\\a\\carpeta_con_zips

Opciones: --excel RUTA (por defecto contratos_pinto.xlsx)
          --fuentes licitaciones,menores (por defecto ambas)
          --desde AAAA-MM-DD (para la actualización: leer el feed hasta esa fecha)

Requisitos: Python 3.9 o superior y el paquete openpyxl  (pip install openpyxl)
"""

import argparse
import datetime as dt
import io
import json
import os
import re
import sys
import time
import unicodedata
import urllib.error
import urllib.request
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

# ---------------------------------------------------------------------------
# CONFIGURACIÓN (se puede ajustar sin tocar el resto)
# ---------------------------------------------------------------------------

# Filtro por nombre: se aplica al nombre del órgano y a toda su jerarquía
# (órgano superior). Sin tildes ni mayúsculas.
FILTRO_NOMBRE = r"\bpinto\b"

# Filtro por identificadores. Si se rellena (ID de la Plataforma, DIR3 o NIF),
# se usa ESTE filtro en lugar del de nombre. Los identificadores aparecen en la
# hoja "Órganos" del Excel después de la primera ejecución.
# Ejemplo: IDS_ORGANOS = {"L01281132", "P2811300A"}
IDS_ORGANOS = set()

# Órganos a excluir aunque contengan "Pinto" (por nombre, sin tildes, en minúscula)
# (basta con que el nombre del órgano CONTENGA el texto; sin tildes, en minúscula)
EXCLUIR_NOMBRES = {"mercedes pinto"}   # CIS Mercedes Pinto (Tenerife), Ministerio del Interior

BASE = "https://contrataciondelsectorpublico.gob.es/sindicacion"
FUENTES = {
    "licitaciones": {
        "etiqueta": "Licitación",
        "feed": BASE + "/sindicacion_643/licitacionesPerfilesContratanteCompleto3.atom",
        "zip": BASE + "/sindicacion_643/licitacionesPerfilesContratanteCompleto3_{p}.zip",
    },
    "menores": {
        "etiqueta": "Contrato menor",
        "feed": BASE + "/sindicacion_1143/contratosMenoresPerfilesContratantes.atom",
        "zip": BASE + "/sindicacion_1143/contratosMenoresPerfilesContratantes_{p}.zip",
    },
}

CARPETA = Path(__file__).resolve().parent
CARPETA_DESCARGAS = CARPETA / "descargas"
FICHERO_DATOS = CARPETA / "datos_contratos_pinto.json"
EXCEL_POR_DEFECTO = CARPETA / "contratos_pinto.xlsx"
MAX_PAGINAS_FEED = 400          # tope de seguridad al seguir los enlaces "next"
MARGEN_DIAS = 1                 # se relee un pequeño solape para no perder nada
DIAS_PRIMERA_VEZ = 1            # primera actualización sin datos previos: días hacia atrás
MAX_MINUTOS_FEED = 60           # tope de tiempo leyendo el feed de cada fuente
BORRAR_DESCARGAS = False       # True: borra cada ZIP tras leerlo (ahorra disco, p. ej. en GitHub)

# ---------------------------------------------------------------------------
# Tablas de códigos CODICE
# ---------------------------------------------------------------------------

ESTADOS = {
    "PRE": "Anuncio previo", "PUB": "En plazo", "EV": "Pendiente de adjudicación",
    "ADJ": "Adjudicada", "RES": "Resuelta", "ANUL": "Anulada",
}
TIPOS = {
    "1": "Suministros", "2": "Servicios", "3": "Obras",
    "21": "Gestión de servicios públicos", "22": "Concesión de servicios",
    "31": "Concesión de obras públicas", "32": "Concesión de obras",
    "40": "Colaboración público-privada", "7": "Administrativo especial",
    "8": "Privado", "50": "Patrimonial", "999": "Otros",
}
PROCEDIMIENTOS = {
    "1": "Abierto", "2": "Restringido", "3": "Negociado sin publicidad",
    "4": "Negociado con publicidad", "5": "Diálogo competitivo", "6": "Contrato menor",
    "7": "Derivado de acuerdo marco", "8": "Concurso de proyectos",
    "9": "Abierto simplificado", "10": "Asociación para la innovación",
    "11": "Derivado de asociación para la innovación",
    "12": "Basado en sistema dinámico de adquisición",
    "13": "Licitación con negociación", "100": "Normas internas", "999": "Otros",
}
RESULTADOS = {
    "1": "Adjudicado provisionalmente", "2": "Adjudicado definitivamente",
    "3": "Desierto", "4": "Desistimiento", "5": "Renuncia",
    "8": "Adjudicado", "9": "Formalizado",
}


def codigo(tabla, valor):
    if not valor:
        return ""
    return tabla.get(valor, valor)


# ---------------------------------------------------------------------------
# Utilidades XML (sin depender de los prefijos de espacio de nombres)
# ---------------------------------------------------------------------------

def P(*tags):
    return "/".join("{*}" + t for t in tags)


def txt(el, *tags):
    if el is None:
        return ""
    x = el.find(P(*tags)) if tags else el
    return (x.text or "").strip() if x is not None and x.text else ""


def num(s):
    try:
        return float(s) if s not in ("", None) else None
    except ValueError:
        return None


def fecha(s):
    """'2026-10-01', '2026-10-01+02:00' o '2026-10-01T10:00:00...' -> 'AAAA-MM-DD'."""
    return s[:10] if s and re.match(r"\d{4}-\d{2}-\d{2}", s) else ""


def normaliza(s):
    s = unicodedata.normalize("NFKD", s or "")
    return "".join(c for c in s if not unicodedata.combining(c)).lower().strip()


# ---------------------------------------------------------------------------
# Lectura de una entrada del feed
# ---------------------------------------------------------------------------

def lee_entrada(entry, origen):
    cfs = entry.find(P("ContractFolderStatus"))
    if cfs is None:
        return None
    lcp = cfs.find(P("LocatedContractingParty"))
    party = lcp.find(P("Party")) if lcp is not None else None

    ids = {}
    if party is not None:
        for pid in party.findall(P("PartyIdentification", "ID")):
            ids[(pid.get("schemeName") or "ID").upper()] = (pid.text or "").strip()

    jerarquia = []
    padre = lcp.find(P("ParentLocatedParty")) if lcp is not None else None
    while padre is not None:
        nombre = txt(padre, "PartyName", "Name")
        if nombre:
            jerarquia.append(nombre)
        padre = padre.find(P("ParentLocatedParty"))

    proj = cfs.find(P("ProcurementProject"))
    budget = proj.find(P("BudgetAmount")) if proj is not None else None
    dur = proj.find(P("PlannedPeriod", "DurationMeasure")) if proj is not None else None
    cpvs = [(c.text or "").strip() for c in proj.findall(P("RequiredCommodityClassification", "ItemClassificationCode"))] if proj is not None else []
    tp = cfs.find(P("TenderingProcess"))

    publicaciones = sorted(fecha(x.text or "")
                           for vni in cfs.iter() if vni.tag.endswith("}ValidNoticeInfo")
                           for x in vni.iter() if x.tag.endswith("}IssueDate") and fecha(x.text or ""))

    adjudicaciones = []
    for tr in cfs.findall(P("TenderResult")):
        win = tr.find(P("WinningParty"))
        atp = tr.find(P("AwardedTenderedProject"))
        adjudicaciones.append({
            "Lote": txt(atp, "ProcurementProjectLotID"),
            "Resultado": codigo(RESULTADOS, txt(tr, "ResultCode")),
            "Adjudicatario": txt(win, "PartyName", "Name"),
            "NIF adjudicatario": txt(win, "PartyIdentification", "ID"),
            "Importe adjudicación sin IVA": num(txt(atp, "LegalMonetaryTotal", "TaxExclusiveAmount")),
            "Importe adjudicación con IVA": num(txt(atp, "LegalMonetaryTotal", "PayableAmount")),
            "Fecha adjudicación": fecha(txt(tr, "AwardDate")),
            "Fecha formalización": fecha(txt(tr, "Contract", "IssueDate")),
            "Nº ofertas": num(txt(tr, "ReceivedTenderQuantity")),
        })

    link = entry.find(P("link"))
    unidad = {"DAY": "días", "MON": "meses", "ANN": "años"}.get((dur.get("unitCode") if dur is not None else "") or "", "")
    return {
        "id": txt(entry, "id"),
        "Origen": origen,
        "Expediente": txt(cfs, "ContractFolderID"),
        "Objeto": txt(proj, "Name") or txt(entry, "title"),
        "Órgano de contratación": txt(party, "PartyName", "Name"),
        "Jerarquía del órgano": " > ".join(reversed(jerarquia)),
        "ID órgano (Plataforma)": ids.get("ID_PLATAFORMA", ""),
        "DIR3 órgano": ids.get("DIR3", ""),
        "NIF órgano": ids.get("NIF", ""),
        "Estado": codigo(ESTADOS, txt(cfs, "ContractFolderStatusCode")),
        "Tipo de contrato": codigo(TIPOS, txt(proj, "TypeCode")),
        "Procedimiento": codigo(PROCEDIMIENTOS, txt(tp, "ProcedureCode")) or ("Contrato menor" if origen == "Contrato menor" else ""),
        "Presupuesto base sin IVA": num(txt(budget, "TaxExclusiveAmount")),
        "Presupuesto base con IVA": num(txt(budget, "TotalAmount")),
        "Valor estimado": num(txt(budget, "EstimatedOverallContractAmount")),
        "CPV": ", ".join(c for c in cpvs if c),
        "Duración": (txt(dur) + " " + unidad).strip() if dur is not None else "",
        "Fin plazo ofertas": fecha(txt(tp, "TenderSubmissionDeadlinePeriod", "EndDate")),
        "Primera publicación": publicaciones[0] if publicaciones else "",
        "Última actualización": txt(entry, "updated"),
        "Enlace": link.get("href", "") if link is not None else "",
        "Borrado en Plataforma": "",
        "adjudicaciones": adjudicaciones,
    }


def candidata(entry):
    """Filtro previo barato: ¿algún nombre o identificador del órgano puede ser de Pinto?"""
    lcp = entry.find(P("ContractFolderStatus", "LocatedContractingParty"))
    if lcp is None:
        return False
    if IDS_ORGANOS:
        return any((x.text or "").strip() in IDS_ORGANOS
                   for x in lcp.iter() if x.tag.endswith("}ID"))
    return any(re.search(FILTRO_NOMBRE, normaliza(x.text or ""))
               for x in lcp.iter() if x.tag.endswith("}Name"))


def es_de_pinto(reg):
    nombres = [reg["Órgano de contratación"]] + reg["Jerarquía del órgano"].split(" > ")
    if any(x in normaliza(reg["Órgano de contratación"]) for x in EXCLUIR_NOMBRES):
        return False
    if IDS_ORGANOS:
        return bool(IDS_ORGANOS & {reg["ID órgano (Plataforma)"], reg["DIR3 órgano"], reg["NIF órgano"]})
    return any(re.search(FILTRO_NOMBRE, normaliza(n)) for n in nombres if n)


# ---------------------------------------------------------------------------
# Almacén (JSON con la última versión de cada expediente)
# ---------------------------------------------------------------------------

class Almacen:
    def __init__(self, ruta):
        self.ruta = ruta
        self.registros, self.control = {}, {}
        if ruta.exists():
            d = json.loads(ruta.read_text(encoding="utf-8"))
            self.registros, self.control = d.get("registros", {}), d.get("control", {})
        self.nuevos = self.actualizados = self.borrados = 0
        self.pendientes = {}

    def guarda_registro(self, reg):
        prev = self.registros.get(reg["id"])
        if prev is None:
            self.nuevos += 1
        elif reg["Última actualización"] > prev["Última actualización"]:
            self.actualizados += 1
            reg["Borrado en Plataforma"] = prev.get("Borrado en Plataforma", "")
        else:
            return
        self.registros[reg["id"]] = reg

    def marca_borrado(self, ref, cuando):
        # Los borrados se anotan y se aplican al final, porque en el feed suelen
        # aparecer antes que la entrada a la que se refieren.
        self.pendientes[ref] = cuando

    def aplica_borrados(self):
        for ref, cuando in list(self.pendientes.items()):
            reg = self.registros.get(ref)
            if reg is None:
                continue
            if not reg.get("Borrado en Plataforma") and reg["Última actualización"] <= cuando:
                reg["Borrado en Plataforma"] = fecha(cuando) or cuando
                self.borrados += 1
            del self.pendientes[ref]
        self.pendientes = {}  # los que no son de Pinto no interesan

    def guarda(self):
        self.aplica_borrados()
        tmp = self.ruta.with_suffix(".tmp")
        # Con sangría y claves ordenadas, para que en GitHub los cambios diarios ocupen poco
        tmp.write_text(json.dumps({"control": self.control, "registros": self.registros},
                                  ensure_ascii=False, indent=1, sort_keys=True), encoding="utf-8")
        os.replace(tmp, self.ruta)


# ---------------------------------------------------------------------------
# Procesado de ficheros ATOM
# ---------------------------------------------------------------------------

def procesa_atom(fobj, origen, almacen):
    """Lee un fichero ATOM. Devuelve (enlace_next, fecha_mas_antigua, nº_entradas, fecha_mas_reciente)."""
    siguiente, mas_antigua, mas_reciente, n = None, None, "", 0
    for _, el in ET.iterparse(fobj, events=("end",)):
        tag = el.tag.rsplit("}", 1)[-1]
        if tag == "link" and el.get("rel") == "next":
            siguiente = el.get("href")
        elif tag == "entry":
            n += 1
            upd = txt(el, "updated")
            if upd and (mas_antigua is None or upd < mas_antigua):
                mas_antigua = upd
            mas_reciente = max(mas_reciente, upd)
            # comprobación rápida del órgano antes de leer la entrada completa
            if candidata(el):
                reg = lee_entrada(el, origen)
                if reg and es_de_pinto(reg):
                    almacen.guarda_registro(reg)
            el.clear()
        elif tag == "deleted-entry":
            almacen.marca_borrado(el.get("ref", ""), el.get("when", ""))
            el.clear()
    return siguiente, mas_antigua, n, mas_reciente


def procesa_zip(ruta, origen, almacen):
    total = 0
    with zipfile.ZipFile(ruta) as zf:
        for nombre in zf.namelist():
            if nombre.lower().endswith(".atom"):
                with zf.open(nombre) as f:
                    total += procesa_atom(f, origen, almacen)[2]
    print(f"   {ruta.name}: {total} entradas leídas · expedientes de Pinto acumulados: {len(almacen.registros)}", flush=True)


# ---------------------------------------------------------------------------
# Descargas
# ---------------------------------------------------------------------------

CABECERAS = {"User-Agent": "Ayto-Pinto-GobiernoAbierto/1.0"}
PARTES_DESCARGA = 8             # conexiones simultáneas para los ZIP grandes
MB = 1 << 20


def _abre(url, rango=None):
    h = dict(CABECERAS)
    if rango:
        h["Range"] = "bytes=%d-%d" % rango
    return urllib.request.urlopen(urllib.request.Request(url, headers=h), timeout=120)


def _tamano_y_rangos(url):
    """Devuelve (tamaño, admite_rangos) o (None, False). 404 -> ('404', False)."""
    try:
        with _abre(url, (0, 0)) as r:
            cr = r.headers.get("Content-Range", "")
            if r.status == 206 and "/" in cr and cr.rsplit("/", 1)[1].isdigit():
                return int(cr.rsplit("/", 1)[1]), True
            return int(r.headers.get("Content-Length") or 0) or None, False
    except urllib.error.HTTPError as e:
        if e.code == 404:
            return "404", False
        return None, False
    except Exception:
        return None, False


class _Progreso:
    def __init__(self, total, nombre):
        import threading
        self.total, self.nombre, self.hecho = total, nombre, 0
        self.t0 = self.ultimo = time.time()
        self.lock = threading.Lock()

    def suma(self, n):
        with self.lock:
            self.hecho += n
            ahora = time.time()
            if ahora - self.ultimo >= 60:
                self.ultimo = ahora
                vel = self.hecho / MB / max(ahora - self.t0, 1)
                tot = f" de {self.total / MB:.0f}" if self.total else ""
                print(f"      {self.nombre}: {self.hecho / MB:.0f}{tot} MB · {vel:.1f} MB/s", flush=True)


def _descarga_parte(url, fichero, inicio, fin, progreso, intentos=6):
    pos = inicio
    for i in range(intentos):
        try:
            with _abre(url, (pos, fin)) as r, open(fichero, "r+b") as f:
                f.seek(pos)
                while pos <= fin:
                    bloque = r.read(MB)
                    if not bloque:
                        break
                    f.write(bloque)
                    pos += len(bloque)
                    progreso.suma(len(bloque))
            if pos > fin:
                return
        except Exception:
            pass
        time.sleep(5 * (i + 1))   # se reanuda desde donde se quedó
    raise RuntimeError(f"parte {inicio}-{fin} incompleta")


def descarga(url, destino=None, intentos=3):
    # Ficheros grandes: varias conexiones a la vez (si el servidor lo admite)
    if destino is not None:
        tam, rangos = _tamano_y_rangos(url)
        if tam == "404":
            return None
        if rangos and tam and tam > 20 * MB:
            from concurrent.futures import ThreadPoolExecutor
            print(f"      {tam / MB:.0f} MB, en {PARTES_DESCARGA} conexiones", flush=True)
            tmp = destino.with_suffix(".part")
            with open(tmp, "wb") as f:
                f.truncate(tam)
            trozo = -(-tam // PARTES_DESCARGA)
            prog = _Progreso(tam, destino.name)
            with ThreadPoolExecutor(PARTES_DESCARGA) as ex:
                tareas = [ex.submit(_descarga_parte, url, tmp, a, min(a + trozo, tam) - 1, prog)
                          for a in range(0, tam, trozo)]
                for t in tareas:
                    t.result()
            os.replace(tmp, destino)
            return destino
    req = urllib.request.Request(url, headers=CABECERAS)
    for i in range(intentos):
        try:
            with urllib.request.urlopen(req, timeout=120) as r:
                if destino is None:
                    return r.read()
                tmp = destino.with_suffix(".part")
                prog = _Progreso(int(r.headers.get("Content-Length") or 0), destino.name)
                with open(tmp, "wb") as f:
                    while True:
                        bloque = r.read(MB)
                        if not bloque:
                            break
                        f.write(bloque)
                        prog.suma(len(bloque))
                os.replace(tmp, destino)
                return destino
        except urllib.error.HTTPError as e:
            if e.code == 404:
                return None
            err = e
        except Exception as e:   # red, timeout...
            err = e
        time.sleep(5 * (i + 1))
    raise RuntimeError(f"No se pudo descargar {url}: {err}")


def lista_periodos(desde_anio):
    """[[año, mes1..mes12], ...] para años cerrados y [[mes]] para el año en curso."""
    hoy = dt.date.today()
    periodos = [[str(a)] + [f"{a}{m:02d}" for m in range(1, 13)] for a in range(desde_anio, hoy.year)]
    # meses del año en curso, incluido el actual (Hacienda lo publica hasta el día anterior)
    periodos += [[f"{hoy.year}{m:02d}"] for m in range(1, hoy.month + 1)]
    return periodos


def plan_paralelo(desde_anio, fuentes, almacen):
    """Tareas pendientes para cargar en paralelo (una por fuente y año/mes)."""
    tareas = []
    for clave in fuentes:
        hechos = set(almacen.control.get(clave, {}).get("periodos_cargados", []))
        for opciones in lista_periodos(desde_anio):
            anio = opciones[0][:4]
            if anio in hechos:          # año ya cargado entero con el ZIP anual
                continue
            # una tarea por mes: un año completo de licitaciones (~3 GB) no cabe
            # en las 6 h de un servidor de GitHub a la velocidad que da Hacienda
            for mes in (opciones[1:] if len(opciones) > 1 else opciones):
                if mes not in hechos:
                    tareas.append({"fuente": clave, "periodo": mes})
    return tareas


def fusiona(carpeta, almacen):
    """Une los resultados parciales (JSON) de las tareas en paralelo."""
    for ruta in sorted(Path(carpeta).rglob("*.json")):
        d = json.loads(ruta.read_text(encoding="utf-8"))
        for reg in d.get("registros", {}).values():
            almacen.guarda_registro(reg)
            if reg.get("Borrado en Plataforma"):
                almacen.registros[reg["id"]]["Borrado en Plataforma"] = reg["Borrado en Plataforma"]
        for clave, ctl in d.get("control", {}).items():
            mio = almacen.control.setdefault(clave, {})
            mio["periodos_cargados"] = sorted(set(mio.get("periodos_cargados", [])) | set(ctl.get("periodos_cargados", [])))
            if ctl.get("historico_desde") and (not mio.get("historico_desde") or ctl["historico_desde"] < mio["historico_desde"]):
                mio["historico_desde"] = ctl["historico_desde"]
        print(f"   {ruta.name}: {len(d.get('registros', {}))} expedientes", flush=True)


def carga_historico(desde_anio, fuentes, almacen, solo=None):
    hoy = dt.date.today()
    CARPETA_DESCARGAS.mkdir(exist_ok=True)
    for clave in fuentes:
        f = FUENTES[clave]
        print(f"\n== Histórico de {clave} desde {desde_anio} ==")
        periodos = lista_periodos(desde_anio)
        if solo:
            # un año o un mes concreto (los meses de años pasados también se publican sueltos)
            periodos = [o for o in periodos if o[0] in solo] or [[p] for p in sorted(solo)]
        mes_actual = f"{hoy.year}{hoy.month:02d}"
        ctl = almacen.control.setdefault(clave, {})
        hechos = set(ctl.get("periodos_cargados", []))
        for opciones in periodos:
            # primero el ZIP anual; si no existe, los mensuales de ese año
            candidatos = opciones[:1]
            for p in candidatos:
                if p in hechos:
                    print(f"   {p}: ya cargado en una ejecución anterior", flush=True)
                    continue
                destino = CARPETA_DESCARGAS / f"{clave}_{p}.zip"
                if not destino.exists():
                    print(f"   descargando {p}...", flush=True)
                    t0 = time.time()
                    if descarga(f["zip"].format(p=p), destino) is None:
                        if len(opciones) > 1 and p == opciones[0]:
                            candidatos.extend(opciones[1:])
                        else:
                            print(f"   (no publicado: {p})", flush=True)
                        continue
                    print(f"   {p}: {destino.stat().st_size / 1e6:.0f} MB en {time.time() - t0:.0f} s", flush=True)
                procesa_zip(destino, f["etiqueta"], almacen)
                if BORRAR_DESCARGAS:
                    destino.unlink(missing_ok=True)
                # se guarda el avance tras cada periodo: si la ejecución se corta,
                # la siguiente continúa donde se quedó
                if p != mes_actual:   # el mes en curso aún está incompleto
                    hechos.add(p)
                ctl["periodos_cargados"] = sorted(hechos)
                almacen.guarda()
        if not solo and (ctl.get("historico_desde") is None or desde_anio < ctl["historico_desde"]):
            ctl["historico_desde"] = desde_anio


def carga_zips_locales(ruta, fuentes, almacen):
    ruta = Path(ruta)
    zips = sorted(ruta.glob("*.zip")) if ruta.is_dir() else [ruta]
    for z in zips:
        n = z.name.lower()
        clave = "menores" if "menor" in n else "licitaciones"
        if clave in fuentes:
            procesa_zip(z, FUENTES[clave]["etiqueta"], almacen)


def actualiza_feed(fuentes, almacen, desde=None):
    for clave in fuentes:
        f = FUENTES[clave]
        ctl = almacen.control.setdefault(clave, {})
        if desde:
            limite = desde
        elif ctl.get("ultima_entrada"):
            limite = (dt.date.fromisoformat(ctl["ultima_entrada"][:10]) - dt.timedelta(days=MARGEN_DIAS)).isoformat()
        else:
            # sin datos previos: solo los últimos días (lo anterior llega con --historico)
            limite = (dt.date.today() - dt.timedelta(days=DIAS_PRIMERA_VEZ)).isoformat()
        print(f"\n== Actualización de {clave} (hasta {limite}) ==", flush=True)
        url, paginas, mas_reciente = f["feed"], 0, ctl.get("ultima_entrada", "")
        t0, completo = time.time(), False
        while url and paginas < MAX_PAGINAS_FEED:
            if time.time() - t0 > MAX_MINUTOS_FEED * 60:
                print(f"   se alcanzó el tope de {MAX_MINUTOS_FEED} min; se guarda lo leído", flush=True)
                break
            datos = descarga(url)
            if datos is None:
                break
            paginas += 1
            url, mas_antigua, n, reciente = procesa_atom(io.BytesIO(datos), f["etiqueta"], almacen)
            mas_reciente = max(mas_reciente, reciente)
            print(f"   página {paginas}: hasta {(mas_antigua or '')[:10]} · "
                  f"expedientes de Pinto acumulados: {len(almacen.registros)}", flush=True)
            if not n or not url or (mas_antigua and mas_antigua[:10] < limite):
                completo = True
                break
        print(f"   {paginas} páginas leídas en {time.time() - t0:.0f} s", flush=True)
        # solo se avanza la marca si se llegó hasta el límite; si no, la próxima vez se relee
        if completo or not ctl.get("ultima_entrada"):
            ctl["ultima_entrada"] = mas_reciente
        ctl["ultima_actualizacion"] = dt.datetime.now().isoformat(timespec="seconds")


# ---------------------------------------------------------------------------
# Excel
# ---------------------------------------------------------------------------

COLUMNAS = [
    ("Origen", 14, None), ("Expediente", 18, None), ("Objeto", 60, None),
    ("Órgano de contratación", 38, None), ("Estado", 22, None),
    ("Tipo de contrato", 16, None), ("Procedimiento", 22, None),
    ("Presupuesto base sin IVA", 16, "eur"), ("Presupuesto base con IVA", 16, "eur"),
    ("Valor estimado", 16, "eur"), ("Adjudicatario", 36, None), ("NIF adjudicatario", 14, None),
    ("Importe adjudicación sin IVA", 16, "eur"), ("Importe adjudicación con IVA", 16, "eur"),
    ("Fecha adjudicación", 13, "fecha"), ("Fecha formalización", 13, "fecha"),
    ("Nº ofertas", 9, "int"), ("Primera publicación", 13, "fecha"), ("Año", 7, "año"),
    ("CPV", 14, None), ("Duración", 11, None), ("Fin plazo ofertas", 13, "fecha"),
    ("Última actualización", 17, "fechahora"), ("Borrado en Plataforma", 13, "fecha"),
    ("ID órgano (Plataforma)", 16, None), ("DIR3 órgano", 13, None),
    ("Jerarquía del órgano", 40, None), ("Enlace", 16, "link"),
]
COLUMNAS_ADJ = [
    ("Expediente", 18, None), ("Origen", 14, None), ("Órgano de contratación", 38, None),
    ("Objeto", 50, None), ("Lote", 8, None), ("Resultado", 20, None),
    ("Adjudicatario", 36, None), ("NIF adjudicatario", 14, None),
    ("Importe adjudicación sin IVA", 16, "eur"), ("Importe adjudicación con IVA", 16, "eur"),
    ("Fecha adjudicación", 13, "fecha"), ("Fecha formalización", 13, "fecha"),
    ("Nº ofertas", 9, "int"), ("Enlace", 16, "link"),
]


def a_fecha(s, con_hora=False):
    if not s:
        return None
    try:
        if con_hora:
            return dt.datetime.fromisoformat(s[:19])
        return dt.date.fromisoformat(s[:10])
    except ValueError:
        return s


def fila_contrato(r):
    adj = r["adjudicaciones"]
    validas = [a for a in adj if a["Adjudicatario"]]
    suma = lambda k: sum(a[k] for a in validas if a[k] is not None) if any(a[k] is not None for a in validas) else None
    primera = lambda k: next((a[k] for a in validas if a[k]), "")
    ofertas = [a["Nº ofertas"] for a in adj if a["Nº ofertas"] is not None]
    d = dict(r)
    d["Adjudicatario"] = "; ".join(dict.fromkeys(a["Adjudicatario"] for a in validas))
    d["NIF adjudicatario"] = "; ".join(dict.fromkeys(a["NIF adjudicatario"] for a in validas if a["NIF adjudicatario"]))
    d["Importe adjudicación sin IVA"] = suma("Importe adjudicación sin IVA")
    d["Importe adjudicación con IVA"] = suma("Importe adjudicación con IVA")
    d["Fecha adjudicación"] = primera("Fecha adjudicación")
    d["Fecha formalización"] = primera("Fecha formalización")
    d["Nº ofertas"] = max(ofertas) if ofertas else None
    base = d["Fecha adjudicación"] or r["Primera publicación"] or r["Última actualización"]
    d["Año"] = int(base[:4]) if base else None
    return d


def escribe_excel(almacen, ruta_excel, filtro_texto):
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.table import Table, TableStyleInfo

    FUENTE = "Arial"
    cab_font = Font(name=FUENTE, bold=True, color="FFFFFF")
    cab_fill = PatternFill("solid", fgColor="1F4E78")
    base_font = Font(name=FUENTE, size=10)
    link_font = Font(name=FUENTE, size=10, color="0563C1", underline="single")
    titulo_font = Font(name=FUENTE, size=14, bold=True, color="1F4E78")
    negrita = Font(name=FUENTE, size=10, bold=True)
    gris = Font(name=FUENTE, size=9, italic=True, color="666666")
    FMT = {"eur": '#,##0.00 €;-#,##0.00 €;"-"', "fecha": "dd/mm/yyyy",
           "fechahora": "dd/mm/yyyy hh:mm", "int": "0", "año": "0"}

    registros = sorted(almacen.registros.values(),
                       key=lambda r: r["Última actualización"], reverse=True)
    filas = [fila_contrato(r) for r in registros]

    wb = Workbook()

    def hoja_tabla(ws, columnas, datos, nombre_tabla):
        for c, (nombre, ancho, _) in enumerate(columnas, 1):
            celda = ws.cell(row=1, column=c, value=nombre)
            celda.font, celda.fill = cab_font, cab_fill
            celda.alignment = Alignment(wrap_text=True, vertical="center")
            ws.column_dimensions[get_column_letter(c)].width = ancho
        ws.row_dimensions[1].height = 32
        for i, d in enumerate(datos, 2):
            for c, (nombre, _, tipo) in enumerate(columnas, 1):
                v = d.get(nombre)
                if tipo == "fecha":
                    v = a_fecha(v)
                elif tipo == "fechahora":
                    v = a_fecha(v, True)
                celda = ws.cell(row=i, column=c)
                if tipo == "link":
                    if v:
                        celda.value, celda.hyperlink, celda.font = "Ver en Plataforma", v, link_font
                    continue
                celda.value = v if v != "" else None
                celda.font = base_font
                if tipo in FMT:
                    celda.number_format = FMT[tipo]
        ultima = max(2, len(datos) + 1)
        ref = f"A1:{get_column_letter(len(columnas))}{ultima}"
        tabla = Table(displayName=nombre_tabla, ref=ref)
        tabla.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
        ws.add_table(tabla)
        ws.freeze_panes = "D2" if nombre_tabla == "TablaContratos" else "B2"
        return ultima

    # --- Contratos
    ws = wb.active
    ws.title = "Contratos"
    ult = hoja_tabla(ws, COLUMNAS, filas, "TablaContratos")
    col = {n: get_column_letter(i) for i, (n, _, _) in enumerate(COLUMNAS, 1)}
    rng = lambda n: f"Contratos!${col[n]}$2:${col[n]}${ult}"

    # --- Adjudicaciones (una fila por lote / adjudicatario)
    ws_a = wb.create_sheet("Adjudicaciones")
    filas_adj = []
    for r in registros:
        for a in r["adjudicaciones"]:
            filas_adj.append({**{k: r[k] for k in ("Expediente", "Origen", "Órgano de contratación", "Objeto", "Enlace")}, **a})
    hoja_tabla(ws_a, COLUMNAS_ADJ, filas_adj, "TablaAdjudicaciones")

    # --- Resumen (fórmulas: se recalcula al abrir y al filtrar/editar)
    ws_r = wb.create_sheet("Resumen", 0)
    ws_r["A1"], ws_r["A1"].font = "Contratación del Ayuntamiento de Pinto", titulo_font
    ws_r["A2"] = f"Datos: Plataforma de Contratación del Sector Público (Ministerio de Hacienda). Actualizado: {dt.datetime.now():%d/%m/%Y %H:%M}"
    ws_r["A2"].font = gris
    for c, w in zip("ABCDEF", (48, 14, 14, 14, 22, 22)):
        ws_r.column_dimensions[c].width = w

    def bloque(fila, titulo, campo, valores):
        cabs = ["Total", "Licitaciones", "Menores", "Importe adjudicado sin IVA", "Importe adjudicado con IVA"]
        ws_r.cell(row=fila, column=1, value=titulo)
        for j, t in enumerate([titulo] + cabs, 1):
            celda = ws_r.cell(row=fila, column=j, value=t)
            celda.font, celda.fill = cab_font, cab_fill
        for i, v in enumerate(valores, fila + 1):
            ws_r.cell(row=i, column=1, value=v).font = base_font
            a = f"$A{i}"
            formulas = [
                f"=COUNTIFS({rng(campo)},{a})",
                f'=COUNTIFS({rng(campo)},{a},{rng("Origen")},"Licitación")',
                f'=COUNTIFS({rng(campo)},{a},{rng("Origen")},"Contrato menor")',
                f'=SUMIFS({rng("Importe adjudicación sin IVA")},{rng(campo)},{a})',
                f'=SUMIFS({rng("Importe adjudicación con IVA")},{rng(campo)},{a})',
            ]
            for j, fml in enumerate(formulas, 2):
                celda = ws_r.cell(row=i, column=j, value=fml)
                celda.font = base_font
                celda.number_format = FMT["eur"] if j >= 5 else "#,##0"
        tot = fila + len(valores) + 1
        ws_r.cell(row=tot, column=1, value="TOTAL").font = negrita
        for j in range(2, 7):
            L = get_column_letter(j)
            celda = ws_r.cell(row=tot, column=j, value=f"=SUM({L}{fila + 1}:{L}{tot - 1})" if valores else 0)
            celda.font = negrita
            celda.number_format = FMT["eur"] if j >= 5 else "#,##0"
        return tot + 2

    organos = sorted({f["Órgano de contratación"] for f in filas if f["Órgano de contratación"]})
    estados = sorted({f["Estado"] for f in filas if f["Estado"]})
    anios = sorted({f["Año"] for f in filas if f["Año"]}, reverse=True)
    fila = bloque(4, "Por órgano de contratación", "Órgano de contratación", organos)
    fila = bloque(fila, "Por estado", "Estado", estados)
    bloque(fila, "Por año", "Año", anios)
    ws_r.freeze_panes = "A4"

    # --- Órganos detectados
    ws_o = wb.create_sheet("Órganos")
    cabs = ["Órgano de contratación", "ID órgano (Plataforma)", "DIR3 órgano", "NIF órgano", "Jerarquía del órgano", "Nº expedientes"]
    for j, (t, w) in enumerate(zip(cabs, (44, 20, 14, 14, 50, 14)), 1):
        celda = ws_o.cell(row=1, column=j, value=t)
        celda.font, celda.fill = cab_font, cab_fill
        ws_o.column_dimensions[get_column_letter(j)].width = w
    vistos = {}
    for r in registros:   # un órgano por nombre, con los identificadores que aparezcan
        v = vistos.setdefault(r["Órgano de contratación"], dict(r))
        for k in ("ID órgano (Plataforma)", "DIR3 órgano", "NIF órgano", "Jerarquía del órgano"):
            v[k] = v.get(k) or r.get(k)
    for i, (nombre, r) in enumerate(sorted(vistos.items()), 2):
        idp = r["ID órgano (Plataforma)"]
        for j, v in enumerate([nombre, idp, r["DIR3 órgano"], r["NIF órgano"], r["Jerarquía del órgano"]], 1):
            ws_o.cell(row=i, column=j, value=v or None).font = base_font
        ws_o.cell(row=i, column=6, value=f"=COUNTIFS({rng('Órgano de contratación')},A{i})").font = base_font
    nota = len(vistos) + 3
    ws_o.cell(row=nota, column=1, value=(
        "Revise esta lista: son los órganos que el filtro ha identificado como del Ayuntamiento de Pinto. "
        "Si sobra o falta alguno, ajuste IDS_ORGANOS o EXCLUIR_NOMBRES al principio del programa.")).font = gris

    # --- Control
    ws_c = wb.create_sheet("Control")
    ws_c.column_dimensions["A"].width = 34
    ws_c.column_dimensions["B"].width = 70
    filas_c = [
        ("Última ejecución", dt.datetime.now().strftime("%d/%m/%Y %H:%M")),
        ("Filtro aplicado", filtro_texto),
        ("Expedientes en el Excel", len(filas)),
        ("Adjudicaciones (filas)", len(filas_adj)),
    ]
    for clave, ctl in almacen.control.items():
        filas_c.append((f"{clave}: entrada más reciente leída", ctl.get("ultima_entrada", "")))
        if ctl.get("historico_desde"):
            filas_c.append((f"{clave}: histórico cargado desde", ctl["historico_desde"]))
    filas_c += [
        ("", ""),
        ("Fuente", "Plataforma de Contratación del Sector Público – datos abiertos (Ministerio de Hacienda)"),
        ("Aviso", "Las hojas Contratos, Adjudicaciones, Órganos y Resumen se regeneran en cada ejecución: "
                  "no escriba en ellas. Para anotaciones propias, use una hoja nueva."),
        ("Aviso", "Una misma licitación se publica varias veces al modificarse; aquí solo figura su última versión."),
    ]
    for i, (a, b) in enumerate(filas_c, 1):
        ws_c.cell(row=i, column=1, value=a).font = negrita
        ws_c.cell(row=i, column=2, value=b).font = base_font
        ws_c.cell(row=i, column=2).alignment = Alignment(wrap_text=True)

    ruta_excel = Path(ruta_excel)
    try:
        wb.save(ruta_excel)
    except PermissionError:
        alternativa = ruta_excel.with_name(f"{ruta_excel.stem}_{dt.datetime.now():%Y%m%d_%H%M}.xlsx")
        wb.save(alternativa)
        print(f"\n!! {ruta_excel.name} está abierto en Excel; se ha guardado como {alternativa.name}")
        ruta_excel = alternativa
    return ruta_excel, len(filas)


# ---------------------------------------------------------------------------

def main():
    ap = argparse.ArgumentParser(description="Contratos del Ayuntamiento de Pinto desde la Plataforma de Contratación")
    ap.add_argument("--historico", type=int, metavar="AÑO", help="carga los ZIP desde ese año")
    ap.add_argument("--zip", metavar="RUTA", help="procesa ZIP ya descargados (fichero o carpeta)")
    ap.add_argument("--desde", metavar="AAAA-MM-DD", help="actualización: leer el feed hasta esa fecha")
    ap.add_argument("--sin-feed", action="store_true", help="no leer el feed en vivo")
    ap.add_argument("--borrar-descargas", action="store_true", help="borra cada ZIP tras procesarlo")
    ap.add_argument("--fuentes", default="licitaciones,menores")
    ap.add_argument("--excel", default=str(EXCEL_POR_DEFECTO))
    ap.add_argument("--datos", default=str(FICHERO_DATOS), help=argparse.SUPPRESS)
    # carga en paralelo (la usa el workflow historico.yml)
    ap.add_argument("--plan", type=int, metavar="AÑO", help="imprime en JSON las tareas pendientes desde ese año")
    ap.add_argument("--periodo", metavar="AAAA|AAAAMM", help="carga solo ese año o mes (con --historico implícito)")
    ap.add_argument("--fusionar", metavar="CARPETA", help="une los JSON parciales de esa carpeta")
    ap.add_argument("--sin-excel", action="store_true")
    a = ap.parse_args()
    global BORRAR_DESCARGAS
    BORRAR_DESCARGAS = BORRAR_DESCARGAS or a.borrar_descargas

    fuentes = [f.strip() for f in a.fuentes.split(",") if f.strip() in FUENTES]
    almacen = Almacen(Path(a.datos))

    # si se ha cambiado el filtro, se quitan los expedientes que ya no cumplen
    fuera = [k for k, r in almacen.registros.items() if not es_de_pinto(r)]
    for k in fuera:
        del almacen.registros[k]
    if fuera:
        print(f"Quitados {len(fuera)} expedientes que ya no cumplen el filtro")
        almacen.guarda()

    if a.plan:
        print(json.dumps(plan_paralelo(a.plan, fuentes, almacen)))
        return

    filtro = f"IDs: {', '.join(sorted(IDS_ORGANOS))}" if IDS_ORGANOS else f"nombre del órgano o superior contiene 'Pinto' (expresión {FILTRO_NOMBRE})"
    print(f"Filtro: {filtro}")

    if a.periodo:
        carga_historico(int(a.periodo[:4]), fuentes, almacen, solo={a.periodo})
        almacen.guarda()
    if a.fusionar:
        fusiona(a.fusionar, almacen)
        almacen.guarda()
    if a.historico:
        carga_historico(a.historico, fuentes, almacen)
        almacen.guarda()
    if a.zip:
        carga_zips_locales(a.zip, fuentes, almacen)
        almacen.guarda()
    if not a.sin_feed:
        actualiza_feed(fuentes, almacen, a.desde)
        almacen.guarda()

    if a.sin_excel:
        print(f"\nHecho: {len(almacen.registros)} expedientes en {a.datos}")
        return
    ruta, n = escribe_excel(almacen, a.excel, filtro)
    print(f"\nHecho: {n} expedientes en {ruta}")
    print(f"  nuevos: {almacen.nuevos} · actualizados: {almacen.actualizados} · marcados como borrados: {almacen.borrados}")


if __name__ == "__main__":
    main()
