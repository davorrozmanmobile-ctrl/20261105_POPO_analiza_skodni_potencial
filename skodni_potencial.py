# -*- coding: utf-8 -*-
"""
Analiza škodnega potenciala za celotno Slovenijo (mreža 50 x 50 m).

Postopek (metodologija, pogl. 3.3):
  Nivo 1: za vsak element ranljivosti se v celico zapiše
            točke     -> vsota ocen (oz. ocena x vrednost polja, npr. št. prebivalcev),
            linije    -> vsota (dolžina v celici [m] x ocena),
            poligoni  -> vsota (površina v celici [m2] x ocena);
          vrednosti v območju nevarnostnega potenciala se razvrstijo (privzeto Quantile, 5 razredov).
  Nivo 2: ponderirana vsota razredov elementov znotraj skupine (GD, IN, KD, OK, SI) -> razvrstitev v 10 razredov.
          ZL se prenese z nivoja 1 (10 razredov).
  Nivo 3: ponderirana vsota skupin -> razvrstitev v 10 razredov (privzeto Natural Breaks).

Vsi parametri se berejo iz input.xlsx (glej README.md).
Celice brez elementov in celice izven maske so NoData (0).

Zagon (ArcGIS Pro Python, potrebuje GDAL, numpy, openpyxl):
  "C:\\Program Files\\ArcGIS\\Pro\\bin\\Python\\envs\\arcgispro-py3\\python.exe" skodni_potencial.py [input.xlsx]
"""
import datetime as dt
import logging
import os
import sys
import time
from collections import Counter, OrderedDict

import numpy as np
import openpyxl
from osgeo import gdal, ogr, osr

gdal.UseExceptions()
ogr.UseExceptions()
osr.UseExceptions()

METODE = ("Quantile", "EqualInterval", "NaturalBreaks", "Manual")
# Porazdelitev, kadar je različnih vrednosti manj kot razredov: (st_razredov, st_vrednosti) -> razredi
MALO_VREDNOSTI = {
    (5, 1): [5], (5, 2): [1, 5], (5, 3): [1, 3, 5], (5, 4): [1, 3, 4, 5],
    (10, 1): [10], (10, 2): [5, 10], (10, 3): [2, 6, 10], (10, 4): [2, 4, 7, 10],
}
NODATA = 0
log = logging.getLogger("skodni_potencial")


# ===========================================================================
# Branje input.xlsx
# ===========================================================================
def _prazno(v):
    return v is None or (isinstance(v, str) and v.strip() == "")


def preberi_list(wb, ime, obvezni):
    if ime not in wb.sheetnames:
        raise ValueError(f"V input.xlsx manjka list '{ime}'.")
    vrstice = list(wb[ime].iter_rows(values_only=True))
    glava = [str(h).strip() if h is not None else "" for h in vrstice[0]]
    for o in obvezni:
        if o not in glava:
            raise ValueError(f"List '{ime}': manjka stolpec '{o}'.")
    out = []
    for r in vrstice[1:]:
        if all(_prazno(v) for v in r):
            continue
        d = {glava[i]: (r[i].strip() if isinstance(r[i], str) else r[i]) for i in range(min(len(glava), len(r)))}
        out.append(d)
    return out


def razberi_meje(v):
    if _prazno(v):
        return None
    if isinstance(v, (int, float)):
        return [float(v)]
    return [float(x.replace(",", ".")) for x in str(v).replace(" ", "").split(";") if x != ""]


def preveri_metodo(metoda, kje):
    if metoda not in METODE:
        raise ValueError(f"{kje}: neznana metoda '{metoda}'. Dovoljene: {', '.join(METODE)}")


def preberi_input(pot):
    baza = os.path.dirname(os.path.abspath(pot))

    def abs_pot(x):
        x = str(x)
        return x if os.path.isabs(x) else os.path.normpath(os.path.join(baza, x))

    wb = openpyxl.load_workbook(pot, read_only=True, data_only=True)
    nast = {r["parameter"]: r.get("vrednost") for r in preberi_list(wb, "nastavitve", ["parameter", "vrednost"])}
    cfg = {
        "pot_t50m": abs_pot(nast["pot_t50m"]),
        "pot_np": abs_pot(nast["pot_nevarnostni_potencial"]),
        "izhod": abs_pot(nast["izhodna_mapa"]),
        "celica": float(nast.get("velikost_celice") or 50),
        "podcelice": int(nast.get("podcelice_poligoni") or 10),
        "ime_izhoda": str(nast.get("ime_izhoda") or "skodni_potencial_SLO"),
        "shrani_vmesne": str(nast.get("shrani_vmesne") or "DA").upper().startswith("D"),
        "ponovno_n1": str(nast.get("ponovno_uporabi_nivo1") or "NE").upper().startswith("D"),
    }
    n3 = preberi_list(wb, "nivo3_razvrstitev", ["metoda", "st_razredov"])
    if len(n3) != 1:
        raise ValueError("List 'nivo3_razvrstitev' mora imeti natanko eno vrstico.")
    preveri_metodo(n3[0]["metoda"], "nivo3_razvrstitev/metoda")
    cfg["metoda3"] = n3[0]["metoda"]
    cfg["n3"] = int(n3[0]["st_razredov"])
    cfg["meje3"] = razberi_meje(n3[0].get("rocne_meje"))
    if cfg["metoda3"] == "Manual" and not cfg["meje3"]:
        raise ValueError("nivo3_razvrstitev: metoda Manual zahteva rocne_meje.")
    cfg["malo"] = MALO_VREDNOSTI

    sloji = OrderedDict()
    for r in preberi_list(wb, "nivo1_sloji", ["koda", "skupina", "metoda", "st_razredov"]):
        k = str(r["koda"])
        preveri_metodo(r["metoda"], f"nivo1_sloji/{k}")
        sloji[k] = {
            "koda": k, "skupina": str(r["skupina"]), "ime": r.get("ime") or k,
            "polje_vrednosti": None if _prazno(r.get("polje_vrednosti")) else str(r["polje_vrednosti"]),
            "polje_unikatno": None if _prazno(r.get("polje_unikatno")) else str(r["polje_unikatno"]),
            "metoda": r["metoda"], "n": int(r["st_razredov"]), "meje": razberi_meje(r.get("rocne_meje")),
            "viri": [], "pravila": [], "utez": None,
        }
        if r["metoda"] == "Manual" and not sloji[k]["meje"]:
            raise ValueError(f"nivo1_sloji/{k}: metoda Manual zahteva rocne_meje.")

    for r in preberi_list(wb, "nivo1_viri", ["koda", "pot_sloja"]):
        k = str(r["koda"])
        if k not in sloji:
            raise ValueError(f"nivo1_viri: koda '{k}' ni v listu nivo1_sloji.")
        pot_s = abs_pot(r["pot_sloja"])
        if not os.path.exists(pot_s):
            raise FileNotFoundError(f"nivo1_viri/{k}: {pot_s}")
        oc = r.get("ocena_konstantna")
        sloji[k]["viri"].append({"pot": pot_s, "ocena": None if _prazno(oc) else float(oc)})

    for r in preberi_list(wb, "nivo1_ocene", ["koda", "polje", "vrednost_od", "ocena"]):
        k = str(r["koda"])
        if k not in sloji:
            raise ValueError(f"nivo1_ocene: koda '{k}' ni v listu nivo1_sloji.")
        sloji[k]["pravila"].append({"polje": str(r["polje"]), "od": r["vrednost_od"],
                                    "do": None if _prazno(r.get("vrednost_do")) else r["vrednost_do"],
                                    "ocena": float(r["ocena"])})

    for r in preberi_list(wb, "nivo2_utezi", ["koda", "utez"]):
        k = str(r["koda"])
        if k not in sloji:
            raise ValueError(f"nivo2_utezi: koda '{k}' ni v listu nivo1_sloji.")
        sloji[k]["utez"] = float(r["utez"])

    skupine = OrderedDict()
    for r in preberi_list(wb, "nivo2_razvrstitev", ["skupina", "metoda", "st_razredov"]):
        s = str(r["skupina"])
        m = str(r["metoda"])
        if m != "PRENOS":
            preveri_metodo(m, f"nivo2_razvrstitev/{s}")
        skupine[s] = {"skupina": s, "metoda": m, "n": int(r["st_razredov"]),
                      "meje": razberi_meje(r.get("rocne_meje")), "utez": None}
        if m == "Manual" and not skupine[s]["meje"]:
            raise ValueError(f"nivo2_razvrstitev/{s}: metoda Manual zahteva rocne_meje.")
    for r in preberi_list(wb, "nivo3_utezi", ["skupina", "utez"]):
        s = str(r["skupina"])
        if s not in skupine:
            raise ValueError(f"nivo3_utezi: skupina '{s}' ni v listu nivo2_razvrstitev.")
        skupine[s]["utez"] = float(r["utez"])
    for s, sk in skupine.items():
        if sk["utez"] is None:
            raise ValueError(f"Skupina '{s}' nima uteži v listu nivo3_utezi.")

    for k, s in sloji.items():
        if not s["viri"]:
            raise ValueError(f"Element '{k}' nima nobenega vira v listu nivo1_viri.")
        if s["skupina"] not in skupine:
            raise ValueError(f"Element '{k}': skupina '{s['skupina']}' ni v listu nivo2_razvrstitev.")
        if skupine[s["skupina"]]["metoda"] != "PRENOS" and s["utez"] is None:
            raise ValueError(f"Element '{k}' nima uteži v listu nivo2_utezi.")
        for v in s["viri"]:
            if v["ocena"] is None and not s["pravila"]:
                raise ValueError(f"Element '{k}': vir {v['pot']} nima konstantne ocene in ni pravil v nivo1_ocene.")
    return cfg, sloji, skupine


# ===========================================================================
# Mreža in maska
# ===========================================================================
class Mreza:
    def __init__(self, xmin, ymax, ncols, nrows, celica, srs_wkt):
        self.xmin, self.ymax, self.ncols, self.nrows, self.c = xmin, ymax, ncols, nrows, celica
        self.srs_wkt = srs_wkt
        self.gt = (xmin, celica, 0.0, ymax, 0.0, -celica)

    @property
    def oblika(self):
        return (self.nrows, self.ncols)

    def indeks(self, x, y):
        """Ploski indeks celic za koordinate; -1 za točke izven mreže."""
        col = np.floor((x - self.xmin) / self.c).astype(np.int64)
        row = np.floor((self.ymax - y) / self.c).astype(np.int64)
        ok = (col >= 0) & (col < self.ncols) & (row >= 0) & (row < self.nrows)
        return np.where(ok, row * self.ncols + col, -1)


def mreza_iz_t50m(pot, celica):
    ds = ogr.Open(pot)
    lyr = ds.GetLayer()
    xmin, xmax, ymin, ymax = lyr.GetExtent()
    srs = lyr.GetSpatialRef()
    ncols = int(round((xmax - xmin) / celica))
    nrows = int(round((ymax - ymin) / celica))
    if abs(xmin + ncols * celica - xmax) > 1e-6 or abs(ymax - nrows * celica - ymin) > 1e-6:
        raise ValueError(f"Obseg T50m ({xmin}, {ymin}, {xmax}, {ymax}) ni večkratnik velikosti celice {celica}.")
    return Mreza(xmin, ymax, ncols, nrows, celica, srs.ExportToWkt())


def zapisi_raster(pot, arr, mreza, dtype, nodata=NODATA):
    os.makedirs(os.path.dirname(pot), exist_ok=True)
    drv = gdal.GetDriverByName("GTiff")
    ds = drv.Create(pot, mreza.ncols, mreza.nrows, 1, dtype,
                    options=["COMPRESS=LZW", "TILED=YES", "BIGTIFF=IF_SAFER"])
    ds.SetGeoTransform(mreza.gt)
    ds.SetProjection(mreza.srs_wkt)
    b = ds.GetRasterBand(1)
    b.SetNoDataValue(nodata)
    b.WriteArray(arr)
    b.FlushCache()
    ds = None


def preberi_na_mrezo(pot, mreza):
    """Prebere raster in ga po potrebi (nearest) prilagodi mreži."""
    src = gdal.Open(pot)
    if (src.RasterXSize, src.RasterYSize) == (mreza.ncols, mreza.nrows) and \
            np.allclose(src.GetGeoTransform(), mreza.gt):
        return src.GetRasterBand(1).ReadAsArray()
    log.info("  raster %s se prilagaja mreži T50m (nearest)", os.path.basename(pot))
    ds = gdal.Warp("", src, format="MEM", outputBounds=(mreza.xmin, mreza.ymax - mreza.nrows * mreza.c,
                                                        mreza.xmin + mreza.ncols * mreza.c, mreza.ymax),
                   xRes=mreza.c, yRes=mreza.c, dstSRS=mreza.srs_wkt, resampleAlg="near")
    return ds.GetRasterBand(1).ReadAsArray()


def maska_t50m(cfg, mreza, vmesna):
    pot_c = os.path.join(vmesna, "t50m_maska.tif")
    if os.path.exists(pot_c):
        ds = gdal.Open(pot_c)
        if (ds.RasterXSize, ds.RasterYSize) == (mreza.ncols, mreza.nrows) and np.allclose(ds.GetGeoTransform(), mreza.gt):
            log.info("Maska T50m: uporabljam predpomnjeno %s", pot_c)
            return ds.GetRasterBand(1).ReadAsArray() == 1
    log.info("Maska T50m: rasterizacija %s (enkratno, lahko traja nekaj minut) ...", cfg["pot_t50m"])
    t = time.time()
    ds = gdal.GetDriverByName("MEM").Create("", mreza.ncols, mreza.nrows, 1, gdal.GDT_Byte)
    ds.SetGeoTransform(mreza.gt)
    ds.SetProjection(mreza.srs_wkt)
    src = ogr.Open(cfg["pot_t50m"])
    gdal.RasterizeLayer(ds, [1], src.GetLayer(), burn_values=[1])
    arr = ds.GetRasterBand(1).ReadAsArray()
    zapisi_raster(pot_c, arr, mreza, gdal.GDT_Byte)
    log.info("  končano v %.0f s, celic: %d", time.time() - t, int((arr == 1).sum()))
    return arr == 1


# ===========================================================================
# Ocene (preslikava atributov)
# ===========================================================================
def _v_stevilo(v):
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return float(v)
    try:
        return float(str(v).strip().replace(",", "."))
    except (TypeError, ValueError):
        return None


def _norm(v):
    return str(v).strip().casefold()


def ocena_vrednosti(vrednost, pravila):
    """Vrne oceno prvega pravila, ki ustreza vrednosti atributa, ali None."""
    if vrednost is None:
        return None
    vn = _v_stevilo(vrednost)
    for p in pravila:
        od_n = _v_stevilo(p["od"])
        if p["do"] is None:
            if vn is not None and od_n is not None:
                if abs(vn - od_n) < 1e-9:
                    return p["ocena"]
            elif _norm(vrednost) == _norm(p["od"]):
                return p["ocena"]
        else:
            do_n = _v_stevilo(p["do"])
            if vn is not None and od_n is not None and do_n is not None and od_n - 1e-9 <= vn <= do_n + 1e-9:
                return p["ocena"]
    return None


# ===========================================================================
# Zapis elementov na mrežo
# ===========================================================================
def _tip_geometrije(lyr):
    gt = ogr.GT_Flatten(lyr.GetGeomType())
    if gt in (ogr.wkbPoint, ogr.wkbMultiPoint):
        return "tocke"
    if gt in (ogr.wkbLineString, ogr.wkbMultiLineString):
        return "linije"
    if gt in (ogr.wkbPolygon, ogr.wkbMultiPolygon):
        return "poligoni"
    raise ValueError(f"Nepodprt tip geometrije: {ogr.GeometryTypeToName(lyr.GetGeomType())}")


def _deli_geometrije(g):
    if g.GetGeometryCount() > 0 and ogr.GT_Flatten(g.GetGeometryType()) in (
            ogr.wkbMultiPoint, ogr.wkbMultiLineString, ogr.wkbGeometryCollection):
        return [g.GetGeometryRef(i) for i in range(g.GetGeometryCount())]
    return [g]


def dolzine_v_celicah(x0, y0, x1, y1, w, mreza, n_celic):
    """Natančna dolžina segmentov po celicah (rezanje na mrežnih črtah), utežena z w. Vrne ploski array."""
    out = np.zeros(n_celic, dtype=np.float64)
    c = mreza.c
    korak = 2_000_000
    for s in range(0, len(x0), korak):
        a0, b0, a1, b1, ww = x0[s:s + korak], y0[s:s + korak], x1[s:s + korak], y1[s:s + korak], w[s:s + korak]
        L = np.hypot(a1 - a0, b1 - b0)
        gx0, gx1 = (a0 - mreza.xmin) / c, (a1 - mreza.xmin) / c
        gy0, gy1 = (mreza.ymax - b0) / c, (mreza.ymax - b1) / c
        dx, dy = gx1 - gx0, gy1 - gy0
        n = len(a0)
        seg = np.arange(n)
        ts, ids = [np.zeros(n), np.ones(n)], [seg, seg]
        for g0, g1, d in ((gx0, gx1, dx), (gy0, gy1, dy)):
            f0, f1 = np.floor(g0), np.floor(g1)
            k = np.abs(f1 - f0).astype(np.int64)
            if k.sum() == 0:
                continue
            idx = np.repeat(seg, k)
            zac = np.cumsum(k) - k
            j = np.arange(idx.size) - np.repeat(zac, k)
            pos = d[idx] > 0
            linija = np.where(pos, f0[idx] + 1 + j, f0[idx] - j)
            ts.append(np.clip((linija - g0[idx]) / d[idx], 0.0, 1.0))
            ids.append(idx)
        t = np.concatenate(ts)
        sid = np.concatenate(ids)
        o = np.lexsort((t, sid))
        t, sid = t[o], sid[o]
        isti = sid[:-1] == sid[1:]
        ta, tb, si = t[:-1][isti], t[1:][isti], sid[:-1][isti]
        dl = (tb - ta) * L[si]
        ok = dl > 0
        ta, tb, si, dl = ta[ok], tb[ok], si[ok], dl[ok]
        tm = (ta + tb) / 2.0
        col = np.floor(gx0[si] + dx[si] * tm).astype(np.int64)
        row = np.floor(gy0[si] + dy[si] * tm).astype(np.int64)
        v = (col >= 0) & (col < mreza.ncols) & (row >= 0) & (row < mreza.nrows)
        out += np.bincount((row * mreza.ncols + col)[v], weights=(dl * ww[si])[v], minlength=n_celic)
    return out


def povrsine_v_celicah(mem_lyr, mreza, f, maska, n_celic):
    """Površina poligonov [m2] x ocena po celicah; rasterizacija na podcelice z MERGE_ALG=ADD, po ploščicah."""
    out = np.zeros(mreza.oblika, dtype=np.float64)
    ploscica = max(1, 4000 // f)  # celic na stranico ploščice
    pc = mreza.c / f
    xmin_l, xmax_l, ymin_l, ymax_l = mem_lyr.GetExtent()
    for r0 in range(0, mreza.nrows, ploscica):
        r1 = min(r0 + ploscica, mreza.nrows)
        for c0 in range(0, mreza.ncols, ploscica):
            c1 = min(c0 + ploscica, mreza.ncols)
            if not maska[r0:r1, c0:c1].any():
                continue
            tx0, tx1 = mreza.xmin + c0 * mreza.c, mreza.xmin + c1 * mreza.c
            ty1, ty0 = mreza.ymax - r0 * mreza.c, mreza.ymax - r1 * mreza.c
            if tx1 < xmin_l or tx0 > xmax_l or ty1 < ymin_l or ty0 > ymax_l:
                continue
            mem_lyr.SetSpatialFilterRect(tx0, ty0, tx1, ty1)
            if mem_lyr.GetFeatureCount() == 0:
                continue
            nr, nc = (r1 - r0) * f, (c1 - c0) * f
            ds = gdal.GetDriverByName("MEM").Create("", nc, nr, 1, gdal.GDT_Float32)
            ds.SetGeoTransform((tx0, pc, 0.0, ty1, 0.0, -pc))
            ds.SetProjection(mreza.srs_wkt)
            gdal.RasterizeLayer(ds, [1], mem_lyr, options=["ATTRIBUTE=OCENA", "MERGE_ALG=ADD"])
            a = ds.GetRasterBand(1).ReadAsArray().astype(np.float64)
            out[r0:r1, c0:c1] += a.reshape(r1 - r0, f, c1 - c0, f).sum(axis=(1, 3)) * pc * pc
            ds = None
    mem_lyr.SetSpatialFilter(None)
    return out.ravel()


def vrednosti_elementa(sloj, mreza, maska, cfg, opozorila):
    """Izračuna vrednosti nivoja 1 (ploski array float64) za en element ranljivosti."""
    n_celic = mreza.nrows * mreza.ncols
    vrednost = np.zeros(n_celic, dtype=np.float64)
    pu, pv = sloj["polje_unikatno"], sloj["polje_vrednosti"]
    videni = set()
    stat = Counter()
    neprepoznane = Counter()
    for vir in sloj["viri"]:
        ds = ogr.Open(vir["pot"])
        lyr = ds.GetLayer()
        tip = _tip_geometrije(lyr)
        dfn = lyr.GetLayerDefn()
        polja = {dfn.GetFieldDefn(i).GetName().upper(): dfn.GetFieldDefn(i).GetName() for i in range(dfn.GetFieldCount())}

        def polje(ime):
            if ime is None:
                return None
            if ime.upper() not in polja:
                raise ValueError(f"{sloj['koda']}: v {vir['pot']} ni polja '{ime}'.")
            return polja[ime.upper()]

        f_unik, f_vred = polje(pu) if pu and pu.upper() in polja else None, polje(pv)
        if pu and f_unik is None:
            opozorila.append((sloj["koda"], f"vir {os.path.basename(vir['pot'])} nima polja {pu}; brez odstranjevanja dvojnikov"))
        f_prav = sorted({polje(p["polje"]) for p in sloj["pravila"]}) if vir["ocena"] is None else []
        potrebna = {x for x in [f_unik, f_vred] + f_prav if x}
        lyr.SetIgnoredFields([ime for ime in polja.values() if ime not in potrebna])
        predpomnilnik = {}

        def ocena_feature(ft):
            if vir["ocena"] is not None:
                return vir["ocena"]
            for fp in f_prav:
                val = ft.GetField(fp)
                kljuc = (fp, val)
                if kljuc not in predpomnilnik:
                    predpomnilnik[kljuc] = ocena_vrednosti(val, [p for p in sloj["pravila"] if polje(p["polje"]) == fp])
                if predpomnilnik[kljuc] is not None:
                    return predpomnilnik[kljuc]
            neprepoznane[tuple(ft.GetField(fp) for fp in f_prav)] += 1
            return None

        if tip == "poligoni":
            mem_ds = (ogr.GetDriverByName("MEM") or ogr.GetDriverByName("Memory")).CreateDataSource("")
            mem = mem_ds.CreateLayer("p", srs=lyr.GetSpatialRef(), geom_type=ogr.wkbMultiPolygon)
            mem.CreateField(ogr.FieldDefn("OCENA", ogr.OFTReal))
            mdfn = mem.GetLayerDefn()
        xs, ys, ws, segs = [], [], [], []
        for ft in lyr:
            stat["objektov"] += 1
            if f_unik:
                u = ft.GetField(f_unik)
                if u is not None:
                    if u in videni:
                        stat["dvojnikov"] += 1
                        continue
                    videni.add(u)
            oc = ocena_feature(ft)
            if oc is None:
                stat["neprepoznanih"] += 1
                continue
            if f_vred:
                v = ft.GetField(f_vred)
                oc = oc * (float(v) if v is not None else 0.0)
            if oc == 0:
                continue
            g = ft.GetGeometryRef()
            if g is None or g.IsEmpty():
                continue
            stat["uporabljenih"] += 1
            if tip == "tocke":
                for d in _deli_geometrije(g):
                    xs.append(d.GetX())
                    ys.append(d.GetY())
                    ws.append(oc)
            elif tip == "linije":
                for d in _deli_geometrije(g):
                    pts = np.asarray(d.GetPoints(), dtype=np.float64)
                    if pts.ndim == 2 and len(pts) >= 2:
                        segs.append(pts[:, :2])
                        ws.append((oc, len(pts) - 1))
            else:
                nf = ogr.Feature(mdfn)
                nf.SetGeometry(g)
                nf.SetField("OCENA", float(oc))
                mem.CreateFeature(nf)
        if tip == "tocke" and xs:
            idx = mreza.indeks(np.asarray(xs), np.asarray(ys))
            ok = idx >= 0
            vrednost += np.bincount(idx[ok], weights=np.asarray(ws)[ok], minlength=n_celic)
        elif tip == "linije" and segs:
            p0 = np.concatenate([s[:-1] for s in segs])
            p1 = np.concatenate([s[1:] for s in segs])
            w = np.repeat([x[0] for x in ws], [x[1] for x in ws]).astype(np.float64)
            vrednost += dolzine_v_celicah(p0[:, 0], p0[:, 1], p1[:, 0], p1[:, 1], w, mreza, n_celic)
        elif tip == "poligoni" and mem.GetFeatureCount() > 0:
            vrednost += povrsine_v_celicah(mem, mreza, cfg["podcelice"], maska, n_celic)
        ds = None
    if stat["neprepoznanih"]:
        primeri = ", ".join(f"{k}: {v}" for k, v in neprepoznane.most_common(8))
        opozorila.append((sloj["koda"], f"{stat['neprepoznanih']} objektov brez ocene (ni pravila) - npr. {primeri}"))
    if stat["dvojnikov"]:
        opozorila.append((sloj["koda"], f"odstranjenih {stat['dvojnikov']} dvojnikov po polju {pu}"))
    return vrednost.reshape(mreza.oblika), stat


# ===========================================================================
# Razvrščanje
# ===========================================================================
def jenks_meje(vrednosti, n, max_unikatnih=3000):
    """Natural Breaks (Jenks) na histogramu unikatnih vrednosti. Vrne zgornje meje razredov 1..n-1."""
    u, w = np.unique(vrednosti, return_counts=True)
    w = w.astype(np.float64)
    if len(u) > max_unikatnih:  # združevanje v bloke z enakim številom celic (aproksimacija)
        cum = np.cumsum(w)
        skupina = np.minimum((cum - w / 2) / cum[-1] * max_unikatnih, max_unikatnih - 1).astype(np.int64)
        vs = np.bincount(skupina, weights=u * w)
        ws = np.bincount(skupina, weights=w)
        zg = np.zeros(skupina.max() + 1)
        np.maximum.at(zg, skupina, u)
        ok = ws > 0
        u_rep, w, zgornja = vs[ok] / ws[ok], ws[ok], zg[ok]
    else:
        u_rep, zgornja = u.astype(np.float64), u.astype(np.float64)
    m = len(u_rep)
    if m <= n:
        return list(zgornja[:-1])
    S1 = np.concatenate([[0.0], np.cumsum(w * u_rep)])
    S2 = np.concatenate([[0.0], np.cumsum(w * u_rep * u_rep)])
    W = np.concatenate([[0.0], np.cumsum(w)])

    def ssd(i, j):  # odsek i..j (vključno), i je array
        s1 = S1[j + 1] - S1[i]
        s2 = S2[j + 1] - S2[i]
        ww = W[j + 1] - W[i]
        return s2 - s1 * s1 / ww

    cena = np.full((n, m), np.inf)
    zac = np.zeros((n, m), dtype=np.int64)
    cena[0] = ssd(np.zeros(m, dtype=np.int64), np.arange(m))
    for k in range(1, n):
        for j in range(k, m):
            i = np.arange(k, j + 1)
            c = cena[k - 1, i - 1] + ssd(i, j)
            a = int(np.argmin(c))
            cena[k, j], zac[k, j] = c[a], i[a]
    meje = []
    j = m - 1
    for k in range(n - 1, 0, -1):
        i = zac[k, j]
        meje.append(zgornja[i - 1])
        j = i - 1
    return sorted(meje)


def kvantilne_meje(v, n):
    """Kvantilne meje z različnimi vrednostmi (kot ArcGIS pri celoštevilskih podatkih z veliko enakimi vrednostmi).

    Meja i je tista unikatna vrednost, pri kateri je kumulativno število celic najbližje cilju
    (preostale celice / preostali razredi). Meje so strogo naraščajoče, zato se razredi ne združujejo.
    """
    u, w = np.unique(v, return_counts=True)
    if len(u) <= n:
        return list(u[:-1].astype(np.float64))
    cum = np.cumsum(w).astype(np.float64)
    N = cum[-1]
    meje, j_prej, cum_prej = [], -1, 0.0
    for i in range(1, n):
        cilj = cum_prej + (N - cum_prej) / (n - i + 1)
        lo = j_prej + 1
        hi = len(u) - 1 - (n - i)  # pusti dovolj vrednosti za preostale razrede
        kandidati = np.arange(lo, hi + 1)
        j = int(kandidati[np.argmin(np.abs(cum[kandidati] - cilj))])
        meje.append(float(u[j]))
        j_prej, cum_prej = j, cum[j]
    return meje


def izracunaj_meje(v, metoda, n, rocne=None):
    if metoda == "Manual":
        return sorted(float(x) for x in rocne)
    if metoda == "Quantile":
        return kvantilne_meje(v, n)
    if metoda == "EqualInterval":
        lo, hi = float(v.min()), float(v.max())
        return [lo + (hi - lo) * i / n for i in range(1, n)]
    if metoda == "NaturalBreaks":
        return jenks_meje(v, n)
    raise ValueError(metoda)


def razredi_malo_vrednosti(n, m, tabela):
    """Razredi za m različnih vrednosti (m < n): iz tabele malo_vrednosti, sicer enakomerno do razreda n."""
    if (n, m) in tabela:
        return tabela[(n, m)]
    return sorted({max(1, int(round(n * k / m))) for k in range(1, m + 1)})


def razvrsti(arr, veljavno, metoda, n, rocne=None, malo=None, omeji=False):
    """Razvrsti celice arr[veljavno & arr>0] v razrede 1..n (vrednost <= meja_i -> razred i). Ostalo 0.

    Kadar je različnih vrednosti manj kot razredov (pri vseh metodah), i-ta najmanjša vrednost dobi
    i-ti razred iz tabele MALO_VREDNOSTI. Pri omeji=True (nivo 1, vrednosti na lestvici ocen) se
    vrednosti pred tem omejijo na n (npr. IED: 5 in 10 -> 5 -> razred 5).
    Vrne (razredi, meje, celice_po_razredih, opomba).
    """
    izbor = veljavno & (arr > 0)
    out = np.zeros(arr.shape, dtype=np.uint8)
    v = arr[izbor]
    if v.size == 0:
        return out, [], {}, "ni celic z vrednostjo"
    opomba = ""
    u_izv = np.unique(v)
    if len(u_izv) < n:
        vo = np.minimum(v, n) if omeji else v
        u = np.unique(vo)
        razredi_u = np.asarray(razredi_malo_vrednosti(n, len(u), malo or {}), dtype=np.uint8)
        out[izbor] = razredi_u[np.searchsorted(u, vo)]
        meje = list(u[:-1])
        opomba = (f"{len(u_izv)} različnih vrednosti < {n} razredov: vrednosti {fmt_meje(u_izv)} "
                  + (f"(omejeno na {n}: {fmt_meje(u)}) " if omeji and len(u) < len(u_izv) else "")
                  + f"-> razredi {';'.join(str(x) for x in razredi_u)}")
    else:
        meje = np.unique(np.asarray(izracunaj_meje(v, metoda, n, rocne), dtype=np.float64))
        out[izbor] = np.minimum(np.searchsorted(meje, v, side="left") + 1, n).astype(np.uint8)
        meje = list(meje)
    stevila = dict(zip(*np.unique(out[izbor], return_counts=True)))
    return out, meje, {int(k): int(c) for k, c in stevila.items()}, opomba


# ===========================================================================
# Poročilo
# ===========================================================================
def fmt_meje(meje):
    return "; ".join(f"{m:.6g}" for m in meje)


def fmt_razredi(st):
    return "; ".join(f"{k}:{v}" for k, v in sorted(st.items()))


def zapisi_porocilo(pot, vrstice_n1, vrstice_n2, vrstice_n3, opozorila, cfg, pot_input):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "povzetek"
    ws.append(["parameter", "vrednost"])
    ws.append(["datum", dt.datetime.now().strftime("%Y-%m-%d %H:%M")])
    ws.append(["input", os.path.abspath(pot_input)])
    for k in ("pot_t50m", "pot_np", "celica", "podcelice", "metoda3", "n3"):
        ws.append([k, str(cfg[k])])
    glava = ["nivo", "koda", "skupina", "objektov", "uporabljenih", "dvojnikov", "brez_ocene",
             "celic_z_vrednostjo_v_maski", "min", "max", "metoda", "st_razredov", "meje", "celic_po_razredih"]
    for ime, vrst in (("nivo1", vrstice_n1), ("nivo2", vrstice_n2), ("nivo3", vrstice_n3)):
        w = wb.create_sheet(ime)
        w.append(glava)
        for r in vrst:
            w.append([r.get(k) for k in glava])
    w = wb.create_sheet("opozorila")
    w.append(["koda", "opozorilo"])
    for o in opozorila:
        w.append(list(o))
    wb.save(pot)


def vrstica(nivo, koda, skupina, arr, veljavno, metoda, n, meje, st, stat=None):
    v = arr[veljavno & (arr > 0)]
    stat = stat or {}
    return {"nivo": nivo, "koda": koda, "skupina": skupina, "objektov": stat.get("objektov"),
            "uporabljenih": stat.get("uporabljenih"), "dvojnikov": stat.get("dvojnikov"),
            "brez_ocene": stat.get("neprepoznanih"), "celic_z_vrednostjo_v_maski": int(v.size),
            "min": float(v.min()) if v.size else None, "max": float(v.max()) if v.size else None,
            "metoda": metoda, "st_razredov": n, "meje": fmt_meje(meje), "celic_po_razredih": fmt_razredi(st)}


# ===========================================================================
# Glavni postopek
# ===========================================================================
def nastavi_log(mapa, datum):
    os.makedirs(mapa, exist_ok=True)
    log.setLevel(logging.INFO)
    fmt = logging.Formatter("%(asctime)s  %(message)s", "%H:%M:%S")
    for h in (logging.StreamHandler(sys.stdout), logging.FileHandler(os.path.join(mapa, f"log_{datum}.txt"), encoding="utf-8")):
        h.setFormatter(fmt)
        log.addHandler(h)


def je_zaklenjena(pot):
    """True, če je datoteka odprta v drugem programu (preimenovanje ni mogoče)."""
    if not os.path.exists(pot):
        return False
    tmp = pot + ".preveri_zaklep"
    try:
        os.rename(pot, tmp)
        os.rename(tmp, pot)
        return False
    except OSError:
        return True


def preveri_izhodne_datoteke(poti):
    """Pred izračunom preveri, da nobena izhodna datoteka ni odprta; sicer prekine izvajanje."""
    zaklenjene = []
    for p in poti:
        for d in (p, p + ".ovr", p + ".aux.xml"):
            if je_zaklenjena(d):
                zaklenjene.append(d)
    if zaklenjene:
        raise RuntimeError("Izračun prekinjen - naslednje datoteke so odprte v drugem programu "
                           "(zaprite jih v ArcGIS/QGIS/Excelu in ponovno zaženite):\n  " + "\n  ".join(zaklenjene))
    for p in poti:  # odstrani stare piramide/statistiko, da se ne ujemajo z novim rastrom
        for d in (p + ".ovr", p + ".aux.xml"):
            if os.path.exists(d):
                os.remove(d)


def main(pot_input):
    t0 = time.time()
    cfg, sloji, skupine = preberi_input(pot_input)
    datum = dt.datetime.now().strftime("%Y%m%d")
    vmesna = os.path.join(cfg["izhod"], "vmesni")
    os.makedirs(vmesna, exist_ok=True)
    nastavi_log(cfg["izhod"], datum)
    log.info("Vhod: %s  (elementov: %d, skupin: %d)", os.path.abspath(pot_input), len(sloji), len(skupine))
    pot_out = os.path.join(cfg["izhod"], f"{cfg['ime_izhoda']}_{datum}.tif")
    pot_por = os.path.join(cfg["izhod"], f"porocilo_{datum}.xlsx")
    vmesne = [os.path.join(vmesna, f) for f in os.listdir(vmesna) if f.lower().endswith(".tif")]
    preveri_izhodne_datoteke([pot_out, pot_por] + vmesne)
    log.info("Izhodne datoteke niso odprte - začenjam izračun.")

    mreza = mreza_iz_t50m(cfg["pot_t50m"], cfg["celica"])
    log.info("Mreža: %d x %d celic, %.0f m, izhodišče (%.0f, %.0f)", mreza.ncols, mreza.nrows, mreza.c, mreza.xmin, mreza.ymax)
    np_arr = preberi_na_mrezo(cfg["pot_np"], mreza)
    kopno = maska_t50m(cfg, mreza, vmesna)
    maska = kopno & (np_arr == 1)
    log.info("Maska analize (T50m ∩ nevarnostni potencial = 1): %d celic", int(maska.sum()))

    opozorila, vrst_n1, vrst_n2, vrst_n3 = [], [], [], []

    # ---------------- Nivo 1
    razredi_n1 = {}
    for k, s in sloji.items():
        t = time.time()
        pot_v = os.path.join(vmesna, f"N1_vrednost_{k}.tif")
        if cfg["ponovno_n1"] and os.path.exists(pot_v):
            arr = preberi_na_mrezo(pot_v, mreza).astype(np.float64)
            stat = {}
            log.info("N1 %-8s vrednosti prebrane iz %s", k, pot_v)
        else:
            arr, stat = vrednosti_elementa(s, mreza, maska, cfg, opozorila)
            if cfg["shrani_vmesne"]:
                zapisi_raster(pot_v, np.where(maska, arr, 0).astype(np.float32), mreza, gdal.GDT_Float32)
        r, meje, st, opomba = razvrsti(arr, maska, s["metoda"], s["n"], s["meje"], cfg["malo"], omeji=True)
        if opomba:
            opozorila.append((k, f"nivo 1: {opomba}"))
        elif len(st) < s["n"]:
            opozorila.append((k, f"nivo 1: uporabljenih {len(st)} od {s['n']} razredov ({fmt_razredi(st)})"))
        razredi_n1[k] = r
        vrst_n1.append(vrstica(1, k, s["skupina"], arr, maska, s["metoda"], s["n"], meje, st, stat))
        if cfg["shrani_vmesne"]:
            zapisi_raster(os.path.join(vmesna, f"N1_razred_{k}.tif"), r, mreza, gdal.GDT_Byte)
        v = arr[maska & (arr > 0)]
        log.info("N1 %-8s objektov %s, celic>0 v maski %d, razpon %s-%s, %s -> %s  (%.0f s)", k,
                 stat.get("uporabljenih", "-"), v.size, f"{v.min():.0f}" if v.size else "-",
                 f"{v.max():.0f}" if v.size else "-", s["metoda"], fmt_razredi(st), time.time() - t)
        del arr

    # ---------------- Nivo 2
    razredi_n2 = {}
    for g, sk in skupine.items():
        clani = [k for k, s in sloji.items() if s["skupina"] == g]
        if sk["metoda"] == "PRENOS":
            if len(clani) != 1:
                raise ValueError(f"Skupina {g}: PRENOS zahteva natanko en element, najdeni: {clani}")
            razredi_n2[g] = razredi_n1[clani[0]]
            st = dict(zip(*np.unique(razredi_n2[g][maska & (razredi_n2[g] > 0)], return_counts=True)))
            vrst_n2.append(vrstica(2, g, g, razredi_n2[g].astype(np.float64), maska, "PRENOS", sk["n"], [],
                                   {int(a): int(b) for a, b in st.items()}))
            log.info("N2 %-3s prenos iz %s", g, clani[0])
            continue
        vsota = np.zeros(mreza.oblika, dtype=np.float64)
        for k in clani:
            vsota += sloji[k]["utez"] * razredi_n1[k]
        r, meje, st, opomba = razvrsti(vsota, maska, sk["metoda"], sk["n"], sk["meje"], cfg["malo"])
        if opomba:
            opozorila.append((g, f"nivo 2: {opomba}"))
        elif len(st) < sk["n"]:
            opozorila.append((g, f"nivo 2: uporabljenih {len(st)} od {sk['n']} razredov ({fmt_razredi(st)})"))
        razredi_n2[g] = r
        vrst_n2.append(vrstica(2, g, g, vsota, maska, sk["metoda"], sk["n"], meje, st))
        if cfg["shrani_vmesne"]:
            zapisi_raster(os.path.join(vmesna, f"N2_vsota_{g}.tif"), np.where(maska, vsota, 0).astype(np.float32),
                          mreza, gdal.GDT_Float32)
            zapisi_raster(os.path.join(vmesna, f"N2_razred_{g}.tif"), r, mreza, gdal.GDT_Byte)
        v = vsota[maska & (vsota > 0)]
        log.info("N2 %-3s razpon vsote %.0f-%.0f, %s -> %s", g, v.min() if v.size else 0, v.max() if v.size else 0,
                 sk["metoda"], fmt_razredi(st))

    # ---------------- Nivo 3
    vsota3 = np.zeros(mreza.oblika, dtype=np.float64)
    for g, sk in skupine.items():
        vsota3 += sk["utez"] * razredi_n2[g]
    koncni, meje3, st3, opomba3 = razvrsti(vsota3, maska, cfg["metoda3"], cfg["n3"], cfg["meje3"], cfg["malo"])
    if opomba3:
        opozorila.append(("nivo3", opomba3))
    vrst_n3.append(vrstica(3, "SKUPAJ", "", vsota3, maska, cfg["metoda3"], cfg["n3"], meje3, st3))
    v = vsota3[maska & (vsota3 > 0)]
    log.info("N3 razpon vsote %.0f-%.0f, %s meje: %s", v.min(), v.max(), cfg["metoda3"], fmt_meje(meje3))
    log.info("N3 celic po razredih: %s", fmt_razredi(st3))
    if cfg["shrani_vmesne"]:
        zapisi_raster(os.path.join(vmesna, "N3_vsota.tif"), np.where(maska, vsota3, 0).astype(np.float32), mreza,
                      gdal.GDT_Float32)

    zapisi_raster(pot_out, koncni, mreza, gdal.GDT_Byte)
    zapisi_porocilo(pot_por, vrst_n1, vrst_n2, vrst_n3, opozorila, cfg, pot_input)
    for k, o in opozorila:
        log.info("OPOZORILO %s: %s", k, o)
    log.info("Končni raster: %s", pot_out)
    log.info("Poročilo:      %s", pot_por)
    log.info("Skupni čas: %.1f min", (time.time() - t0) / 60)
    return 0


if __name__ == "__main__":
    pot = sys.argv[1] if len(sys.argv) > 1 else os.path.join(os.path.dirname(os.path.abspath(__file__)), "input.xlsx")
    try:
        sys.exit(main(pot))
    except RuntimeError as e:
        log.error("%s", e) if log.handlers else print(e)
        sys.exit(1)
