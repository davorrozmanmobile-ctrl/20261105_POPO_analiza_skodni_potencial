# Analiza škodnega potenciala – Slovenija (50 × 50 m)

Skripta `skodni_potencial.py` izdela raster škodnega potenciala z vrednostmi **1–10** (UInt8, NoData = 0)
po metodologiji (pogl. 3.3, `metodologija_škodni_potencial.pdf`). Vsi parametri se berejo iz `input.xlsx`.

## Zagon

Potreben je Python z GDAL, numpy in openpyxl – uporablja se ArcGIS Pro Python (arcpy ni potreben):

```
"C:\Program Files\ArcGIS\Pro\bin\Python\envs\arcgispro-py3\python.exe" skodni_potencial.py input.xlsx
```

Prvi zagon rasterizira `T50m.shp` (1 GB) in masko shrani v `_rezultati\vmesni\t50m_maska.tif`; naslednji zagoni jo uporabijo znova.



## Rezultati (`_rezultati`)

| datoteka | vsebina |
|---|---|
| `skodni_potencial_SLO_<datum>.tif` | končni raster 1–10 |
| `porocilo_<datum>.xlsx` | razponi vrednosti, meje razredov in število celic po razredih za vsak element/skupino/nivo; opozorila |
| `log_<datum>.txt` | potek izračuna |
| `vmesni\N1_vrednost_<koda>.tif`, `N1_razred_<koda>.tif` | nivo 1 (vrednosti in razredi 1–5) |
| `vmesni\N2_vsota_<skupina>.tif`, `N2_razred_<skupina>.tif` | nivo 2 |
| `vmesni\N3_vsota.tif` | ponderirana vsota nivoja 3 pred razvrstitvijo |

## Postopek

1. **Mreža**: obseg in poravnava iz `T50m.shp`, velikost celice iz nastavitev. Analiza poteka samo v celicah T50m,
   kjer ima raster nevarnostnega potenciala vrednost 1.
2. **Nivo 1** – za vsak element ranljivosti:
   - točke: vsota ocen v celici (pri ZL vsota prebivalcev, `polje_vrednosti = COUNT___`),
   - linije: vsota (dolžina v celici [m] × ocena), natančno rezanje na mrežnih črtah,
   - poligoni: vsota (površina v celici [m²] × ocena), rasterizacija na podcelice (privzeto 5 m), prekrivanja se seštejejo;
   - razvrstitev celic z vrednostjo > 0 v razrede (privzeto Quantile, 5 razredov; ZL 10 razredov).
3. **Nivo 2**: Σ (utež × razred N1) po skupinah GD, IN, KD, OK, SI → 10 razredov (Quantile). ZL se prenese.
4. **Nivo 3**: Σ (utež × razred N2) → 10 razredov (Natural Breaks, natančen Jenks na histogramu vrednosti).

Celice brez elementov ranljivosti in celice izven maske so NoData.

## `input.xlsx`

| list | stolpci | opis |
|---|---|---|
| `nastavitve` | parameter, vrednost, opis | poti do `T50m.shp`, `nevarnostni_potencial_raster.tif`, izhodne mape; velikost celice; shranjevanje vmesnih rastrov |
| `nivo1_viri` | koda, pot_sloja, ocena_konstantna | prostorski sloji elementa (lahko več); `ocena_konstantna` velja za vse objekte sloja |
| `nivo1_sloji` | koda, skupina, ime, polje_vrednosti, polje_unikatno, metoda, st_razredov, rocne_meje | element ranljivosti in njegova razvrstitev na nivoju 1 |
| `nivo1_ocene` | koda, polje, vrednost_od, vrednost_do, ocena, opis | preslikava atributa v oceno 1–5; brez `vrednost_do` = točno ujemanje, sicer interval (npr. SKD `35.111`–`35.230`); velja prvo ustrezno pravilo |
| `nivo2_utezi` | koda, skupina, utez | utež elementa na nivoju 2 |
| `nivo2_razvrstitev` | skupina, metoda, st_razredov, rocne_meje | razvrstitev skupine na nivoju 2 (`PRENOS` = brez seštevanja, za ZL) |
| `nivo3_utezi` | skupina, utez | utež skupine na nivoju 3 |
| `nivo3_razvrstitev` | metoda, st_razredov, rocne_meje | razvrstitev končnega rastra (privzeto NaturalBreaks, 10) |

Metode razvrščanja: `Quantile`, `EqualInterval`, `NaturalBreaks`, `Manual` (seznam je tudi v komentarju glave stolpca metoda v Excelu).

Kadar ima sloj manj različnih vrednosti kot razredov, i-ta najmanjša vrednost (pri vseh metodah) dobi i-ti razred:
5 razredov `5`, `1;5`, `1;3;5`, `1;3;4;5`; 10 razredov `10`, `5;10`, `2;6;10`, `2;4;7;10` (tabela `MALO_VREDNOSTI` v skripti;
za druge kombinacije enakomerno do najvišjega razreda). Na nivoju 1 se vrednosti pred tem omejijo na lestvico ocen
(vrednost > 5 → 5). Primer IED: celice z 1 ali 2 napravama (5 oz. 10) → ena vrednost (5) → razred 5.
Pri `Manual` se v `rocne_meje` vpišejo zgornje meje razredov 1..n-1, ločene s `;` (npr. `1;2;3;4` → vrednost ≤ 1 je razred 1, … > 4 je razred 5).

`polje_unikatno = MATICNA` odstrani subjekte AJPES, ki se ponavljajo v več podslojih (napake pri pripravi podatkov).
Objekti, za katere ni pravila v `nivo1_ocene`, se izpustijo in so navedeni na listu `opozorila` v poročilu.
