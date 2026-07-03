"""
Script de automatización La Madrugada
=====================================
1. Se conecta a Dropbox usando refresh token (renueva access token solo)
2. Descarga las dos planillas Excel
3. Procesa los datos (cultivos, ambientes, órdenes de trabajo)
   - Universo de lotes: SOLO los que están en planilla Agricultura
   - Los tratamientos se filtran para incluir solo esos lotes
   - Log de tratamientos ignorados por no corresponder a lote agrícola
4. Genera el HTML actualizado y lo guarda como index.html

Variables de entorno requeridas:
- DROPBOX_APP_KEY
- DROPBOX_APP_SECRET
- DROPBOX_REFRESH_TOKEN
"""
import os
import sys
import json
import requests
from datetime import datetime, timezone, timedelta
from io import BytesIO

# === Zona horaria Argentina (UTC-3) ===
ARG_TZ = timezone(timedelta(hours=-3))

def ahora_argentina():
    return datetime.now(ARG_TZ)

# === Configuración ===
DROPBOX_FOLDER = "/JG/TRABAJO/Clientes/La Madrugada/Agricultura La Madrugada"
PLANILLA_CULTIVOS = "Planilla de Cultivos La Madrugada.xlsx"
PLANILLA_AGRICULTURA = "Planilla Agricultura La Madrugada 26-27.xlsx"

APP_KEY = os.environ.get("DROPBOX_APP_KEY")
APP_SECRET = os.environ.get("DROPBOX_APP_SECRET")
REFRESH_TOKEN = os.environ.get("DROPBOX_REFRESH_TOKEN")

if not all([APP_KEY, APP_SECRET, REFRESH_TOKEN]):
    print("ERROR: faltan variables de entorno (APP_KEY, APP_SECRET, REFRESH_TOKEN)")
    sys.exit(1)


# =========================================================================
# 1. Autenticación Dropbox
# =========================================================================
def get_access_token():
    """Renueva el access_token usando el refresh_token (válido 4hs)."""
    print("→ Renovando access token de Dropbox...")
    r = requests.post(
        "https://api.dropboxapi.com/oauth2/token",
        data={
            "grant_type": "refresh_token",
            "refresh_token": REFRESH_TOKEN,
            "client_id": APP_KEY,
            "client_secret": APP_SECRET,
        },
        timeout=30,
    )
    r.raise_for_status()
    token = r.json()["access_token"]
    print("  ✓ Token renovado")
    return token


def download_file(access_token, dropbox_path):
    """Descarga un archivo de Dropbox y lo devuelve como BytesIO."""
    print(f"→ Descargando: {dropbox_path}")
    r = requests.post(
        "https://content.dropboxapi.com/2/files/download",
        headers={
            "Authorization": f"Bearer {access_token}",
            "Dropbox-API-Arg": json.dumps({"path": dropbox_path}),
        },
        timeout=120,
    )
    r.raise_for_status()
    print(f"  ✓ {len(r.content):,} bytes descargados")
    return BytesIO(r.content)


# =========================================================================
# 2. Procesamiento de datos
# =========================================================================
def procesar_planillas(buf_cultivos, buf_agricultura):
    import pandas as pd

    # --- Diccionario de cultivos abreviados ---
    df_listados = pd.read_excel(buf_cultivos, sheet_name="Listados", header=None)
    cultivos_dict = {}
    for i in range(1, len(df_listados)):
        abrev = df_listados.iloc[i, 1]
        nombre = df_listados.iloc[i, 2]
        if pd.notna(abrev) and pd.notna(nombre):
            cultivos_dict[str(abrev).strip()] = str(nombre).strip()
    cultivos_dict["."] = "—"

    # --- Órdenes de trabajo (Base de Datos) ---
    buf_cultivos.seek(0)
    df = pd.read_excel(buf_cultivos, sheet_name="Base de Datos")
    df = df.dropna(how='all')
    df = df[df['Establecimiento'].notna() & (df['Establecimiento'] != 'TOTALES')]
    df = df[df['Lote'] != '.']
    df['Fecha'] = pd.to_datetime(df['Fecha'], errors='coerce')
    df['Fecha realizado'] = pd.to_datetime(df['Fecha realizado'], errors='coerce')

    ordenes_todas = []
    for (fecha, est, lote), grupo in df.groupby(['Fecha', 'Establecimiento', 'Lote'], dropna=False):
        primera = grupo.iloc[0]
        cultivo_abrev = str(primera['Actividad']).strip() if pd.notna(primera['Actividad']) else "."
        antecesor_abrev = str(primera['Antecesor']).strip() if pd.notna(primera['Antecesor']) else "."

        tipos = grupo['Tipo'].dropna().unique().tolist()
        if 'Herbicida' in tipos or 'Insecticida' in tipos or 'Fungicida' in tipos:
            tipo_labor = "Pulverización"
        elif 'Fertilizante' in tipos:
            tipo_labor = "Fertilización"
        elif any('semilla' in str(t).lower() for t in tipos):
            tipo_labor = "Siembra"
        else:
            tipo_labor = "Labor"

        items = []
        for _, row in grupo.iterrows():
            if pd.notna(row['Producto/Labor']) and str(row['Producto/Labor']).strip() != '.':
                items.append({
                    "producto": str(row['Producto/Labor']).strip(),
                    "dosis": float(row['Dosis/ha']) if pd.notna(row['Dosis/ha']) else None,
                    "unidad": str(row['Unidad']).strip() if pd.notna(row['Unidad']) else "",
                    "tipo": str(row['Tipo']).strip() if pd.notna(row['Tipo']) else "",
                    "pa": str(row['p.a']).strip() if pd.notna(row['p.a']) else "",
                    "total": float(row['Total']) if pd.notna(row['Total']) else None,
                })

        ordenes_todas.append({
            "fecha": fecha.strftime("%Y-%m-%d") if pd.notna(fecha) else None,
            "fecha_realizado": primera['Fecha realizado'].strftime("%Y-%m-%d") if pd.notna(primera['Fecha realizado']) else None,
            "establecimiento": est,
            "lote": lote,
            "superficie": float(primera['Sup']) if pd.notna(primera['Sup']) else None,
            "campaña": str(primera['Campaña']).strip() if pd.notna(primera['Campaña']) else "",
            "cultivo_abrev": cultivo_abrev,
            "cultivo": cultivos_dict.get(cultivo_abrev, cultivo_abrev),
            "antecesor_abrev": antecesor_abrev,
            "antecesor": cultivos_dict.get(antecesor_abrev, antecesor_abrev),
            "tipo_labor": tipo_labor,
            "items": items,
            "n_productos": len(items),
            "contratista": str(primera.get('Contratista/Proovedorr', '')).strip() if pd.notna(primera.get('Contratista/Proovedorr')) else "",
            "real_presup": "",
            "observaciones": "",
        })

    # --- Planificación 26/27 por ambiente ---
    df_plan = pd.read_excel(buf_agricultura, sheet_name="Info")
    df_plan = df_plan.dropna(how='all')
    df_plan = df_plan[df_plan['Campo'].notna()]

    ambientes_por_lote = {}
    lotes_agricolas = set()  # <-- UNIVERSO DE LOTES VÁLIDOS

    for _, row in df_plan.iterrows():
        campo = str(row['Campo']).strip() if pd.notna(row['Campo']) else ""
        lote = str(row['Lote']).strip() if pd.notna(row['Lote']) else ""
        if not campo or not lote:
            continue
        key = (campo, lote)
        lotes_agricolas.add(key)  # marca este lote como válido
        ambientes_por_lote.setdefault(key, [])
        amb_nombre = str(row['Ambiente Regional RIDZO']).strip() if pd.notna(row['Ambiente Regional RIDZO']) else None
        ambientes_por_lote[key].append({
            "ambiente": amb_nombre,
            "campaña": str(row['Campaña']).strip() if pd.notna(row['Campaña']) else "",
            "superficie": float(row['Superficie Sembrada']) if pd.notna(row['Superficie Sembrada']) else None,
            "cultivo": str(row['Cultivo']).strip() if pd.notna(row['Cultivo']) else "",
            "destino": str(row['Destino']).strip() if pd.notna(row['Destino']) else "",
            "antecesor": str(row['Antecesor']).strip() if pd.notna(row['Antecesor']) else "",
            "fecha_siembra": row['Fecha de Siembra'].strftime("%Y-%m-%d") if pd.notna(row['Fecha de Siembra']) and hasattr(row['Fecha de Siembra'], 'strftime') else None,
            "semillero": str(row['Semillero']).strip() if pd.notna(row['Semillero']) else "",
            "genetica": str(row['Genética']).strip() if pd.notna(row['Genética']) else "",
            "densidad_recomendada": float(row['Densidad siembra Recomendada']) if pd.notna(row['Densidad siembra Recomendada']) else None,
            "rinde_esperado": float(row['Rinde Esperado (kg/ha)']) if pd.notna(row['Rinde Esperado (kg/ha)']) else None,
            "fertilizante_p": str(row['Fertilizante P Producto 1']).strip() if pd.notna(row['Fertilizante P Producto 1']) else "",
            "dosis_fert_p": float(row['Dosis Fert. P Producto 1']) if pd.notna(row['Dosis Fert. P Producto 1']) else None,
            "fertilizante_n": str(row['Fertilizante N Producto 2']).strip() if pd.notna(row['Fertilizante N Producto 2']) else "",
            "dosis_fert_n": float(row['Dosis 2']) if pd.notna(row['Dosis 2']) else None,
        })

    # === FILTRAR ÓRDENES DE TRABAJO ===
    # Solo mantener las órdenes cuyo (establecimiento, lote) esté en el universo agrícola
    ordenes = []
    ordenes_ignoradas = []
    for o in ordenes_todas:
        key = (o['establecimiento'], o['lote'])
        if key in lotes_agricolas:
            ordenes.append(o)
        else:
            ordenes_ignoradas.append(o)

    # Log de tratamientos ignorados (para diagnóstico)
    if ordenes_ignoradas:
        lotes_ignorados = {}
        for o in ordenes_ignoradas:
            key = f"{o['establecimiento']}/{o['lote']}"
            lotes_ignorados[key] = lotes_ignorados.get(key, 0) + 1
        print(f"\n⚠ {len(ordenes_ignoradas)} tratamientos ignorados por no corresponder a lote agrícola:")
        for lote_key, count in sorted(lotes_ignorados.items()):
            print(f"    {lote_key}: {count} tratamiento(s)")
    else:
        print("\n✓ Todos los tratamientos corresponden a lotes agrícolas")

    ordenes.sort(key=lambda x: x['fecha'] or '', reverse=True)

    # === CONSTRUIR LISTA DE LOTES ===
    # Universo = SOLO lotes de planilla Agricultura (no unión)
    lotes_lista = []
    for (est, lote_code) in sorted(lotes_agricolas):
        ambientes = ambientes_por_lote.get((est, lote_code), [])
        ordenes_lote = [o for o in ordenes if o['establecimiento'] == est and o['lote'] == lote_code]

        principal = None
        sub_ambientes = []
        sin_ambiente = []
        for amb in ambientes:
            if amb['ambiente'] is None or amb['ambiente'] == '':
                sin_ambiente.append(amb)
            else:
                sub_ambientes.append(amb)

        if sub_ambientes:
            principal = max(sub_ambientes, key=lambda a: a['superficie'] or 0)
        elif sin_ambiente:
            principal = sin_ambiente[0]
        else:
            # Lote está en Agricultura pero sin datos de planificación cargados
            principal = {
                "ambiente": None, "campaña": "",
                "superficie": 0, "cultivo": "—",
                "destino": "", "antecesor": "",
                "fecha_siembra": None, "semillero": "", "genetica": "",
                "densidad_recomendada": None, "rinde_esperado": None,
                "fertilizante_p": "", "dosis_fert_p": None,
                "fertilizante_n": "", "dosis_fert_n": None,
            }

        if sub_ambientes:
            sup_total = sum(a['superficie'] or 0 for a in sub_ambientes)
        else:
            sup_total = principal['superficie']

        lotes_lista.append({
            "establecimiento": est,
            "lote": lote_code,
            "superficie": sup_total,
            "cultivo_actual": principal['cultivo'],
            "destino_actual": principal.get('destino', ''),
            "antecesor": principal['antecesor'],
            "campaña_planificada": principal['campaña'],
            "fecha_siembra_planificada": principal['fecha_siembra'],
            "semillero": principal['semillero'],
            "genetica": principal['genetica'],
            "fertilizante_p": principal['fertilizante_p'],
            "dosis_fert_p": principal['dosis_fert_p'],
            "fertilizante_n": principal['fertilizante_n'],
            "dosis_fert_n": principal['dosis_fert_n'],
            "tiene_ambientes": len(sub_ambientes) > 0,
            "n_ambientes": len(sub_ambientes),
            "ambientes": sub_ambientes,
            "ordenes_trabajo": ordenes_lote,
        })

    data = {
        "empresa": "La Madrugada",
        "generado_en": ahora_argentina().strftime("%Y-%m-%d %H:%M"),
        "lotes": lotes_lista,
        "ordenes_trabajo": ordenes,
        "cultivos_dict": cultivos_dict,
        "totales": {
            "n_lotes": len(lotes_lista),
            "n_ordenes": len(ordenes),
            "establecimientos": sorted(list(set(l['establecimiento'] for l in lotes_lista))),
            "superficie_total": sum(l['superficie'] or 0 for l in lotes_lista),
        }
    }
    return data


# =========================================================================
# 3. Generar HTML desde template
# =========================================================================
def generar_html(data):
    from pathlib import Path
    template_path = Path(__file__).parent / "template.html"
    html = template_path.read_text(encoding="utf-8")
    data_json = json.dumps(data, ensure_ascii=False, separators=(',', ':'), default=str)
    return html.replace("__DATA_PLACEHOLDER__", data_json)


# =========================================================================
# 4. Main
# =========================================================================
def main():
    print(f"=== Actualización La Madrugada · {ahora_argentina().strftime('%Y-%m-%d %H:%M:%S')} (ART) ===\n")

    access_token = get_access_token()

    path_cultivos = f"{DROPBOX_FOLDER}/{PLANILLA_CULTIVOS}"
    path_agricultura = f"{DROPBOX_FOLDER}/{PLANILLA_AGRICULTURA}"

    buf_cultivos = download_file(access_token, path_cultivos)
    buf_agricultura = download_file(access_token, path_agricultura)

    print("\n→ Procesando datos...")
    data = procesar_planillas(buf_cultivos, buf_agricultura)
    print(f"\n  ✓ {data['totales']['n_lotes']} lotes agrícolas")
    print(f"  ✓ {data['totales']['n_ordenes']} órdenes de trabajo (solo en lotes agrícolas)")
    print(f"  ✓ {data['totales']['superficie_total']} hectáreas")

    print("\n→ Generando HTML...")
    html = generar_html(data)
    output_path = "index.html"
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"  ✓ {output_path} ({len(html):,} bytes)")

    print("\n=== Actualización OK ===")


if __name__ == "__main__":
    main()
