"""
Buscador multi-portal de empleos remotos en español (Colombia / LATAM)
Salida: Archivo Excel (.xlsx) con pestañas por categoría + Resumen a Telegram.
Sirve como backup organizado con todos los enlaces, sueldos y descripciones.
"""

import os
import json
import re
import html
import time
import urllib.request
import urllib.parse
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Cargar variables desde .env local si existe
def cargar_env():
    env_path = os.path.join(BASE_DIR, ".env")
    if os.path.exists(env_path):
        with open(env_path, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith("#") and "=" in line:
                    key, val = line.split("=", 1)
                    val = val.strip().strip('"').strip("'")
                    os.environ.setdefault(key.strip(), val)

cargar_env()

# --- CONFIGURACIÓN ---
TELEGRAM_BOT_TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN", "").strip()
TELEGRAM_CHAT_ID = os.environ.get("TELEGRAM_CHAT_ID", "").strip()

HISTORIAL_JSON = os.path.join(BASE_DIR, "empleos_vistos.json")
CARPETA_REPORTES = os.path.join(BASE_DIR, "reportes")
MAX_IDS_HISTORIAL = 2000

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36"
}

GRUPOS_CONFIG = {
    "PYTHON": {"icono": "🐍", "hoja": "🐍 Python"},
    "INGENIERÍA DE DATOS": {"icono": "📊", "hoja": "📊 Ingeniería de Datos"},
    "DATA ENTRY": {"icono": "⌨️", "hoja": "⌨️ Data Entry"},
    "INGENIERÍA DE SOFTWARE": {"icono": "💻", "hoja": "💻 Software"}
}

def clasificar_oferta(titulo, tags=None):
    """Clasifica la vacante con vocabulario ampliado para Colombia y LATAM."""
    tit = (titulo or "").lower()
    tags_text = " ".join(tags).lower() if tags else ""
    
    # 1. DATA ENTRY Y DIGITACIÓN
    data_entry_kws = [
        "data entry", "capturista", "transcripcion", "transcripción",
        "digitador", "digitadora", "digitación", "digitacion",
        "captura de datos", "data clerk", "data operator", "asistente de datos",
        "ingreso de datos", "auxiliar de datos", "transcriptor", "transcriptor(a)"
    ]
    if any(k in tit for k in data_entry_kws):
        return "DATA ENTRY"
        
    # 2. INGENIERÍA Y ANÁLISIS DE DATOS
    datos_kws = [
        "data engineer", "ingeniero de datos", "ingeniera de datos", "ingeniería de datos",
        "analista de datos", "data analyst", "científico de datos", "cientifico de datos",
        "data scientist", "data science", "etl", "big data", "data pipeline",
        "data warehouse", "analytics engineer", "bases de datos", "business intelligence",
        "power bi", "bi developer", "sql developer", "dba", "administrador de base de datos"
    ]
    if any(k in tit for k in datos_kws):
        return "INGENIERÍA DE DATOS"
        
    # 3. PYTHON
    python_kws = ["python", "django", "flask", "fastapi", "desarrollador python", "python developer"]
    if any(k in tit for k in python_kws):
        return "PYTHON"
        
    # 4. INGENIERÍA DE SOFTWARE Y SISTEMAS
    software_kws = [
        "software engineer", "software developer", "backend", "full stack", "fullstack",
        "frontend", "developer", "ingeniero de software", "desarrollador", "desarrolladora",
        "programador", "programadora", "ingeniero de sistemas", "ingeniera de sistemas",
        "desarrollador web", "web developer", "devops", "qa engineer", "qa",
        "soporte técnico", "soporte ti", "arquitecto de software"
    ]
    if any(k in tit for k in software_kws):
        return "INGENIERÍA DE SOFTWARE"
        
    # 5. Evaluación por etiquetas complementarias
    if tags_text:
        if any(k in tags_text for k in ["data entry", "digitador", "capturista"]):
            return "DATA ENTRY"
        if any(k in tags_text for k in ["data engineer", "etl", "big data", "data warehouse", "data analyst"]):
            return "INGENIERÍA DE DATOS"
        if "python" in tags_text:
            return "PYTHON"
        if any(k in tags_text for k in ["software engineer", "backend", "fullstack", "full stack", "desarrollo"]):
            return "INGENIERÍA DE SOFTWARE"
            
    return None

def limpiar_texto_html(texto_html, max_caracteres=350):
    """Limpia etiquetas HTML y prepara descripciones legibles para celdas de Excel."""
    if not texto_html:
        return "Descripción detallada disponible en el enlace de la oferta."
    texto = re.sub(r'<[^>]+>', ' ', texto_html)
    texto = html.unescape(texto)
    texto = re.sub(r'\s+', ' ', texto).strip()
    if len(texto) > max_caracteres:
        return texto[:max_caracteres].rstrip() + "..."
    return texto

def cargar_empleos_vistos():
    """Carga los IDs de empleos ya notificados anteriormente."""
    if os.path.exists(HISTORIAL_JSON):
        try:
            with open(HISTORIAL_JSON, "r", encoding="utf-8") as f:
                return set(json.load(f))
        except Exception:
            return set()
    return set()

def guardar_empleos_vistos(vistos):
    """Guarda los IDs aplicando el límite rotativo."""
    try:
        lista_ids = list(vistos)
        if len(lista_ids) > MAX_IDS_HISTORIAL:
            lista_ids = lista_ids[-MAX_IDS_HISTORIAL:]
        with open(HISTORIAL_JSON, "w", encoding="utf-8") as f:
            json.dump(lista_ids, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"Error guardando historial: {e}")

# --- FUENTES DE EMPLEO ---
def obtener_ofertas_linkedin():
    """Consulta LinkedIn para Colombia y LATAM."""
    queries = [
        ("python", "PYTHON"),
        ("desarrollador%20python", "PYTHON"),
        ("data%20engineer", "INGENIERÍA DE DATOS"),
        ("analista%20de%20datos", "INGENIERÍA DE DATOS"),
        ("data%20entry", "DATA ENTRY"),
        ("digitador", "DATA ENTRY"),
        ("software%20engineer", "INGENIERÍA DE SOFTWARE"),
        ("desarrollador%20software", "INGENIERÍA DE SOFTWARE")
    ]
    
    ofertas = []
    vistos_locales = set()
    
    for kw, grupo_defecto in queries:
        try:
            url = f"https://www.linkedin.com/jobs-guest/jobs/api/seeMoreJobPostings/search?keywords={kw}&location=Colombia&f_WT=2"
            req = urllib.request.Request(url, headers=HEADERS)
            with urllib.request.urlopen(req, timeout=12) as resp:
                html_li = resp.read().decode("utf-8")
                
            cards = re.findall(r'<div class="base-card[^"]*"[^>]*>(.*?)</div>\s*</div>', html_li, re.DOTALL)
            for c in cards:
                link_m = re.search(r'<a[^>]+href="([^"]+)"[^>]*class="[^"]*base-card__full-link[^"]*"', c) or re.search(r'<a[^>]*class="[^"]*base-card__full-link[^"]*"[^>]+href="([^"]+)"', c)
                if not link_m:
                    continue
                url_limpia = link_m.group(1).split('?')[0]
                
                id_m = re.search(r'-(\d+)$', url_limpia)
                job_id = f"li_{id_m.group(1)}" if id_m else f"li_{url_limpia}"
                
                if job_id in vistos_locales:
                    continue
                vistos_locales.add(job_id)
                
                title_m = re.search(r'<h3 class="base-search-card__title">\s*([^<]+)\s*</h3>', c)
                title = title_m.group(1).strip() if title_m else kw.replace('%20', ' ').title()
                
                comp_m = re.search(r'<h4 class="base-search-card__subtitle">\s*<a[^>]*>([^<]+)</a>', c) or re.search(r'<h4 class="base-search-card__subtitle">\s*([^<]+)\s*</h4>', c)
                company = comp_m.group(1).strip() if comp_m else "Empresa en LinkedIn"
                
                loc_m = re.search(r'<span class="job-search-card__location">\s*([^<]+)\s*</span>', c)
                loc = loc_m.group(1).strip() if loc_m else "Colombia (Remoto)"
                
                grupo = clasificar_oferta(title) or grupo_defecto
                
                ofertas.append({
                    "id": job_id,
                    "portal": "LinkedIn",
                    "grupo": grupo,
                    "titulo": title,
                    "empresa": company,
                    "lugar": loc,
                    "remoto": "Sí (100% Remoto)",
                    "sueldo": "Ver en LinkedIn",
                    "descripcion": f"Vacante remota publicada en LinkedIn para {title}.",
                    "url": url_limpia
                })
        except Exception as e:
            print(f"Error en LinkedIn ({kw}): {e}")
            
    return ofertas

def obtener_ofertas_computrabajo():
    """Consulta ofertas remotas en CompuTrabajo Colombia."""
    urls = [
        ("https://co.computrabajo.com/trabajo-de-python-remoto", "PYTHON"),
        ("https://co.computrabajo.com/trabajo-de-datos-remoto", "INGENIERÍA DE DATOS"),
        ("https://co.computrabajo.com/trabajo-de-analista-de-datos-remoto", "INGENIERÍA DE DATOS"),
        ("https://co.computrabajo.com/trabajo-de-ingeniero-de-datos-remoto", "INGENIERÍA DE DATOS"),
        ("https://co.computrabajo.com/trabajo-de-data-entry-remoto", "DATA ENTRY"),
        ("https://co.computrabajo.com/trabajo-de-digitador-remoto", "DATA ENTRY"),
        ("https://co.computrabajo.com/trabajo-de-digitadora-remoto", "DATA ENTRY"),
        ("https://co.computrabajo.com/trabajo-de-desarrollador-remoto", "INGENIERÍA DE SOFTWARE"),
        ("https://co.computrabajo.com/trabajo-de-programador-remoto", "INGENIERÍA DE SOFTWARE"),
        ("https://co.computrabajo.com/trabajo-de-ingeniero-de-sistemas-remoto", "INGENIERÍA DE SOFTWARE")
    ]
    
    ofertas = []
    vistos_locales = set()
    
    for url, grupo_defecto in urls:
        try:
            req = urllib.request.Request(url, headers=HEADERS)
            with urllib.request.urlopen(req, timeout=12) as resp:
                html_ct = resp.read().decode("utf-8")
                
            articles = re.findall(r'<article[^>]*id="([0-9A-F]+)"[^>]*>(.*?)</article>', html_ct, re.DOTALL)
            for aid, art in articles:
                job_id = f"ct_{aid}"
                if job_id in vistos_locales:
                    continue
                vistos_locales.add(job_id)
                
                t_m = re.search(r'<a class="js-o-link[^"]*"[^>]*href="([^"]*)"[^>]*>([^<]+)</a>', art)
                if not t_m:
                    continue
                    
                title = html.unescape(t_m.group(2).strip())
                link_rel = t_m.group(1).split('#')[0]
                link = f"https://co.computrabajo.com{link_rel}"
                
                comp_m = re.search(r'<p class="fs16[^"]*">\s*<a[^>]*>([^<]+)</a>', art) or re.search(r'<p class="fs16[^"]*">([^<]+)</p>', art)
                company = html.unescape(comp_m.group(1).strip()) if comp_m else "Empresa Confidencial"
                
                sal_m = re.search(r'<span class="dIB mr10">([^<]+)</span>', art)
                salary = html.unescape(sal_m.group(1).strip()) if sal_m else "No especificado"
                
                loc_m = re.search(r'<p class="fs13 text-secondary[^"]*">([^<]+)</p>', art)
                loc = html.unescape(loc_m.group(1).strip()) if loc_m else "Colombia (Remoto)"
                
                grupo = clasificar_oferta(title) or grupo_defecto
                
                ofertas.append({
                    "id": job_id,
                    "portal": "CompuTrabajo",
                    "grupo": grupo,
                    "titulo": title,
                    "empresa": company,
                    "lugar": f"{loc} / Remoto",
                    "remoto": "Sí (Remoto)",
                    "sueldo": salary,
                    "descripcion": f"Oferta remota en CompuTrabajo Colombia: {title}.",
                    "url": link
                })
        except Exception as e:
            print(f"Error en CompuTrabajo ({url}): {e}")
            
    return ofertas

def obtener_ofertas_getonbrd():
    """Consulta Get on Board con categorías tech y de datos."""
    endpoints = [
        "https://www.getonbrd.com/api/v0/categories/programming/jobs?per_page=30",
        "https://www.getonbrd.com/api/v0/search/jobs?query=data+engineer",
        "https://www.getonbrd.com/api/v0/search/jobs?query=python",
        "https://www.getonbrd.com/api/v0/search/jobs?query=analista+datos",
        "https://www.getonbrd.com/api/v0/search/jobs?query=data+entry"
    ]
    
    ofertas = []
    vistos_locales = set()
    
    for url in endpoints:
        try:
            req = urllib.request.Request(url, headers=HEADERS)
            with urllib.request.urlopen(req, timeout=12) as response:
                data = json.loads(response.read().decode("utf-8"))
                jobs = data.get("data", [])
                for item in jobs:
                    job_id = f"gob_{item.get('id')}"
                    if job_id in vistos_locales:
                        continue
                    vistos_locales.add(job_id)
                    
                    attrs = item.get("attributes", {})
                    es_remoto = attrs.get("remote", False)
                    modalidad = attrs.get("remote_modality", "")
                    if not es_remoto and modalidad != "fully_remote":
                        continue
                        
                    title = attrs.get("title", "")
                    tags = attrs.get("tags_names", [])
                    grupo = clasificar_oferta(title, tags)
                    
                    if grupo:
                        desc = limpiar_texto_html(attrs.get("description", ""))
                        pais = attrs.get("country", "América Latina / Remoto")
                        
                        min_sal = attrs.get("min_salary")
                        max_sal = attrs.get("max_salary")
                        if min_sal and max_sal:
                            sueldo = f"${min_sal:,} - ${max_sal:,} USD / mes"
                        elif min_sal:
                            sueldo = f"Desde ${min_sal:,} USD / mes"
                        else:
                            sueldo = "No especificado en la oferta"
                            
                        ofertas.append({
                            "id": job_id,
                            "portal": "Get on Board",
                            "grupo": grupo,
                            "titulo": title,
                            "empresa": attrs.get("company", {}).get("data", {}).get("attributes", {}).get("name", "Confidencial"),
                            "lugar": f"{pais} (Remoto)",
                            "remoto": "Sí (100% Remoto)",
                            "sueldo": sueldo,
                            "descripcion": desc,
                            "url": attrs.get("url", f"https://www.getonbrd.com/jobs/{item.get('id')}")
                        })
        except Exception as e:
            print(f"Error en Get on Board ({url}): {e}")
            
    return ofertas

def obtener_ofertas_remotive():
    """Consulta Remotive filtrando vacantes disponibles para LATAM y en español."""
    endpoints = [
        "https://remotive.com/api/remote-jobs?category=software-dev&limit=40",
        "https://remotive.com/api/remote-jobs?category=data&limit=40",
        "https://remotive.com/api/remote-jobs?search=python"
    ]
    
    ofertas = []
    vistos_locales = set()
    regiones_ok = ["latam", "latin america", "worldwide", "colombia", "cualquier país", "remote"]
    
    for url in endpoints:
        try:
            req = urllib.request.Request(url, headers=HEADERS)
            with urllib.request.urlopen(req, timeout=12) as resp:
                data = json.loads(resp.read().decode("utf-8"))
                for job in data.get("jobs", []):
                    job_id = f"rem_{job.get('id', job.get('url'))}"
                    if job_id in vistos_locales:
                        continue
                    vistos_locales.add(job_id)
                    
                    lugar = (job.get("candidate_required_location") or "").lower()
                    if not any(r in lugar for r in regiones_ok) and lugar != "":
                        continue
                        
                    title = job.get("title", "")
                    tags = [t.lower() for t in job.get("tags", [])]
                    grupo = clasificar_oferta(title, tags)
                    
                    if grupo:
                        desc = limpiar_texto_html(job.get("description", ""))
                        ofertas.append({
                            "id": job_id,
                            "portal": "Remotive",
                            "grupo": grupo,
                            "titulo": title,
                            "empresa": job.get("company_name", "Confidencial"),
                            "lugar": job.get("candidate_required_location") or "Remoto Mundial / LATAM",
                            "remoto": "Sí (100% Remoto)",
                            "sueldo": job.get("salary") or "No especificado en la oferta",
                            "descripcion": desc,
                            "url": job.get("url", "")
                        })
        except Exception as e:
            print(f"Error en Remotive ({url}): {e}")
            
    return ofertas

# --- GENERADOR DE ARCHIVO EXCEL (.xlsx) PROFESIONAL ---
def generar_excel_empleos(agrupadas, ruta_archivo):
    """
    Crea un archivo Excel profesional con estilo corporativo:
    - Pestaña 'Todas las Vacantes'
    - Pestañas individuales por Grupo (Python, Datos, Data Entry, Software)
    - Hipervínculos directos a las ofertas
    - Columnas auto-ajustadas
    """
    wb = openpyxl.Workbook()
    ws_todas = wb.active
    ws_todas.title = "Todas las Vacantes"
    
    header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    cell_font = Font(name="Calibri", size=10)
    
    headers = [
        "Grupo", "Puesto", "Empresa", "Portal", "Lugar",
        "¿Remoto?", "Sueldo", "Enlace de Postulación", "Descripción", "Fecha"
    ]
    
    # Crear pestañas por grupo
    pestanas = [("Todas las Vacantes", ws_todas)]
    for grp in ["PYTHON", "INGENIERÍA DE DATOS", "DATA ENTRY", "INGENIERÍA DE SOFTWARE"]:
        if grp in agrupadas:
            nombre_pestana = GRUPOS_CONFIG.get(grp, {}).get("hoja", grp)[:31]
            ws_grp = wb.create_sheet(title=nombre_pestana)
            pestanas.append((grp, ws_grp))
            
    for nombre_filtro, ws in pestanas:
        ws.append(headers)
        ws.row_dimensions[1].height = 26
        
        for col_idx in range(1, len(headers) + 1):
            cell = ws.cell(row=1, column=col_idx)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            
        if nombre_filtro == "Todas las Vacantes":
            lista_ofertas = [of for g, ofs in agrupadas.items() for of in ofs]
        else:
            lista_ofertas = agrupadas.get(nombre_filtro, [])
            
        for of in lista_ofertas:
            row_data = [
                of.get("grupo", ""),
                of.get("titulo", ""),
                of.get("empresa", ""),
                of.get("portal", ""),
                of.get("lugar", ""),
                of.get("remoto", "Sí"),
                of.get("sueldo", "No especificado"),
                of.get("url", ""),
                of.get("descripcion", ""),
                datetime.now().strftime("%Y-%m-%d %H:%M")
            ]
            ws.append(row_data)
            current_row = ws.max_row
            
            # Estilos de celda
            for c_idx in range(1, len(row_data) + 1):
                c = ws.cell(row=current_row, column=c_idx)
                c.font = cell_font
                if c_idx in (1, 4, 6):
                    c.alignment = Alignment(horizontal="center", vertical="center")
                else:
                    c.alignment = Alignment(vertical="center")
                    
            # Hipervínculo clickeable en columna de Enlace
            link_cell = ws.cell(row=current_row, column=8)
            link_cell.value = "Postularme aquí"
            link_cell.hyperlink = of.get("url", "")
            link_cell.font = Font(name="Calibri", size=10, color="0563C1", underline="single")
            link_cell.alignment = Alignment(horizontal="center", vertical="center")
            
        # Ajustar ancho de columnas
        for col in ws.columns:
            col_letter = get_column_letter(col[0].column)
            if col[0].column == 9:  # Descripción
                ws.column_dimensions[col_letter].width = 45
            elif col[0].column == 8:  # Link
                ws.column_dimensions[col_letter].width = 18
            else:
                max_len = max(len(str(c.value or '')) for c in col)
                ws.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 35)
                
    wb.save(ruta_archivo)
    return ruta_archivo

# --- ENVÍO DE ARCHIVO ADJUNTO A TELEGRAM ---
def enviar_documento_telegram(ruta_archivo, caption=""):
    """Envía el archivo Excel (.xlsx) directamente como documento adjunto a Telegram."""
    if not TELEGRAM_BOT_TOKEN or not TELEGRAM_CHAT_ID:
        print("Error: Credenciales de Telegram no configuradas.")
        return False
        
    boundary = "----WebKitFormBoundary" + os.urandom(8).hex()
    filename = os.path.basename(ruta_archivo)
    
    with open(ruta_archivo, "rb") as f:
        file_bytes = f.read()
        
    parts = []
    # chat_id
    parts.append(f"--{boundary}\r\nContent-Disposition: form-data; name=\"chat_id\"\r\n\r\n{TELEGRAM_CHAT_ID}\r\n".encode('utf-8'))
    
    # caption
    if caption:
        parts.append(f"--{boundary}\r\nContent-Disposition: form-data; name=\"caption\"\r\n\r\n{caption}\r\n".encode('utf-8'))
        parts.append(f"--{boundary}\r\nContent-Disposition: form-data; name=\"parse_mode\"\r\n\r\nHTML\r\n".encode('utf-8'))
        
    # document
    parts.append(
        f"--{boundary}\r\nContent-Disposition: form-data; name=\"document\"; filename=\"{filename}\"\r\n"
        f"Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet\r\n\r\n".encode('utf-8')
    )
    parts.append(file_bytes)
    parts.append(f"\r\n--{boundary}--\r\n".encode('utf-8'))
    
    payload = b''.join(parts)
    url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendDocument"
    req = urllib.request.Request(
        url,
        data=payload,
        headers={"Content-Type": f"multipart/form-data; boundary={boundary}"}
    )
    
    for intento in range(1, 4):
        try:
            with urllib.request.urlopen(req, timeout=40) as resp:
                data = json.loads(resp.read().decode('utf-8'))
                return data.get("ok", False)
        except Exception as e:
            if intento < 3:
                time.sleep(2)
            else:
                print(f"Error enviando documento Excel a Telegram: {e}")
def limpiar_reportes_antiguos(dias_retencion=30):
    """Elimina automáticamente archivos Excel con más de 30 días de antigüedad en la carpeta reportes."""
    if not os.path.exists(CARPETA_REPORTES):
        return
    limite_segundos = time.time() - (dias_retencion * 86400)
    for nombre in os.listdir(CARPETA_REPORTES):
        if nombre.endswith(".xlsx"):
            ruta_completa = os.path.join(CARPETA_REPORTES, nombre)
            try:
                if os.path.getmtime(ruta_completa) < limite_segundos:
                    os.remove(ruta_completa)
                    print(f"Limpieza automática: reporte antiguo eliminado ({nombre})")
            except Exception as e:
                print(f"No se pudo eliminar {nombre}: {e}")

def ejecutar_busqueda():
    """Ejecuta la búsqueda horaria completa y genera el reporte en Excel."""
    fecha_no_militar = datetime.now().strftime("%d/%m/%Y - %I:%M %p")
    print(f"[{fecha_no_militar}] Iniciando escaneo multi-portal de vacantes...")
    
    vistos = cargar_empleos_vistos()
    
    todas_ofertas = []
    todas_ofertas.extend(obtener_ofertas_linkedin())
    todas_ofertas.extend(obtener_ofertas_computrabajo())
    todas_ofertas.extend(obtener_ofertas_getonbrd())
    todas_ofertas.extend(obtener_ofertas_remotive())
    
    # Desduplicar y filtrar nuevas
    nuevas = []
    vistos_esta_corrida = set()
    for of in todas_ofertas:
        if of["id"] not in vistos and of["id"] not in vistos_esta_corrida:
            nuevas.append(of)
            vistos_esta_corrida.add(of["id"])
            
    if not nuevas:
        print("No hay nuevas vacantes en esta hora. Esperando la siguiente ejecución...")
        return
        
    print(f"Se encontraron {len(nuevas)} nueva(s) vacante(s). Generando Excel...")
    
    agrupadas = {}
    for of in nuevas:
        grp = of["grupo"]
        if grp not in agrupadas:
            agrupadas[grp] = []
        agrupadas[grp].append(of)
        vistos.add(of["id"])
        
    guardar_empleos_vistos(vistos)
    
    # Generar archivo Excel con fecha y hora (hora y minuto separados por _)
    os.makedirs(CARPETA_REPORTES, exist_ok=True)
    stamp = datetime.now().strftime("%Y-%m-%d_%H_%M")
    ruta_excel = os.path.join(CARPETA_REPORTES, f"empleos_remotos_{stamp}.xlsx")
    
    generar_excel_empleos(agrupadas, ruta_excel)
    print(f"Excel generado exitosamente: {ruta_excel}")
    
    # Autolimpieza: eliminar archivos de más de 30 días de antigüedad
    limpiar_reportes_antiguos(dias_retencion=30)
    
    # Preparar resumen conciso para Telegram
    total = sum(len(agrupadas[g]) for g in agrupadas)
    c_py = len(agrupadas.get("PYTHON", []))
    c_datos = len(agrupadas.get("INGENIERÍA DE DATOS", []))
    c_entry = len(agrupadas.get("DATA ENTRY", []))
    c_soft = len(agrupadas.get("INGENIERÍA DE SOFTWARE", []))
    
    caption = (
        f"📊 <b>Reporte de Empleos Remotos ({fecha_no_militar})</b>\n\n"
        f"Se encontraron <b>{total} nuevas vacantes</b>:\n"
        f"• 🐍 <b>Python:</b> {c_py}\n"
        f"• 📊 <b>Ingeniería y Análisis de Datos:</b> {c_datos}\n"
        f"• ⌨️ <b>Data Entry y Digitación:</b> {c_entry}\n"
        f"• 💻 <b>Ingeniería de Software:</b> {c_soft}\n\n"
        f"🌐 <i>Portales: LinkedIn, CompuTrabajo, Get on Board y Remotive</i>\n"
        f"📎 <i>Descarga el archivo Excel adjunto para ver todos los enlaces, sueldos y descripciones.</i>"
    )
    
    exito = enviar_documento_telegram(ruta_excel, caption=caption)
    if exito:
        print("Archivo Excel enviado con éxito a Telegram.")
    else:
        print("Fallo al enviar el archivo Excel a Telegram.")

if __name__ == "__main__":
    ejecutar_busqueda()
