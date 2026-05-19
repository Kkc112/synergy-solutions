import csv

def wa(raw):
    if not raw:
        return ""
    d = "".join(c for c in raw if c.isdigit())
    if not d:
        return ""
    if d.startswith("52"):
        return f"https://wa.me/{d}"
    return f"https://wa.me/52{d}"

LEADS = [
    # ── CDMX ────────────────────────────────────────────────────────────
    ("Hughes Events",              "Corporativos / General",   "CDMX","CDMX","5554552934","agencia-de-eventos-cdmx","20+ años. Corporativos y sociales."),
    ("iRHOD Producer",             "Producción general",       "CDMX","CDMX","5539858314","","Producción audiovisual y eventos."),
    ("Eufra México",               "Social / Corporativo",     "CDMX","CDMX","5574834716","eufra-mexico.ueniweb.com",""),
    ("MICE México (Neruc)",        "MICE / Corporativos",      "CDMX","CDMX","5589509577","micemexico.events","Reuniones, incentivos, congresos."),
    ("J&J Eventos Especiales",     "Social / Empresarial",     "CDMX","CDMX","5556054155","jjeventos.com","20+ años. Pitagoras 916, Del Valle."),
    ("Banquetes Ambrosia",         "Social / Corporativo",     "CDMX","CDMX","5546249800","ambrosia.mx","Salones propios. Periférico Sur 3395."),
    ("SARAO",                      "Bodas / Fiestas / Corpor.","CDMX (multi)","CDMX","5552712032","sarao.com.mx","Líder 20+ años. Fiestas temáticas."),
    ("Imagic Group",               "Congresos / Corporativos", "CDMX","CDMX","5556694973","imagicgroup.com","Insurgentes Sur 428. info@imagicgroup.com"),
    ("Eventos VB",                 "Producción audiovisual",   "CDMX","CDMX","5512375887","eventosvb.com","Tecnología innovadora en espectáculos."),
    ("Fiestas y Eventos EGO",      "Social / Empresarial",     "CDMX","CDMX","5565893354","fiestasyeventosego.com","Asesoría profesional en coordinación."),
    ("Intermeeting",               "Congresos / Convenciones", "CDMX","CDMX","5556630402","intermeeting.com.mx","19 años en congresos y exposiciones."),
    ("Social Casino Eventos",      "Social / Entretenimiento", "CDMX Benito Juárez","CDMX","5537091828","","Planificación y organización."),
    ("Pilares Eventos",            "Social / Corporativo",     "CDMX","CDMX","5632630657","","Decoración + entretenimiento completo."),
    ("Ángel Eventos",              "Social / Bodas",           "CDMX Coyoacán","CDMX","5544704037","","200+ eventos realizados."),
    ("Gruposmart Agencia Creativa","Corporativos / Creativos", "CDMX Cuauhtémoc","CDMX","5631218424","","Proyectos creativos y operativos."),
    ("Syncretic",                  "Corporativos / Híbridos",  "CDMX","CDMX","5564688368","","Presencial + híbrido + virtual."),
    ("Happy Planners",             "Bodas / Social",           "CDMX Benito Juárez","CDMX","5624406148","","Planificación de eventos diversos."),
    ("Proeesa Eventos",            "Empresarial / Social",     "CDMX Cuauhtémoc","CDMX","5523281428","","Académicos + empresariales + sociales."),
    ("Vansi Experience Marketing", "Experiencias de marca",    "CDMX Álvaro Obregón","CDMX","5565382198","","Marketing experiencial personalizado."),
    ("NAYA events",                "Social / Empresarial",     "CDMX Cuauhtémoc","CDMX","5510145574","",""),
    ("Starfall",                   "Asesoría / Organización",  "CDMX V. Carranza","CDMX","5537229721","",""),
    ("Bauri Event Planning",       "Social / Empresarial",     "CDMX Tlalpan","CDMX","5529539113","",""),
    ("MH Producciones",            "Corporativos / Ferias",    "CDMX Miguel Hidalgo","CDMX","5555126896","","Lanzamientos, ferias, corporativos."),
    ("Congrexpo",                  "Congresos / Escenografía", "CDMX","CDMX","5527556706","","Especialistas en escenografía para congresos."),
    ("El Antro House Eventos",     "Todo incluido",            "CDMX Álvaro Obregón","CDMX","5546200478","","Paquetes todo incluido."),
    ("HARI Eventos S.A de C.V",   "Corporativos / Social",    "CDMX Miguel Hidalgo","CDMX","5523987547","","Coordinación y planificación."),
    ("Karams Producciones",        "Producción / Experiencias","CDMX Cuauhtémoc","CDMX","5538942470","","Eventos memorables."),
    ("Valermo Experience Group",   "Lujo / Experiencias",      "CDMX","CDMX","5549386854","grupoavanttia.com","Experiencias de lujo personalizadas."),
    ("Grupo Avanttia",             "AV / Producción integral", "CDMX","CDMX","5575751368","grupoavanttia.com","Audio, video e iluminación. Cuauhtémoc 722."),
    ("Event Planner México",       "Bodas / Corporativos",     "CDMX","CDMX","5561257703","eventplannermexico.mx","200+ bodas. Paquetes $85k-$200k MXN."),
    ("Yaber Supreme Events",       "Bodas / Gala",             "CDMX","CDMX","5527053894","yabersupreme.com","Bodas con elegancia como eje central."),
    ("Revery Eventos",             "Social / Diseño contemp.", "CDMX","CDMX","5519894364","instagram.com/revery_eventos_especiales","Flexibilidad y resolución de crisis."),
    ("Protocolo Eventos",          "Corporativos técnicos",    "CDMX","CDMX","5578254081","","Coordinación milimétrica de programas."),
    ("IMAGINE Diseño y Producción","Producción gran escala",   "CDMX","CDMX","5555124420","imaginedype.com","Escenografías que transforman espacios."),
    ("Aster Events",               "Bodas / Social",           "CDMX","CDMX","5591835496","asterevent.com","Estilo minimalista y natural."),
    ("Nappyclub",                  "Baby shower / Social",     "CDMX","CDMX","5551060519","","Dinámicas y entretenimiento."),
    # ── Guadalajara / Jalisco ────────────────────────────────────────────
    ("ARTV ARTPHEST S.A.",         "Shows / Entretenimiento",  "Zapopan","JAL","3331294964","artv.com.mx","38 años. Conciertos y ferias."),
    ("AVM Organización",           "Social / Cultural / Dep.", "Guadalajara","JAL","3336795717","eventosavm.com","Social, cultural y deportivo."),
    ("B&C Eventos",                "Bodas / Banquetes",        "Zapopan","JAL","3331650590","banquetesycoordinaciondeeventos.com","Nacional e internacional."),
    ("SUNSET EVENTCOACHING",       "Corporativos / Convenciones","Zapopan","JAL","3315629424","sunseteventos.net","Convenciones y lanzamientos."),
    ("Event Planner Guadalajara",  "General",                  "Guadalajara","JAL","3314958653","planeadoresdeeventos.com",""),
    ("Coordinación De Eventos GDL","General",                  "Guadalajara","JAL","3336325933","","Lomas De Guevara, Gral. Eulogio Parra."),
    # ── Monterrey / Nuevo León ────────────────────────────────────────────
    ("Carpa Real",                 "Bodas / Sociales",         "Monterrey","NL","8183114444","carpareal.com.mx",""),
    ("DMC Monterrey",              "Congresos / Incentivos",   "Monterrey","NL","8181918142","dmcmonterrey.com","20 años. info@dmcmonterrey.com"),
    ("Aztec Consultoria Empresarial","Corporativos",           "Monterrey","NL","","","5.0 estrellas (14 reseñas) StarOfService."),
    ("Plus Eventos Y Promociones", "General",                  "Monterrey","NL","","","StarOfService Monterrey."),
    ("Lorena Puente Bodas y Eventos","Bodas",                  "Monterrey","NL","","","Desde $15,000 MXN."),
    ("Brandon De León Event Designer","Diseño de eventos",     "Monterrey","NL","","","4.0 estrellas. Event designer."),
    # ── Otros estados ────────────────────────────────────────────────────
    ("MOZÉ Centro de Eventos",     "Corporativos / Social",    "Tuxtla Gutiérrez","Chiapas","9611465809","mozeeventos.com","Salones verdes y corporativos."),
    ("ARIM Banquetes",             "Banquetes / Empresarial",  "La Piedad","Michoacán","3525241002","arim.com.mx","25+ años lider en La Piedad."),
]

FIELDNAMES = [
    "ID","Empresa","Tipo","Ciudad","Estado",
    "Teléfono","WhatsApp_link","Sitio_web","Especialidad_notas",
    "Fecha_llamada","Estado_llamada","Demo_enviada","Cerrado","Notas_seguimiento"
]

rows = []
for i, (empresa, tipo, ciudad, estado, tel, web, notas) in enumerate(LEADS, 1):
    rows.append({
        "ID": i,
        "Empresa": empresa,
        "Tipo": tipo,
        "Ciudad": ciudad,
        "Estado": estado,
        "Teléfono": tel,
        "WhatsApp_link": wa(tel),
        "Sitio_web": web,
        "Especialidad_notas": notas,
        "Fecha_llamada": "",
        "Estado_llamada": "",
        "Demo_enviada": "",
        "Cerrado": "",
        "Notas_seguimiento": ""
    })

path = "/home/user/synergy-solutions/Lumio-Eventos-Mexico.csv"
with open(path, "w", newline="", encoding="utf-8-sig") as f:
    writer = csv.DictWriter(f, fieldnames=FIELDNAMES)
    writer.writeheader()
    writer.writerows(rows)

print(f"SAVED: {path}")
print(f"TOTAL LEADS: {len(rows)}")
con_tel = sum(1 for r in rows if r["Teléfono"])
sin_tel = len(rows) - con_tel
print(f"Con teléfono: {con_tel} | Sin teléfono: {sin_tel}")
from collections import Counter
estados = Counter(r["Estado"] for r in rows)
print("Por estado:", dict(estados))
