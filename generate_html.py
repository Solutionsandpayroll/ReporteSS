"""
Generates index.html — the SS automation form.
Run whenever formato2.xlsx or maestro.xlsx change.
"""
import base64, json

# ─── Assets ───────────────────────────────────────────────────────────────────
with open("logosyp.png", "rb") as f:
    LOGO_B64 = base64.b64encode(f.read()).decode()

MAESTRO = {
    "040267": {"eps": "ALIANSALUD EPS",                                     "afp": "ADMINISTRADORA COLOMBIANA DE PENSIONES COLPENSIONES"},
    "052231": {"eps": "E.P.S SANITAS",                                      "afp": "ADMINISTRADORA COLOMBIANA DE PENSIONES COLPENSIONES"},
    "054755": {"eps": "E.P.S SANITAS",                                      "afp": "OLD MUTUAL FONDO DE PENSIONES OBLIGATORIAS"},
    "055512": {"eps": "E.P.S SANITAS",                                      "afp": "ADMINISTRADORA COLOMBIANA DE PENSIONES COLPENSIONES"},
    "056374": {"eps": "FAMISANAR",                                          "afp": "ADMINISTRADORA COLOMBIANA DE PENSIONES COLPENSIONES"},
    "056918": {"eps": "SALUD TOTAL S.A.",                                   "afp": "ADMINISTRADORA COLOMBIANA DE PENSIONES COLPENSIONES"},
    "058834": {"eps": "NUEVA EPS",                                          "afp": "PORVENIR"},
    "059288": {"eps": "E.P.S SANITAS",                                      "afp": "PROTECCION"},
    "062858": {"eps": "E.P.S SANITAS",                                      "afp": "PROTECCION"},
    "073226": {"eps": "COMPENSAR ENTIDAD PROMOTORA DE SALUD",               "afp": "PORVENIR"},
    "073931": {"eps": "EPS SURA",                                           "afp": "PROTECCION"},
    "074173": {"eps": "EPS SURA",                                           "afp": "ADMINISTRADORA COLOMBIANA DE PENSIONES COLPENSIONES"},
    "074490": {"eps": "EPS SURA",                                           "afp": "OLD MUTUAL FONDO DE PENSIONES OBLIGATORIAS"},
    "075064": {"eps": "EPS SURA",                                           "afp": "COLFONDOS"},
    "075115": {"eps": "FAMISANAR",                                          "afp": "COLFONDOS"},
    "075157": {"eps": "EPS SURA",                                           "afp": "ADMINISTRADORA COLOMBIANA DE PENSIONES COLPENSIONES"},
    "075828": {"eps": "SALUD TOTAL S.A.",                                   "afp": "PROTECCION"},
    "076575": {"eps": "MUTUAL SER",                                         "afp": "ADMINISTRADORA COLOMBIANA DE PENSIONES COLPENSIONES"},
    "076845": {"eps": "E.P.S SANITAS",                                      "afp": "PROTECCION"},
    "077201": {"eps": "ALIANSALUD EPS",                                     "afp": "PORVENIR"},
    "077659": {"eps": "FAMISANAR",                                          "afp": "PROTECCION"},
    "077911": {"eps": "E.P.S SANITAS",                                      "afp": "PORVENIR"},
    "078082": {"eps": "EPS SURA",                                           "afp": "PORVENIR"},
    "078111": {"eps": "EPS SURA",                                           "afp": "PORVENIR"},
    "081214": {"eps": "COMPENSAR ENTIDAD PROMOTORA DE SALUD",               "afp": "PORVENIR"},
}
MAESTRO_JSON = json.dumps(MAESTRO, ensure_ascii=False)

# ─── Read HTML template ────────────────────────────────────────────────────────
with open("index_template.html", "r", encoding="utf-8") as f:
    html = f.read()

html = html.replace("__LOGO_B64__",     LOGO_B64)
html = html.replace("__MAESTRO_JSON__", MAESTRO_JSON)

with open("index.html", "w", encoding="utf-8") as f:
    f.write(html)

print("index.html generated!")
print(f"  Maestro: {len(MAESTRO)} empleados")
