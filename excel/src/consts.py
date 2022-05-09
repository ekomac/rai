from openpyxl.styles.borders import Border, Side
from openpyxl.styles import Alignment, PatternFill
CATEGORIES = [
    'Educación-Rendimiento',
    'Datos Básicos-General',
    'Familia-Estructura Familiar',
    'Educación-General',
    'Educación-Maltrato Escolar',
    'Educación-Educación Tecnológica',
    'Crianza-Cuidados',
    'Crianza-Salud',
    'Trabajo Infantil-General',
    'Tiempo Libre-General',
    'Ámbitos protectores-Pares',
    'Ámbitos protectores-Comunidad',
    'Salud-Mental',
    'Salud-Sexual',
    'Sospechas-General',
    'Salud-General',
    'Salud-Física',
    'Salud-Nutrición',
    'Familia-Vivienda',
    'Familia-Contexto Familiar',
    'Familia-Madre o Padre',
    'Familia-Salud Familiar',
    'Familia-Capacidades',
    'Familia-Economía',
    'Familia-Redes y Ayudas'
]

COLUMNS_TO_MERGE = [
    "A1:A2",
    "B1:F1",
    "G1:G2",
    "H1:H2",
    "I1:I2",
    "J1:J2",
    "K1:K2",
    "L1:L2",
]

COLS_REF = {
    # A
    "CATEGORIA": {'title': (1, 1, "CATEGORIA", ), 'col': 1},
    # B
    "PREGUNTA_TOP": {'title': (1, 2, "PREGUNTA", ), 'col': 2},
    # B
    "ID": {'title': (2, 2, "ID", ), 'col': 2},
    # C
    "PREGUNTA": {'title': (2, 3, "PREGUNTA", ), 'col': 3},
    # D
    "TIPO": {'title': (2, 4, "TIPO", ), 'col': 4},
    # E
    "SUBTIPO": {'title': (2, 5, "SIGNO", ), 'col': 5},
    # F
    "OP": {'title': (1, 6, "OP", ), 'col': 6},
    # G
    "RTA": {'title': (1, 7, "RTA", ), 'col': 7},
    # H
    "VALOR": {'title': (1, 8, "VALOR", ), 'col': 8},
    # I
    "CON_OP": {'title': (1, 9, "CON OP", ), 'col': 9},
    # J
    "SUMA_TOTAL": {'title': (1, 10, "MAX", ), 'col': 10},
    # K
    "PREGUNTAS_TOTALES": {'title': (1, 11, "PREGUNTAS TOTALES", ), 'col': 11},
    # L
    "VALOR_FINAL": {'title': (1, 12, "VALOR FINAL", ), 'col': 12},
}

TITLES = [
    (1, 1, "CATEGORIA"),            # A
    (1, 2, "PREGUNTA"),             # B
    (2, 2, "ID"),                   # B
    (2, 3, "PREGUNTA"),             # C
    (2, 4, "TIPO"),                 # D
    (2, 5, "SIGNO"),              # E
    (2, 6, "OP"),                   # F
    (1, 7, "RTA"),                  # G
    (1, 8, "VALOR"),                # H
    (1, 9, "CON OP"),               # I
    (1, 10, "SUMA TOTAL (PEOR CASO)"),          # J
    (1, 11, "PREGUNTAS TOTALES"),   # K
    (1, 12, "VALOR FINAL"),         # L
]

TYPE_BOL = "BOL"
TYPE_VEM = "VEM"
TYPE_VEX = "VEX"
TYPE_VEC = "VEC"
TYPE_MULT = "MULT"
TYPE_SING = "SING"
TYPE_UNDEFINED = "undefined"

UNDEFINED_FILL = PatternFill(
    start_color="FFFF00", end_color="FFFF00", fill_type="solid")

THIN_BORDER = Border(left=Side(style='thin'),
                     right=Side(style='thin'),
                     top=Side(style='thin'),
                     bottom=Side(style='thin'))

ALIGNMENT = Alignment(horizontal='general',
                      vertical='top',
                      text_rotation=0,
                      wrap_text=False,
                      shrink_to_fit=False,
                      indent=0)

WRAPPED_ALIGNMENT = Alignment(horizontal='general',
                              vertical='top',
                              text_rotation=0,
                              wrap_text=True,
                              shrink_to_fit=False,
                              indent=0)
