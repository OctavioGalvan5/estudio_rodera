import os
import json
from io import BytesIO

import openai
import openpyxl
from dotenv import load_dotenv
from pypdf import PdfReader, PdfWriter

load_dotenv()

PDF_TEMPLATE = os.path.join(
    os.path.dirname(__file__),
    '..', '..', 'datos', 'Planilla_demanda',
    'Planilla de Sorteo Demanda BLANCO.pdf'
)

# Campos del PDF por posición de interviniente (0-based)
CAMPOS_INTERVINIENTE = [
    {
        'nombre':       'Apellido y nombre Interviniente',
        'doc':          'Doc Interviniente',
        'tipo_doc':     'Listado 6',
        'tipo_parte':   'Listado 4',
        'sexo':         'sexo',
        'nacionalidad': 'Nacionalidad',
        'domicilio':    'domicilio',
        'nro_dom':      'nro_domicilio',
        'localidad':    'localidad',
        'provincia':    'Provincia',
    },
    {
        'nombre':       'Apellido y nombre Interviniente_2',
        'doc':          'Doc Interviniente_2',
        'tipo_doc':     'Listado 6_2',
        'tipo_parte':   'Listado 4_2',
        'sexo':         'Listado 7_2',
        'nacionalidad': 'Campo de texto 1_2',
        'domicilio':    'Campo de texto 3',
        'nro_dom':      'Campo de texto 4_2',
        'localidad':    'Campo de texto 5_2',
        'provincia':    'Listado 2_2',
    },
    {
        'nombre':       'Apellido y nombre Interviniente_3',
        'doc':          'Doc Interviniente_3',
        'tipo_doc':     'Listado 6_3',
        'tipo_parte':   'Listado 4_3',
        'sexo':         'Listado 7_3',
        'nacionalidad': 'Campo de texto 1_3',
        'domicilio':    'Campo de texto 3_2',
        'nro_dom':      'Campo de texto 4_3',
        'localidad':    'Campo de texto 5_3',
        'provincia':    'Listado 2_3',
    },
    {
        'nombre':       'Apellido y nombre Interviniente_4',
        'doc':          'Doc Interviniente_4',
        'tipo_doc':     'Listado 6_4',
        'tipo_parte':   'Listado 4_4',
        'sexo':         'Listado 7_4',
        'nacionalidad': 'Campo de texto 1_4',
        'domicilio':    'Campo de texto 3_3',
        'nro_dom':      'Campo de texto 4_4',
        'localidad':    'Campo de texto 5_4',
        'provincia':    'Listado 2_4',
    },
    {
        'nombre':       'Apellido y nombre Interviniente_5',
        'doc':          'Doc Interviniente_5',
        'tipo_doc':     'Listado 6_5',
        'tipo_parte':   'Listado 4_5',
        'sexo':         'Listado 7_5',
        'nacionalidad': 'Campo de texto 1_5',
        'domicilio':    'Campo de texto 3_4',
        'nro_dom':      'Campo de texto 4_5',
        'localidad':    'Campo de texto 5_5',
        'provincia':    'Listado 2_5',
    },
    {
        'nombre':       'Apellido y nombre Interviniente_18',
        'doc':          'Doc Interviniente_6',
        'tipo_doc':     'Listado 6_6',
        'tipo_parte':   'Listado 4_6',
        'sexo':         'Listado 7_6',
        'nacionalidad': 'Campo de texto 1_6',
        'domicilio':    'Campo de texto 3_5',
        'nro_dom':      'Campo de texto 4_6',
        'localidad':    'Campo de texto 5_6',
        'provincia':    'Listado 2_6',
    },
]

ANSES_DEMANDADO = {
    'nombre':       'ANSeS',
    'doc':          '',
    'tipo_doc':     'CI',
    'tipo_parte':   'DEMANDADO',
    'sexo':         'Masc',
    'nacionalidad': '',
    'domicilio':    'Av. Colon',
    'nro_dom':      '279',
    'localidad':    'Bahia Blanca',
    'provincia':    'BUENOS AIRES',
}

# Columnas cuyo encabezado implica el tipo de documento
_ENCABEZADO_TIPO_DOC = {
    'cuil': 'CUIL',
    'cuit': 'CUIT',
    'dni': 'CI',
    'documento': 'CI',
    'nro doc': 'CI',
    'nro_doc': 'CI',
}

COLUMNAS_ALIAS = {
    'apellido y nombre': 'nombre',
    'apellido y nombres': 'nombre',
    'nombre': 'nombre',
    'nombre completo': 'nombre',
    'tipo doc': 'tipo_doc',
    'tipo_doc': 'tipo_doc',
    'tipo de documento': 'tipo_doc',
    'nro doc': 'doc',
    'nro_doc': 'doc',
    'documento': 'doc',
    'dni': 'doc',
    'cuil': 'doc',
    'cuit': 'doc',
    'tipo parte': 'tipo_parte',
    'tipo_parte': 'tipo_parte',
    'tipo': 'tipo_parte',
    'rol': 'tipo_parte',
    'nacionalidad': 'nacionalidad',
    'domicilio': 'domicilio',
    'localidad': 'localidad',
    'provincia': 'provincia',
    'nro domicilio': 'nro_dom',
    'nro_domicilio': 'nro_dom',
    'numero domicilio': 'nro_dom',
    'sexo': 'sexo',
}


def _normalizar_columna(nombre):
    return nombre.strip().lower()


def leer_excel(file_storage):
    wb = openpyxl.load_workbook(file_storage, data_only=True)
    ws = wb.active

    filas = list(ws.iter_rows(values_only=True))
    if not filas:
        return []

    encabezados_raw = [str(c).strip() if c is not None else '' for c in filas[0]]
    mapa_col = {}
    for idx, enc in enumerate(encabezados_raw):
        clave = COLUMNAS_ALIAS.get(_normalizar_columna(enc))
        if clave:
            mapa_col[clave] = idx

    # Inferir tipo de documento desde el nombre del encabezado
    tipo_doc_inferido = 'CI'
    for enc in encabezados_raw:
        td = _ENCABEZADO_TIPO_DOC.get(_normalizar_columna(enc))
        if td:
            tipo_doc_inferido = td
            break

    intervinientes = []
    for fila in filas[1:]:
        if all(c is None for c in fila):
            continue

        def get(clave):
            idx = mapa_col.get(clave)
            if idx is None:
                return ''
            val = fila[idx]
            return str(val).strip() if val is not None else ''

        nombre = get('nombre')
        if not nombre:
            continue

        intervinientes.append({
            'nombre':       nombre,
            'doc':          get('doc'),
            'tipo_doc':     get('tipo_doc') or tipo_doc_inferido,
            'tipo_parte':   get('tipo_parte') or 'ACTOR',
            'nacionalidad': get('nacionalidad').upper() if get('nacionalidad') else '',
            'domicilio':    get('domicilio'),
            'localidad':    get('localidad'),
            'provincia':    _normalizar_provincia(get('provincia')),
            'nro_dom':      get('nro_dom'),
            'sexo':         get('sexo') or 'Masc',
        })

    return intervinientes


PROVINCIA_MAP = {
    'buenos aires': 'BUENOS AIRES', 'caba': 'CABA', 'capital federal': 'CABA',
    'cordoba': 'CORDOBA', 'santa fe': 'SANTA FE', 'mendoza': 'MENDOZA',
    'tucuman': 'TUCUMAN', 'entre rios': 'ENTRE RIOS', 'salta': 'SALTA',
    'misiones': 'MISIONES', 'chaco': 'CHACO', 'corrientes': 'CORRIENTES',
    'santiago del estero': 'SANTIAGO DEL ESTERO', 'san juan': 'SAN JUAN',
    'jujuy': 'JUJUY', 'rio negro': 'RIO NEGRO', 'neuquen': 'NEUQUEN',
    'la pampa': 'LA PAMPA', 'chubut': 'CHUBUT', 'san luis': 'SAN LUIS',
    'catamarca': 'CATAMARCA', 'la rioja': 'LA RIOJA', 'formosa': 'FORMOSA',
    'santa cruz': 'SANTA CRUZ', 'tierra del fuego': 'TIERRA DEL FUEGO',
}


def _normalizar_provincia(val):
    if not val:
        return 'BUENOS AIRES'
    return PROVINCIA_MAP.get(val.strip().lower(), val.strip().upper())


def _nombre_archivo(nombre):
    import re
    n = re.sub(r'[^a-zA-Z0-9 ,._-]', '', nombre or 'SIN_NOMBRE')
    return n.strip().replace(' ', '_')[:80] + '.pdf'


def generar_planillas_individuales(personas, fecha):
    """Una planilla por persona: la persona como ACTOR y ANSeS como DEMANDADO.
    Devuelve lista de (nombre_archivo, BytesIO)."""
    resultados = []
    vistos = {}
    for persona in personas:
        pdf_buf = _llenar_pdf([persona, ANSES_DEMANDADO], fecha)
        fname = _nombre_archivo(persona['nombre'])
        if fname in vistos:
            vistos[fname] += 1
            fname = fname.replace('.pdf', f'_{vistos[fname]}.pdf')
        else:
            vistos[fname] = 0
        resultados.append((fname, pdf_buf))
    return resultados


def generar_planillas(intervinientes, fecha):
    """Genera uno o más PDFs (máximo 6 intervinientes por planilla)."""
    pdfs = []
    for bloque_inicio in range(0, max(len(intervinientes), 1), 6):
        bloque = intervinientes[bloque_inicio:bloque_inicio + 6]
        pdf_buf = _llenar_pdf(bloque, fecha)
        pdfs.append(pdf_buf)
    return pdfs


def determinar_generos_batch(nombres, batch_size=25):
    """Llama a OpenAI para determinar el género de una lista de nombres.
    Devuelve una lista de 'Masc' o 'Fem' en el mismo orden, garantizando
    que la longitud del resultado siempre coincide con la de la entrada."""
    client = openai.OpenAI(api_key=os.getenv('OPENAI_API_KEY'))
    resultados = []

    for i in range(0, len(nombres), batch_size):
        batch = nombres[i:i + batch_size]
        numerados = '\n'.join(f'{j + 1}. {n}' for j, n in enumerate(batch))

        resp = client.chat.completions.create(
            model='gpt-4o-mini',
            messages=[{
                'role': 'user',
                'content': (
                    'Dado este listado de nombres argentinos en formato "APELLIDO, NOMBRE", '
                    'determiná el género de cada persona basándote en su nombre de pila. '
                    f'Son exactamente {len(batch)} nombres. '
                    'Respondé ÚNICAMENTE con un objeto JSON con la clave "generos" cuyo valor '
                    f'es un array de exactamente {len(batch)} strings "Masc" o "Fem" '
                    'en el mismo orden, sin texto adicional.\n\n'
                    f'Nombres:\n{numerados}'
                ),
            }],
            response_format={'type': 'json_object'},
        )

        data = json.loads(resp.choices[0].message.content)
        generos = data.get('generos', [])

        # Si la IA devolvió menos elementos, completar con 'Masc'
        if len(generos) < len(batch):
            generos = list(generos) + ['Masc'] * (len(batch) - len(generos))

        # Normalizar por si la IA devuelve variantes
        generos = ['Fem' if str(g).strip().lower() in ('fem', 'f', 'femenino') else 'Masc'
                   for g in generos[:len(batch)]]
        resultados.extend(generos)

    return resultados


def extraer_domicilio_batch(domicilios, batch_size=25):
    """Llama a OpenAI para extraer calle limpia y número de una lista de domicilios argentinos.
    Devuelve lista de dicts {"calle": ..., "numero": ...} en el mismo orden.
    Garantiza que la longitud del resultado coincide con la de la entrada."""
    client = openai.OpenAI(api_key=os.getenv('OPENAI_API_KEY'))
    resultados = []

    for i in range(0, len(domicilios), batch_size):
        batch = domicilios[i:i + batch_size]
        numerados = '\n'.join(f'{j + 1}. {d}' for j, d in enumerate(batch))

        resp = client.chat.completions.create(
            model='gpt-4o-mini',
            messages=[{
                'role': 'user',
                'content': (
                    'Dado este listado de domicilios argentinos, extraé la calle (sin número), '
                    'el número de domicilio e indicá si el caso es dudoso. Reglas de extracción:\n'
                    '- "Pellegrini 579" → calle:"Pellegrini", numero:"579", dudoso:false\n'
                    '- "Murature N 1665" → calle:"Murature", numero:"1665", dudoso:false\n'
                    '- "calle 31 n° 1494 piso 4° dto A" → calle:"calle 31", numero:"1494", dudoso:false\n'
                    '- Cuando hay "N", "Nro" o "N°" antes del número, ese es el número de domicilio.\n'
                    '  Ej: "Calle 14 Esq 421 Bis N 645" → calle:"Calle 14 Esq 421 Bis", numero:"645", dudoso:true\n'
                    '- Si hay "(segun DNI) / ... (real)", usá el domicilio marcado como real.\n'
                    '  Ej: "España 78 (segun DNI) / San Martín 542 (real)" → calle:"San Martín", numero:"542", dudoso:true\n'
                    '- "B° Centenario Casa 338" → calle:"B° Centenario", numero:"338", dudoso:false\n'
                    '- "B° Atepam 1 Casa 3" → calle:"B° Atepam", numero:"3", dudoso:true\n'
                    '- Si no hay número → numero:"S/N", dudoso:true\n'
                    'dudoso:true cuando: hay múltiples calles, esquinas, "Bis", conflicto DNI/real, '
                    'calles con nombre numérico, o cualquier ambigüedad. dudoso:false para casos simples.\n'
                    f'Son exactamente {len(batch)} domicilios. '
                    'Respondé ÚNICAMENTE con JSON {"domicilios": [{"calle":..., "numero":..., "dudoso":...}, ...]} '
                    f'con exactamente {len(batch)} objetos, sin texto adicional.\n\n'
                    f'Domicilios:\n{numerados}'
                ),
            }],
            response_format={'type': 'json_object'},
        )

        data = json.loads(resp.choices[0].message.content)
        items = data.get('domicilios', [])

        if len(items) < len(batch):
            items = list(items) + [{'calle': '', 'numero': 'S/N'}] * (len(batch) - len(items))

        for item in items[:len(batch)]:
            resultados.append({
                'calle':  str(item.get('calle', '')).strip(),
                'numero': str(item.get('numero', 'S/N')).strip(),
                'dudoso': bool(item.get('dudoso', False)),
            })

    return resultados


def llenar_pdf_individual(persona, fecha):
    """Genera el PDF de una sola persona (ACTOR) + ANSeS (DEMANDADO)."""
    return _llenar_pdf([persona, ANSES_DEMANDADO], fecha)


def nombre_archivo_pdf(nombre):
    return _nombre_archivo(nombre)


def _llenar_pdf(intervinientes, fecha):
    reader = PdfReader(PDF_TEMPLATE)
    writer = PdfWriter()
    writer.clone_reader_document_root(reader)

    datos = {'fecha': fecha}

    for i, interv in enumerate(intervinientes):
        if i >= len(CAMPOS_INTERVINIENTE):
            break
        campos = CAMPOS_INTERVINIENTE[i]
        datos[campos['nombre']]       = interv['nombre']
        datos[campos['doc']]          = interv['doc']
        datos[campos['tipo_doc']]     = interv['tipo_doc']
        datos[campos['tipo_parte']]   = interv['tipo_parte']
        datos[campos['sexo']]         = interv['sexo']
        datos[campos['nacionalidad']] = interv['nacionalidad']
        datos[campos['domicilio']]    = interv['domicilio']
        datos[campos['nro_dom']]      = interv['nro_dom']
        datos[campos['localidad']]    = interv['localidad']
        datos[campos['provincia']]    = interv['provincia']

    for page in writer.pages:
        writer.update_page_form_field_values(page, datos)

    buf = BytesIO()
    writer.write(buf)
    buf.seek(0)
    return buf
