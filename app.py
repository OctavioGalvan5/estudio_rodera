import os
import json
import time
import uuid
import zipfile
import threading
from io import BytesIO
from datetime import datetime

from flask import (Flask, render_template, request, redirect, url_for,
                   flash, send_file, make_response, jsonify, Response,
                   stream_with_context)
from flask_login import (LoginManager, UserMixin, login_user, logout_user,
                         login_required, current_user)
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.utils import secure_filename
from sqlalchemy import text
from pypdf import PdfReader, PdfWriter
from xhtml2pdf import pisa

from models.database import engine, init_db
from services.consultas.consultas import (
    openai_api_extract_data, update_cliente_in_db,
    process_file, FORMULARIOS_MAPPING
)
from services.planilla_demanda.planilla_demanda_service import (
    leer_excel, determinar_generos_batch, extraer_domicilio_batch,
    llenar_pdf_individual, nombre_archivo_pdf,
)
from services.mails_planilla.mail_service import (
    extraer_datos_pdf, enviar_mail_planilla,
)

# Almacén en memoria para el progreso de cada generación
PROGRESO      = {}  # token -> dict (planillas)
PROGRESO_MAIL = {}  # token -> dict (envíos)

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'rodera-secret-key-2024')

login_manager = LoginManager(app)
login_manager.login_view = 'login'
login_manager.login_message = 'Iniciá sesión para continuar.'
login_manager.login_message_category = 'warning'


@app.template_filter('fecha')
def filtro_fecha(val, fmt='%d/%m/%Y'):
    """Convierte un datetime o string 'YYYY-MM-DD' al formato pedido."""
    if not val:
        return '—'
    if hasattr(val, 'strftime'):
        return val.strftime(fmt)
    try:
        return datetime.strptime(str(val)[:10], '%Y-%m-%d').strftime(fmt)
    except (ValueError, TypeError):
        return str(val)


class User(UserMixin):
    def __init__(self, id, username, fullname=''):
        self.id = id
        self.username = username
        self.fullname = fullname


@login_manager.user_loader
def load_user(user_id):
    with engine.connect() as conn:
        row = conn.execute(
            text("SELECT id, username, fullname FROM users WHERE id = :id"),
            {"id": int(user_id)}
        ).fetchone()
    if row:
        return User(row[0], row[1], row[2] or '')
    return None


# ---------------------------------------------------------------------------
# AUTH
# ---------------------------------------------------------------------------

@app.route('/')
@login_required
def index():
    return redirect(url_for('home'))


@app.route('/home')
@login_required
def home():
    return render_template('home.html')


@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = request.form.get('username', '').strip()
        password = request.form.get('password', '')
        with engine.connect() as conn:
            row = conn.execute(
                text("SELECT id, username, password, fullname FROM users WHERE username = :u"),
                {"u": username}
            ).fetchone()
        if row and check_password_hash(row[2], password):
            login_user(User(row[0], row[1], row[3] or ''))
            return redirect(url_for('home'))
        flash('Usuario o contraseña incorrectos.', 'danger')
    return render_template('auth/login.html')


@app.route('/logout')
@login_required
def logout():
    logout_user()
    return redirect(url_for('login'))


# ---------------------------------------------------------------------------
# CONSULTAS
# ---------------------------------------------------------------------------

@app.route('/consultas')
@login_required
def consultas():
    with engine.connect() as conn:
        result = conn.execute(text("SELECT * FROM data_clientes ORDER BY id DESC"))
        data_clientes = [dict(row._mapping) for row in result]
    return render_template('consultas/consultas.html', data_clientes=data_clientes)


@app.route('/upload_dni', methods=['POST'])
@login_required
def upload_dni():
    documentos = request.files.getlist('documentos')
    if len(documentos) < 1:
        flash("Debe enviar al menos un archivo del DNI.", "danger")
        return redirect(url_for('consultas'))

    processed_files = []
    for file in documentos:
        pages = process_file(file)
        if not pages:
            flash("Error al procesar alguno de los archivos.", "danger")
            return redirect(url_for('consultas'))
        processed_files.extend(pages)

    extracted_data, error = openai_api_extract_data(processed_files)
    if error or not extracted_data:
        flash(f"Error al extraer datos del DNI: {error}", "danger")
        return redirect(url_for('consultas'))

    if not extracted_data.get('date_of_birth'):
        extracted_data['date_of_birth'] = None
    if not extracted_data.get('entry_date'):
        extracted_data['entry_date'] = None

    try:
        with engine.begin() as conn:
            result = conn.execute(text("""
                INSERT INTO data_clientes (
                    numero_dni, numero_cuil, numero_celular, nombre, apellido,
                    nombre_completo, nombre_completo_2, sexo, sexo_femenino,
                    sexo_masculino, fecha_de_nacimiento, fecha_de_ingreso,
                    nacionalidad, direccion, numero_direccion, provincia,
                    departamento, ciudad
                ) VALUES (
                    :dni_number, :cuil_number, :phone_number, :name, :surname,
                    :full_name, :full_name_2, :sexo, :sexo_femenino,
                    :sexo_masculino, :date_of_birth, :entry_date,
                    :nationality, :address, :adress_number, :province,
                    :department, :city
                )
            """), extracted_data)
            new_id = result.lastrowid
    except Exception as e:
        flash(f"Error al guardar los datos: {e}", "danger")
        return redirect(url_for('consultas'))

    return redirect(url_for('ver_cliente', id=new_id))


@app.route('/agregar_cliente', methods=['POST'])
@login_required
def agregar_cliente():
    try:
        with engine.begin() as conn:
            result = conn.execute(text("""
                INSERT INTO data_clientes (
                    numero_dni, numero_cuil, numero_celular, nombre, apellido,
                    nombre_completo, nombre_completo_2, sexo, sexo_femenino,
                    sexo_masculino, fecha_de_nacimiento, fecha_de_ingreso,
                    nacionalidad, direccion, numero_direccion, provincia,
                    departamento, ciudad
                ) VALUES (
                    NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL,
                    NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL
                )
            """))
            new_id = result.lastrowid
    except Exception:
        return redirect(url_for('consultas'))
    return redirect(url_for('ver_cliente', id=new_id))


@app.route('/eliminar_cliente/<int:id>', methods=['POST'])
@login_required
def eliminar_cliente(id):
    try:
        with engine.begin() as conn:
            conn.execute(text("DELETE FROM data_clientes WHERE id = :id"), {"id": id})
    except Exception as e:
        print(f"Error al eliminar: {e}")
    return redirect(url_for('consultas'))


@app.route('/ver_cliente/<int:id>', methods=['GET', 'POST'])
@login_required
def ver_cliente(id):
    if request.method == 'POST':
        accion = request.form.get('accion')
        data = request.form.to_dict()
        update_cliente_in_db(data)

        if accion == 'hacer_formulario':
            with engine.connect() as conn:
                row = conn.execute(
                    text("SELECT * FROM data_clientes WHERE id = :id"), {"id": id}
                ).mappings().first()

            nombre = row.get("nombre", "") or ""
            apellido = row.get("apellido", "") or ""

            def fmt(d, f):
                if not d:
                    return ""
                if hasattr(d, 'strftime'):
                    return d.strftime(f)
                try:
                    return datetime.strptime(str(d)[:10], '%Y-%m-%d').strftime(f)
                except (ValueError, TypeError):
                    return str(d)

            datos = {
                "nombre": nombre,
                "apellido": apellido,
                "numero_celular": row.get("numero_celular", "") or "",
                "nombre_completo": f"{apellido} {nombre}",
                "nombre_completo_2": f"{nombre} {apellido}",
                "sexo": row.get("sexo", "") or "",
                "sexo_femenino": row.get("sexo_femenino", "") or "",
                "sexo_masculino": row.get("sexo_masculino", "") or "",
                "numero_dni": row.get("numero_dni", "") or "",
                "fecha_de_nacimiento_formato": fmt(row.get("fecha_de_nacimiento"), '%d/%m/%Y'),
                "fecha_de_nacimiento": fmt(row.get("fecha_de_nacimiento"), '%d%m%Y'),
                "fecha_de_nacimiento_dia": fmt(row.get("fecha_de_nacimiento"), '%d'),
                "fecha_de_nacimiento_mes": fmt(row.get("fecha_de_nacimiento"), '%m'),
                "fecha_de_nacimiento_año": fmt(row.get("fecha_de_nacimiento"), '%Y'),
                "fecha_de_ingreso": fmt(row.get("fecha_de_ingreso"), '%d%m%y'),
                "numero_cuil": row.get("numero_cuil", "") or "",
                "cuil_inicio": (row.get("numero_cuil", "") or "")[:2],
                "cuil_fin": (row.get("numero_cuil", "") or "")[-1:],
                "nacionalidad": row.get("nacionalidad", "") or "",
                "direccion": row.get("direccion", "") or "",
                "numero_direccion": row.get("numero_direccion", "") or "",
                "provincia": row.get("provincia", "") or "",
                "departamento": row.get("departamento", "") or "",
                "ciudad": row.get("ciudad", "") or "",
                "donde_firmar": "X",
            }

            archivos_generados = {}
            lista_formularios = []
            formularios_pdf = []
            formularios_docx = []

            for checkbox_name, info in FORMULARIOS_MAPPING.items():
                if request.form.get(checkbox_name):
                    lista_formularios.append(info["label"])
                    if info["path"].endswith(".docx"):
                        formularios_docx.append(info["path"])
                    else:
                        formularios_pdf.append(info["path"])

            sabana_writer = PdfWriter()

            for formulario in formularios_pdf:
                if not os.path.exists(formulario):
                    continue
                reader = PdfReader(formulario)
                writer = PdfWriter()
                writer.clone_reader_document_root(reader)
                if len(writer.pages) == 0:
                    continue
                writer.update_page_form_field_values(writer.pages[0], datos)
                nombre_formulario = secure_filename(os.path.splitext(os.path.basename(formulario))[0])
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                output = BytesIO()
                writer.write(output)
                output.seek(0)
                archivos_generados[f"{nombre_formulario}_{timestamp}.pdf"] = output

                output.seek(0)
                for page in PdfReader(output).pages:
                    sabana_writer.add_page(page)

            if formularios_docx:
                from docxtpl import DocxTemplate
                for template_path in formularios_docx:
                    if not os.path.exists(template_path):
                        continue
                    doc = DocxTemplate(template_path)
                    doc.render(datos)
                    output_word = BytesIO()
                    doc.save(output_word)
                    output_word.seek(0)
                    nombre_base = os.path.splitext(os.path.basename(template_path))[0]
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    archivos_generados[f"{nombre_base}_{timestamp}.docx"] = output_word

            rendered = render_template('consultas/formularios_impresos.html', filas=lista_formularios)
            pdf_lista_buffer = BytesIO()
            pisa.CreatePDF(rendered, dest=pdf_lista_buffer)
            pdf_lista_buffer.seek(0)
            archivos_generados["lista_formularios.pdf"] = pdf_lista_buffer

            if sabana_writer.pages:
                sabana_buffer = BytesIO()
                sabana_writer.write(sabana_buffer)
                sabana_buffer.seek(0)
                archivos_generados["sabana_formularios.pdf"] = sabana_buffer

            zip_buffer = BytesIO()
            with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
                for nombre_archivo, contenido in archivos_generados.items():
                    contenido.seek(0)
                    zf.writestr(nombre_archivo, contenido.read())
            zip_buffer.seek(0)

            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            nombre_zip = f"formularios_{apellido}_{nombre}_{timestamp}.zip".replace(" ", "_")

            response = make_response(send_file(
                zip_buffer,
                mimetype='application/zip',
                as_attachment=True,
                download_name=nombre_zip
            ))
            response.set_cookie('fileDownloadReady', '1')
            return response

        flash("Cambios guardados correctamente.", "success")
        return redirect(url_for('ver_cliente', id=id))

    with engine.connect() as conn:
        row = conn.execute(
            text("SELECT * FROM data_clientes WHERE id = :id"), {"id": id}
        ).mappings().first()

    if not row:
        flash("Cliente no encontrado.", "danger")
        return redirect(url_for('consultas'))

    data = dict(row)
    for campo in ('fecha_de_nacimiento', 'fecha_de_ingreso'):
        val = data.get(campo)
        if val is None:
            data[campo] = ''
        elif hasattr(val, 'strftime'):
            data[campo] = val.strftime('%Y-%m-%d')
        else:
            data[campo] = str(val)[:10]

    return render_template('consultas/ver_cliente.html', data_cliente=data)


# ---------------------------------------------------------------------------
# PLANILLA DEMANDA
# ---------------------------------------------------------------------------

@app.route('/planilla_demanda')
@login_required
def planilla_demanda():
    hoy = datetime.now().strftime('%Y-%m-%d')
    return render_template('planilla_demanda/planilla_demanda.html', hoy=hoy)


@app.route('/planilla_demanda/iniciar', methods=['POST'])
@login_required
def planilla_demanda_iniciar():
    fecha = request.form.get('fecha', '').strip()
    excel_file = request.files.get('excel')

    if not fecha or not excel_file or not excel_file.filename:
        return jsonify({'error': 'Faltan datos'}), 400

    try:
        personas = leer_excel(excel_file)
    except Exception as e:
        return jsonify({'error': f'Error al leer el Excel: {e}'}), 400

    if not personas:
        return jsonify({'error': 'El archivo Excel no tiene datos válidos.'}), 400

    token = str(uuid.uuid4())
    PROGRESO[token] = {
        'etapa': 'inicio',
        'porcentaje': 0,
        'mensaje': 'Iniciando...',
        'nombre_actual': '',
        'total': len(personas),
        'listo': False,
        'error': None,
        'fecha': fecha,
    }

    def procesar():
        try:
            total = len(personas)
            batch_size = 25

            idx_sin_nro = [i for i, p in enumerate(personas)
                           if not p.get('nro_dom') and p.get('domicilio')]
            hay_dom = bool(idx_sin_nro)
            domicilios_originales = {i: personas[i]['domicilio'] for i in idx_sin_nro}
            genero_max_pct = 50 if hay_dom else 100

            # ── Fase 1a: género ──────────────────────────────────────────────
            PROGRESO[token].update({'etapa': 'genero', 'mensaje': f'Extrayendo géneros... (0/{total})'})
            generos = []
            for i in range(0, total, batch_size):
                batch_nombres = [p['nombre'] for p in personas[i:i + batch_size]]
                generos_batch = determinar_generos_batch(batch_nombres, batch_size=batch_size)
                generos.extend(generos_batch)
                procesados = len(generos)
                PROGRESO[token].update({
                    'porcentaje': int(procesados / total * genero_max_pct),
                    'mensaje': f'Extrayendo géneros... ({procesados}/{total})',
                })
            for persona, genero in zip(personas, generos):
                persona['sexo'] = genero

            # ── Fase 1b: domicilio → pausa para revisión ────────────────────
            if hay_dom:
                total_dom = len(idx_sin_nro)
                domicilios_a_procesar = [personas[i]['domicilio'] for i in idx_sin_nro]
                PROGRESO[token].update({'etapa': 'domicilio',
                                        'mensaje': f'Extrayendo domicilios... (0/{total_dom})'})
                dom_resultado = []
                for i in range(0, total_dom, batch_size):
                    batch_dom = domicilios_a_procesar[i:i + batch_size]
                    dom_resultado.extend(extraer_domicilio_batch(batch_dom, batch_size=batch_size))
                    PROGRESO[token].update({
                        'porcentaje': 50 + int(len(dom_resultado) / total_dom * 50),
                        'mensaje': f'Extrayendo domicilios... ({len(dom_resultado)}/{total_dom})',
                    })

                personas_revision = []
                for idx_p, dom in zip(idx_sin_nro, dom_resultado):
                    personas[idx_p]['nro_dom'] = dom['numero']
                    if dom['calle']:
                        personas[idx_p]['domicilio'] = dom['calle']
                    personas_revision.append({
                        'idx':                idx_p,
                        'nombre':             personas[idx_p]['nombre'],
                        'domicilio_original': domicilios_originales[idx_p],
                        'calle':              dom['calle'],
                        'nro_dom':            dom['numero'],
                        'dudoso':             dom.get('dudoso', False),
                        'sexo':               personas[idx_p].get('sexo', 'Masc'),
                    })

                PROGRESO[token].update({
                    'etapa':             'revision',
                    'porcentaje':        100,
                    'mensaje':           'Revisá los domicilios antes de generar.',
                    'personas':          personas,
                    'personas_revision': personas_revision,
                })

            else:
                # Sin extracción de domicilio → generar PDFs directamente
                PROGRESO[token].update({'etapa': 'pdf', 'mensaje': 'Generando PDFs...'})
                planillas, vistos = [], {}
                for idx, persona in enumerate(personas, 1):
                    PROGRESO[token].update({
                        'porcentaje': int(idx / total * 95),
                        'mensaje': f'Generando PDF {idx}/{total}',
                        'nombre_actual': persona['nombre'],
                    })
                    pdf_buf = llenar_pdf_individual(persona, fecha)
                    fname = nombre_archivo_pdf(persona['nombre'])
                    if fname in vistos:
                        vistos[fname] += 1
                        fname = fname.replace('.pdf', f'_{vistos[fname]}.pdf')
                    else:
                        vistos[fname] = 0
                    planillas.append((fname, pdf_buf))

                PROGRESO[token].update({'etapa': 'zip', 'porcentaje': 95,
                                        'mensaje': 'Comprimiendo...', 'nombre_actual': ''})
                zip_buf = BytesIO()
                with zipfile.ZipFile(zip_buf, 'w', zipfile.ZIP_DEFLATED) as zf:
                    for fname, pdf_buf in planillas:
                        zf.writestr(fname, pdf_buf.read())
                zip_buf.seek(0)
                PROGRESO[token].update({
                    'etapa': 'listo', 'porcentaje': 100,
                    'mensaje': f'¡Listo! {total} planillas generadas.',
                    'zip': zip_buf, 'listo': True,
                })

        except Exception as e:
            PROGRESO[token].update({'error': str(e), 'listo': True})

    threading.Thread(target=procesar, daemon=True).start()
    return jsonify({'token': token, 'total': len(personas)})


@app.route('/planilla_demanda/confirmar/<token>', methods=['POST'])
@login_required
def planilla_demanda_confirmar(token):
    info = PROGRESO.get(token)
    if not info or info.get('etapa') != 'revision':
        return jsonify({'error': 'Token inválido o estado incorrecto'}), 400

    edits = (request.get_json() or {}).get('items', [])
    personas = info['personas']
    for edit in edits:
        idx = int(edit['idx'])
        personas[idx]['domicilio'] = edit.get('calle',   personas[idx]['domicilio'])
        personas[idx]['nro_dom']   = edit.get('nro_dom', personas[idx]['nro_dom'])
        personas[idx]['sexo']      = edit.get('sexo',    personas[idx].get('sexo', 'Masc'))

    fecha = info['fecha']
    total = len(personas)
    PROGRESO[token].update({
        'etapa': 'pdf', 'porcentaje': 0,
        'mensaje': 'Generando PDFs...', 'nombre_actual': '',
        'listo': False, 'error': None, 'personas_revision': [],
    })

    def generar():
        try:
            planillas, vistos = [], {}
            for idx, persona in enumerate(personas, 1):
                PROGRESO[token].update({
                    'porcentaje': int(idx / total * 95),
                    'mensaje': f'Generando PDF {idx}/{total}',
                    'nombre_actual': persona['nombre'],
                })
                pdf_buf = llenar_pdf_individual(persona, fecha)
                fname = nombre_archivo_pdf(persona['nombre'])
                if fname in vistos:
                    vistos[fname] += 1
                    fname = fname.replace('.pdf', f'_{vistos[fname]}.pdf')
                else:
                    vistos[fname] = 0
                planillas.append((fname, pdf_buf))

            PROGRESO[token].update({'etapa': 'zip', 'porcentaje': 95,
                                    'mensaje': 'Comprimiendo...', 'nombre_actual': ''})
            zip_buf = BytesIO()
            with zipfile.ZipFile(zip_buf, 'w', zipfile.ZIP_DEFLATED) as zf:
                for fname, pdf_buf in planillas:
                    zf.writestr(fname, pdf_buf.read())
            zip_buf.seek(0)
            PROGRESO[token].update({
                'etapa': 'listo', 'porcentaje': 100,
                'mensaje': f'¡Listo! {total} planillas generadas.',
                'zip': zip_buf, 'listo': True,
            })
        except Exception as e:
            PROGRESO[token].update({'error': str(e), 'listo': True})

    threading.Thread(target=generar, daemon=True).start()
    return jsonify({'ok': True})


@app.route('/planilla_demanda/progreso/<token>')
@login_required
def planilla_demanda_progreso(token):
    def stream():
        while True:
            info = PROGRESO.get(token)
            if not info:
                yield f"data: {json.dumps({'error': 'Token inválido'})}\n\n"
                break
            excluir = {'zip', 'personas'}
            payload = {k: v for k, v in info.items() if k not in excluir}
            yield f"data: {json.dumps(payload)}\n\n"
            if info.get('listo') or info.get('etapa') == 'revision':
                break
            time.sleep(0.4)

    return Response(
        stream_with_context(stream()),
        mimetype='text/event-stream',
        headers={'Cache-Control': 'no-cache', 'X-Accel-Buffering': 'no'},
    )


@app.route('/planilla_demanda/descargar/<token>')
@login_required
def planilla_demanda_descargar(token):
    info = PROGRESO.pop(token, None)
    if not info or not info.get('zip'):
        flash('No hay archivo disponible para descargar.', 'danger')
        return redirect(url_for('planilla_demanda'))
    zip_buf = info['zip']
    zip_buf.seek(0)
    fecha = info.get('fecha', 'planillas')
    return send_file(
        zip_buf,
        mimetype='application/zip',
        as_attachment=True,
        download_name=f'planillas_demanda_{fecha}.zip',
    )


# ---------------------------------------------------------------------------
# MAILS PLANILLA
# ---------------------------------------------------------------------------

@app.route('/mails_planilla')
@login_required
def mails_planilla():
    return render_template('mails_planilla/mails_planilla.html')


@app.route('/mails_planilla/cargar', methods=['POST'])
@login_required
def mails_planilla_cargar():
    archivos = request.files.getlist('pdfs')
    if not archivos:
        return jsonify({'error': 'No se recibieron archivos.'}), 400

    token = str(uuid.uuid4())
    items_meta = []
    items_store = []

    for i, archivo in enumerate(archivos):
        pdf_bytes = archivo.read()
        nombre, cuil = extraer_datos_pdf(pdf_bytes)
        items_meta.append({'idx': i, 'nombre': nombre, 'cuil': cuil, 'filename': archivo.filename})
        items_store.append({
            'nombre':   nombre,
            'cuil':     cuil,
            'filename': archivo.filename,
            'bytes':    pdf_bytes,
            'estado':   'pendiente',
            'error':    None,
        })

    PROGRESO_MAIL[token] = {
        'pdfs':         items_store,
        'etapa':        'cargado',
        'porcentaje':   0,
        'mensaje':      '',
        'nombre_actual':'',
        'enviados':     0,
        'errores':      0,
        'total':        len(items_store),
        'listo':        False,
        'error':        None,
    }
    return jsonify({'token': token, 'items': items_meta})


@app.route('/mails_planilla/enviar/iniciar', methods=['POST'])
@login_required
def mails_planilla_enviar_iniciar():
    data  = request.get_json() or {}
    token = data.get('token')
    info  = PROGRESO_MAIL.get(token)
    if not info:
        return jsonify({'error': 'Token inválido.'}), 400

    for edit in data.get('edits', []):
        idx = int(edit['idx'])
        info['pdfs'][idx]['nombre'] = edit.get('nombre', info['pdfs'][idx]['nombre'])
        info['pdfs'][idx]['cuil']   = edit.get('cuil',   info['pdfs'][idx]['cuil'])

    destinatario = (data.get('destinatario') or '').strip()
    total = info['total']
    info.update({'etapa': 'enviando', 'porcentaje': 0, 'listo': False, 'enviados': 0, 'errores': 0})

    def enviar():
        for i, item in enumerate(info['pdfs'], 1):
            info.update({
                'porcentaje':    int(i / total * 100),
                'mensaje':       f'Enviando {i}/{total}: {item["nombre"]}',
                'nombre_actual': item['nombre'],
            })
            try:
                enviar_mail_planilla(item['nombre'], item['cuil'], item['bytes'], item['filename'],
                                     destinatario=destinatario)
                item['estado'] = 'enviado'
                info['enviados'] += 1
            except Exception as e:
                item['estado'] = 'error'
                item['error']  = str(e)
                info['errores'] += 1

        info.update({
            'etapa':         'listo',
            'porcentaje':    100,
            'mensaje':       f'{info["enviados"]} enviados, {info["errores"]} errores.',
            'nombre_actual': '',
            'listo':         True,
        })

    threading.Thread(target=enviar, daemon=True).start()
    return jsonify({'ok': True})


@app.route('/mails_planilla/progreso/<token>')
@login_required
def mails_planilla_progreso(token):
    def stream():
        while True:
            info = PROGRESO_MAIL.get(token)
            if not info:
                yield f"data: {json.dumps({'error': 'Token inválido'})}\n\n"
                break
            payload = {k: v for k, v in info.items() if k != 'pdfs'}
            payload['items_estado'] = [
                {'nombre': it['nombre'], 'filename': it['filename'],
                 'estado': it['estado'], 'error': it.get('error')}
                for it in info['pdfs']
            ]
            yield f"data: {json.dumps(payload)}\n\n"
            if info.get('listo'):
                break
            time.sleep(0.5)

    return Response(
        stream_with_context(stream()),
        mimetype='text/event-stream',
        headers={'Cache-Control': 'no-cache', 'X-Accel-Buffering': 'no'},
    )


# ---------------------------------------------------------------------------
# INICIO
# ---------------------------------------------------------------------------

if __name__ == '__main__':
    init_db()
    app.run(debug=True, port=5001)
