import os
from sqlalchemy import create_engine, text

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, '..', 'clientes.db')

# Para migrar a PostgreSQL, reemplazar esta línea por:
# engine = create_engine('postgresql+psycopg2://usuario:contraseña@localhost/estudio_rodera')
engine = create_engine(
    f'sqlite:///{os.path.abspath(DB_PATH)}',
    pool_pre_ping=True,
    connect_args={"check_same_thread": False}
)


def init_db():
    with engine.begin() as conn:
        conn.execute(text("""
            CREATE TABLE IF NOT EXISTS users (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT NOT NULL UNIQUE,
                password TEXT NOT NULL,
                fullname TEXT DEFAULT ''
            )
        """))
        conn.execute(text("""
            CREATE TABLE IF NOT EXISTS data_clientes (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                numero_dni TEXT,
                numero_cuil TEXT,
                numero_celular TEXT,
                nombre TEXT,
                apellido TEXT,
                nombre_completo TEXT,
                nombre_completo_2 TEXT,
                sexo TEXT,
                sexo_femenino TEXT,
                sexo_masculino TEXT,
                fecha_de_nacimiento DATE,
                fecha_de_ingreso DATE,
                nacionalidad TEXT,
                direccion TEXT,
                numero_direccion TEXT,
                provincia TEXT,
                departamento TEXT,
                ciudad TEXT,
                formularios_reconocimiento_servicios BOOLEAN DEFAULT 0,
                formularios_jubilacion BOOLEAN DEFAULT 0,
                formularios_jubilacion_con_24_476 BOOLEAN DEFAULT 0,
                formularios_jubilacion_con_27_705 BOOLEAN DEFAULT 0,
                formularios_jubilacion_docente BOOLEAN DEFAULT 0,
                formularios_jubilacion_serv_diferenciales BOOLEAN DEFAULT 0,
                formularios_jubilacion_servicio_domestico BOOLEAN DEFAULT 0,
                formularios_jubilacion_hiv BOOLEAN DEFAULT 0,
                formularios_jubilacion_minusvalidos_ceguera BOOLEAN DEFAULT 0,
                formularios_pension_directa_casados BOOLEAN DEFAULT 0,
                formularios_pension_directa_convivientes BOOLEAN DEFAULT 0,
                formularios_pension_derivada_casados BOOLEAN DEFAULT 0,
                formularios_pension_derivada_convivientes BOOLEAN DEFAULT 0,
                formularios_pension_directa_derivada_hijo_discapacitada BOOLEAN DEFAULT 0,
                formularios_retiro_transitorio_por_invalidez BOOLEAN DEFAULT 0,
                formularios_retiro_transitorio_por_invalidez_sdm BOOLEAN DEFAULT 0,
                formularios_puam BOOLEAN DEFAULT 0,
                formularios_ucap BOOLEAN DEFAULT 0,
                formularios_reajuste_de_haberes BOOLEAN DEFAULT 0,
                formularios_asignacion_fliar_hijo_discapacitado BOOLEAN DEFAULT 0,
                beneficios_nuevo_convenio_de_honorarios_numerado BOOLEAN DEFAULT 0,
                juicios_nuevo_convenio_de_honorarios_numerado BOOLEAN DEFAULT 0,
                convenio_magistrados BOOLEAN DEFAULT 0,
                convenio_de_gastos_administrativos_judiciales BOOLEAN DEFAULT 0,
                _2_91_guarda_documental BOOLEAN DEFAULT 0,
                _6_18_solicitud_prestaciones_previsionales BOOLEAN DEFAULT 0,
                _6_18_solicitud_prestaciones_previsionales_pension BOOLEAN DEFAULT 0,
                acta_poder BOOLEAN DEFAULT 0,
                anexo_baja_puam BOOLEAN DEFAULT 0,
                anexo_i_ley_27625 BOOLEAN DEFAULT 0,
                anexo_ii_dec_894_01 BOOLEAN DEFAULT 0,
                anexo_ii_980_05 BOOLEAN DEFAULT 0,
                anexo_ii_socioeconomico_24_476 BOOLEAN DEFAULT 0,
                baja_pnc BOOLEAN DEFAULT 0,
                carta_poder_srt BOOLEAN DEFAULT 0,
                ddjj_de_salud_resol_300 BOOLEAN DEFAULT 0,
                ddjj_ley_17562_6_9 BOOLEAN DEFAULT 0,
                f_3283_autorizacion_arca BOOLEAN DEFAULT 0,
                formulario_carta_poder_css BOOLEAN DEFAULT 0,
                formulario_encuesta_rti BOOLEAN DEFAULT 0,
                ps_1_75_carta_poder_cap_iii_27705 BOOLEAN DEFAULT 0,
                ps_5_7_derivacion_aportes_obra_social BOOLEAN DEFAULT 0,
                ps_5_11_aceptacion_de_la_obra_social BOOLEAN DEFAULT 0,
                ps_6_292_ddjj_solicitante_sdm BOOLEAN DEFAULT 0,
                ps_6_293_ddjj_dador_de_trabajo_sdm BOOLEAN DEFAULT 0,
                ps_6_294_ddjj_renuncia_sdm BOOLEAN DEFAULT 0,
                ps_6_2_certific_de_servicios BOOLEAN DEFAULT 0,
                ps_6_3_nivel_de_estudios_rti BOOLEAN DEFAULT 0,
                ps_6_4_carta_poder BOOLEAN DEFAULT 0,
                ps_6_8_ddjj_testimonial_acred_servicios BOOLEAN DEFAULT 0,
                ps_6_13_ddjj_testimonial_dependencia_economica BOOLEAN DEFAULT 0,
                ps_6_268_certific_de_servicios_ampliatoria BOOLEAN DEFAULT 0,
                ps_6_273_certific_complementaria_investigadores BOOLEAN DEFAULT 0,
                ps_6_278_dto_de_cuotas_jubilacion BOOLEAN DEFAULT 0,
                ps_6_279_dto_de_cuotas_pension BOOLEAN DEFAULT 0,
                ps_6_284_ddjj_fzas_armadas BOOLEAN DEFAULT 0,
                ps_6_305_carta_poder BOOLEAN DEFAULT 0,
                renuncia_condicionada BOOLEAN DEFAULT 0,
                telegrama_revocando_poder BOOLEAN DEFAULT 0
            )
        """))
