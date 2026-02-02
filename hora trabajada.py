import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
from datetime import datetime, time, date, timedelta, timezone
import calendar
import os
import json
import gspread
from google.oauth2.service_account import Credentials

# Zona horaria de Colombia (UTC-5)
COLOMBIA_UTC_OFFSET = timedelta(hours=-5)
COLOMBIA_TZ = timezone(COLOMBIA_UTC_OFFSET)

def obtener_hora_colombia():
    """Obtener la hora actual en zona horaria de Colombia (UTC-5)"""
    return datetime.now(COLOMBIA_TZ)

def obtener_hora_colombia_time():
    """Obtener solo el objeto time en zona horaria de Colombia"""
    return obtener_hora_colombia().time()

def obtener_fecha_colombia():
    """Obtener la fecha actual en zona horaria de Colombia"""
    return obtener_hora_colombia().date()

# Configuración de la página
st.set_page_config(
    page_title="ChronoTrack - Tekpro",
    page_icon="⏱",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# CSS personalizado - Tema Tekpro
st.markdown("""
<style>
    /* Importar fuente moderna */
    @import url('https://fonts.googleapis.com/css2?family=Poppins:wght@300;400;500;600;700&display=swap');
    
    /* Variables de colores Tekpro */
    :root {
        --tekpro-primary: #3EAEA5;
        --tekpro-secondary: #5BC4BC;
        --tekpro-light: #7DD4CE;
        --tekpro-dark: #2D8B84;
        --tekpro-darker: #1E5E59;
        --text-dark: #2C3E50;
        --text-light: #FFFFFF;
        --bg-light: #F8FFFE;
        --shadow: rgba(62, 174, 165, 0.15);
    }
    
    /* Fondo general con patrón pixelado sutil */
    .stApp {
        background: linear-gradient(135deg, #F8FFFE 0%, #E8F8F7 100%);
        font-family: 'Poppins', sans-serif;
    }
    
    /* Ocultar elementos de Streamlit */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    
    /* Contenedor principal */
    .main .block-container {
        padding-top: 2rem;
        padding-bottom: 2rem;
        max-width: 1200px;
    }
    
    /* Títulos principales */
    h1, h2, h3 {
        font-family: 'Poppins', sans-serif;
        font-weight: 600;
        color: var(--tekpro-dark);
    }
    
    /* Logo y header estilo Tekpro */
    .tekpro-header {
        background: linear-gradient(135deg, var(--tekpro-primary) 0%, var(--tekpro-secondary) 100%);
        padding: 25px 40px;
        border-radius: 20px;
        box-shadow: 0 10px 30px var(--shadow);
        margin-bottom: 30px;
        position: relative;
        overflow: hidden;
    }
    
    .tekpro-header::before {
        content: '';
        position: absolute;
        top: -50px;
        right: -50px;
        width: 200px;
        height: 200px;
        background: rgba(255, 255, 255, 0.1);
        border-radius: 20px;
        transform: rotate(45deg);
    }
    
    .tekpro-logo {
        font-size: 32px;
        font-weight: 700;
        color: white;
        text-align: center;
        margin: 0;
        text-shadow: 2px 2px 4px rgba(0,0,0,0.1);
    }
    
    .tekpro-subtitle {
        font-size: 18px;
        color: rgba(255, 255, 255, 0.9);
        text-align: center;
        margin-top: 5px;
        font-weight: 300;
    }
    
    /* Contenedor de código de barras */
    .codigo-barras-container {
        background: white;
        padding: 25px;
        border-radius: 15px;
        border-left: 5px solid var(--tekpro-primary);
        box-shadow: 0 5px 20px var(--shadow);
        margin-bottom: 25px;
        transition: transform 0.3s ease;
    }
    
    .codigo-barras-container:hover {
        transform: translateY(-2px);
        box-shadow: 0 8px 25px var(--shadow);
    }
    
    /* Mensajes de estado */
    .success-message {
        background: linear-gradient(135deg, #D4F4F2 0%, #E8FAF9 100%);
        color: var(--tekpro-darker);
        padding: 20px;
        border-radius: 12px;
        border-left: 5px solid var(--tekpro-primary);
        margin: 15px 0;
        box-shadow: 0 4px 15px var(--shadow);
        animation: slideIn 0.4s ease;
    }
    
    .error-message {
        background: linear-gradient(135deg, #FFE5E5 0%, #FFF0F0 100%);
        color: #C62828;
        padding: 20px;
        border-radius: 12px;
        border-left: 5px solid #EF5350;
        margin: 15px 0;
        box-shadow: 0 4px 15px rgba(239, 83, 80, 0.15);
        animation: slideIn 0.4s ease;
    }
    
    .info-message {
        background: linear-gradient(135deg, #E3F5F4 0%, #F0FFFE 100%);
        color: var(--tekpro-dark);
        padding: 20px;
        border-radius: 12px;
        border-left: 5px solid var(--tekpro-secondary);
        margin: 15px 0;
        box-shadow: 0 4px 15px var(--shadow);
        animation: slideIn 0.4s ease;
    }
    
    /* Botones estilo Tekpro */
    .stButton > button {
        width: 100%;
        background: linear-gradient(135deg, var(--tekpro-primary) 0%, var(--tekpro-secondary) 100%);
        color: white;
        border: none;
        padding: 15px 30px;
        border-radius: 12px;
        font-size: 16px;
        font-weight: 600;
        font-family: 'Poppins', sans-serif;
        box-shadow: 0 5px 20px var(--shadow);
        transition: all 0.3s ease;
        letter-spacing: 0.5px;
    }
    
    .stButton > button:hover {
        background: linear-gradient(135deg, var(--tekpro-dark) 0%, var(--tekpro-primary) 100%);
        box-shadow: 0 8px 30px var(--shadow);
        transform: translateY(-2px);
    }
    
    .stButton > button:active {
        transform: translateY(0);
        box-shadow: 0 3px 15px var(--shadow);
    }
    
    /* Botón con ícono */
    .stButton > button[kind="primary"]::before {
        content: '';
        display: inline-block;
        width: 20px;
        height: 20px;
        margin-right: 10px;
        background-image: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' fill='white'%3E%3Cpath d='M12 2L15 8.5L22 9.5L17 14.5L18 21.5L12 18L6 21.5L7 14.5L2 9.5L9 8.5L12 2Z'/%3E%3C/svg%3E");
        background-size: contain;
        background-repeat: no-repeat;
        vertical-align: middle;
    }
    
    /* Campo de código de barras */
    .barcode-scanner-field input {
        font-size: 24px !important;
        font-family: 'Courier New', monospace !important;
        text-align: center !important;
        padding: 20px !important;
        border: 3px solid var(--tekpro-primary) !important;
        border-radius: 12px !important;
        background-color: white !important;
        box-shadow: 0 5px 20px var(--shadow) !important;
        transition: all 0.3s ease !important;
        font-weight: 600 !important;
    }
    
    .barcode-scanner-field input:focus {
        border-color: var(--tekpro-secondary) !important;
        box-shadow: 0 0 25px rgba(62, 174, 165, 0.3) !important;
        transform: scale(1.02) !important;
        outline: none !important;
    }
    
    /* Instrucciones del escáner */
    .scanner-instructions {
        background: linear-gradient(135deg, var(--tekpro-primary) 0%, var(--tekpro-secondary) 100%);
        color: white;
        padding: 20px;
        border-radius: 12px;
        text-align: center;
        margin: 15px 0;
        animation: pulse 2s infinite;
        box-shadow: 0 5px 20px var(--shadow);
    }
    
    .scanner-instructions h4 {
        color: white;
        margin: 0;
        font-weight: 600;
    }
    
    .scanner-instructions p {
        color: rgba(255, 255, 255, 0.95);
        margin: 5px 0 0 0;
    }
    
    /* Tarjetas de bienvenida */
    .welcome-card {
        background: white;
        padding: 40px;
        border-radius: 20px;
        box-shadow: 0 10px 40px var(--shadow);
        text-align: center;
        border: 3px solid var(--tekpro-primary);
        transition: transform 0.3s ease;
        position: relative;
    }
    
    .welcome-card:hover {
        transform: translateY(-5px);
        box-shadow: 0 15px 50px var(--shadow);
    }
    
    /* Patrón de bloques pixelados (como en el diseño Tekpro) */
    .pixel-pattern {
        position: relative;
        overflow: hidden;
    }
    
    .pixel-pattern::before {
        content: '';
        position: absolute;
        top: -30px;
        right: -30px;
        width: 200px;
        height: 200px;
        background-image: 
            repeating-linear-gradient(
                90deg,
                rgba(125, 212, 206, 0.15) 0px,
                rgba(125, 212, 206, 0.15) 25px,
                rgba(91, 196, 188, 0.15) 25px,
                rgba(91, 196, 188, 0.15) 50px,
                rgba(62, 174, 165, 0.15) 50px,
                rgba(62, 174, 165, 0.15) 75px,
                transparent 75px,
                transparent 100px
            ),
            repeating-linear-gradient(
                0deg,
                rgba(125, 212, 206, 0.15) 0px,
                rgba(125, 212, 206, 0.15) 25px,
                rgba(91, 196, 188, 0.15) 25px,
                rgba(91, 196, 188, 0.15) 50px,
                rgba(62, 174, 165, 0.15) 50px,
                rgba(62, 174, 165, 0.15) 75px,
                transparent 75px,
                transparent 100px
            );
        transform: rotate(45deg);
        opacity: 0.6;
    }
    
    .pixel-pattern::after {
        content: '';
        position: absolute;
        bottom: -30px;
        left: -30px;
        width: 150px;
        height: 150px;
        background-image: 
            repeating-linear-gradient(
                90deg,
                rgba(125, 212, 206, 0.1) 0px,
                rgba(125, 212, 206, 0.1) 20px,
                transparent 20px,
                transparent 40px
            ),
            repeating-linear-gradient(
                0deg,
                rgba(91, 196, 188, 0.1) 0px,
                rgba(91, 196, 188, 0.1) 20px,
                transparent 20px,
                transparent 40px
            );
        opacity: 0.4;
    }
    
    /* Animaciones */
    @keyframes pulse {
        0%, 100% { transform: scale(1); }
        50% { transform: scale(1.02); }
    }
    
    @keyframes slideIn {
        from {
            opacity: 0;
            transform: translateY(-10px);
        }
        to {
            opacity: 1;
            transform: translateY(0);
        }
    }
    
    @keyframes fadeIn {
        from { opacity: 0; }
        to { opacity: 1; }
    }
    
    /* Expanders personalizados */
    .streamlit-expanderHeader {
        background-color: white;
        border-radius: 10px;
        border: 2px solid var(--tekpro-light);
        font-weight: 500;
        color: var(--tekpro-dark);
    }
    
    .streamlit-expanderHeader:hover {
        border-color: var(--tekpro-primary);
        background-color: var(--bg-light);
    }
    
    /* Inputs de Streamlit */
    .stTextInput input {
        border-radius: 10px;
        border: 2px solid var(--tekpro-light);
        padding: 12px;
        font-family: 'Poppins', sans-serif;
    }
    
    .stTextInput input:focus {
        border-color: var(--tekpro-primary);
        box-shadow: 0 0 10px var(--shadow);
    }
    
    /* Selectbox */
    .stSelectbox select {
        border-radius: 10px;
        border: 2px solid var(--tekpro-light);
        font-family: 'Poppins', sans-serif;
    }
    
    /* Badges y etiquetas */
    .badge {
        display: inline-block;
        padding: 5px 15px;
        border-radius: 20px;
        background: var(--tekpro-primary);
        color: white;
        font-size: 12px;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    
    /* Efectos de carga */
    .stSpinner > div {
        border-color: var(--tekpro-primary) !important;
    }
</style>
""", unsafe_allow_html=True)

# Configuración de archivos - Usar rutas absolutas basadas en la ubicación del script
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_FILE = os.path.join(SCRIPT_DIR, 'horas_trabajadas.csv')
CONFIG_FILE = os.path.join(SCRIPT_DIR, 'config.json')
CREDENTIALS_FILE = os.path.join(SCRIPT_DIR, 'credentials.json')

# Inicializar session state
if 'screen' not in st.session_state:
    st.session_state.screen = 'inicio'
if 'admin_mode' not in st.session_state:
    st.session_state.admin_mode = False
if 'admin_authenticated' not in st.session_state:
    st.session_state.admin_authenticated = False
if 'login_attempts' not in st.session_state:
    st.session_state.login_attempts = 0

def load_config():
    """Cargar configuración desde archivo JSON"""
    config = {
        'empleados': [],
        'codigos_barras': {},
        'admin': {
            'password': 'admin123',  # Contraseña por defecto
            'max_attempts': 3,
            'lockout_time': 300  # 5 minutos en segundos
        },
        'horarios_laborales': {
            'lunes_a_jueves': {
                'hora_entrada': '07:00',
                'hora_salida': '16:30',
                'horas_normales': 8.5,
                'tolerancia_entrada': 15,  # minutos de tolerancia
                'tolerancia_salida': 15
            },
            'viernes': {
                'hora_entrada': '07:00',
                'hora_salida': '15:30',
                'horas_normales': 7.5,
                'tolerancia_entrada': 15,
                'tolerancia_salida': 15
            },
            'sabado': {
                'hora_entrada': '07:00',
                'hora_salida': '12:00',
                'horas_normales': 5,
                'tolerancia_entrada': 15,
                'tolerancia_salida': 15
            }
        },
        'google_sheets': {
            'enabled': True,
            'spreadsheet_id': '1r3M71nQK_SxVFycYvmoeDek9KVKfBjFZuPax-v5oIb0',
            'worksheet_empleados': 'Datos_colab',
            'worksheet_servicios': 'Servicio',
            'worksheet_registros': 'Registros',
            'credentials_file': 'credentials.json'
        },
        'adecuacion_locativa': {
            'habilitado': True,
            'lunes_jueves': {
                'hora_inicio': '16:20',  # Desde las 4:20 PM
                'hora_fin': '16:30',     # Hasta las 4:30 PM
                'hora_registro': '16:30'  # Se registra como 4:30 PM
            },
            'viernes': {
                'hora_inicio': '15:20',  # Desde las 3:20 PM
                'hora_fin': '15:30',     # Hasta las 3:30 PM
                'hora_registro': '15:30'  # Se registra como 3:30 PM
            },
            'servicio_nombre': 'Adecuación Locativa',
            'servicio_codigo': '29'   # Código del servicio
        }
    }
    
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r') as f:
                saved_config = json.load(f)
                config.update(saved_config)
        except:
            pass
    
    return config

def save_config(config):
    """Guardar configuración en archivo JSON"""
    with open(CONFIG_FILE, 'w') as f:
        json.dump(config, f, indent=2)

def verificar_contraseña_admin(password):
    """Verificar contraseña de administrador"""
    config = load_config()
    admin_config = config.get('admin', {})
    contraseña_correcta = admin_config.get('password', 'admin123')
    
    return password == contraseña_correcta

def esta_bloqueado():
    """Verificar si el acceso está temporalmente bloqueado"""
    config = load_config()
    admin_config = config.get('admin', {})
    max_attempts = admin_config.get('max_attempts', 3)
    
    return st.session_state.login_attempts >= max_attempts

# ============================================
# SISTEMA DE ADECUACIÓN LOCATIVA
# Lunes a Jueves: 4:20 PM - 4:30 PM
# Viernes: 3:20 PM - 3:30 PM
# ============================================

def es_horario_adecuacion_locativa():
    """
    Verificar si la hora actual está en el rango de adecuación locativa.
    - Lunes a Jueves: 4:20 PM - 4:30 PM (se registra como 4:30 PM)
    - Viernes: 3:20 PM - 3:30 PM (se registra como 3:30 PM)
    """
    config = load_config()
    adecuacion = config.get('adecuacion_locativa', {})
    
    if not adecuacion.get('habilitado', True):
        return False, None
    
    hora_actual = obtener_hora_colombia_time()
    dia_semana = obtener_fecha_colombia().weekday()
    
    # Sábado (5) y Domingo (6) no aplica
    if dia_semana > 4:
        return False, None
    
    try:
        servicio_nombre = adecuacion.get('servicio_nombre', 'Adecuación Locativa')
        servicio_codigo = adecuacion.get('servicio_codigo', '29')
        
        # Determinar horarios según el día
        if dia_semana == 4:  # Viernes
            config_dia = adecuacion.get('viernes', {})
            hora_inicio = datetime.strptime(config_dia.get('hora_inicio', '15:20'), '%H:%M').time()
            hora_fin = datetime.strptime(config_dia.get('hora_fin', '15:30'), '%H:%M').time()
            hora_registro = config_dia.get('hora_registro', '15:30')
        else:  # Lunes a Jueves (0-3)
            config_dia = adecuacion.get('lunes_jueves', {})
            hora_inicio = datetime.strptime(config_dia.get('hora_inicio', '16:20'), '%H:%M').time()
            hora_fin = datetime.strptime(config_dia.get('hora_fin', '16:30'), '%H:%M').time()
            hora_registro = config_dia.get('hora_registro', '16:30')
        
        if hora_inicio <= hora_actual <= hora_fin:
            dia_nombre = 'Viernes' if dia_semana == 4 else 'Lunes-Jueves'
            return True, {
                'hora_registro': hora_registro,
                'servicio_nombre': servicio_nombre,
                'servicio_codigo': servicio_codigo,
                'mensaje': f'Registro automático a las {hora_registro} - {servicio_nombre} ({dia_nombre})'
            }
    except Exception as e:
        print(f"Error al verificar horario adecuación locativa: {e}")
    
    return False, None

def aplicar_adecuacion_locativa(hora_actual):
    """Aplicar la lógica de adecuación locativa y retornar hora ajustada si aplica"""
    es_adecuacion, info = es_horario_adecuacion_locativa()
    
    if es_adecuacion and info:
        # Retornar la hora fija de registro (16:30)
        hora_registro = datetime.strptime(info['hora_registro'], '%H:%M').time()
        return True, hora_registro, info
    
    return False, hora_actual, None

def obtener_servicio_adecuacion_locativa():
    """Obtener información del servicio de adecuación locativa"""
    config = load_config()
    adecuacion = config.get('adecuacion_locativa', {})
    
    return {
        'numero': adecuacion.get('servicio_codigo', '99'),
        'nomservicio': adecuacion.get('servicio_nombre', 'Adecuación Locativa'),
        'display': f"{adecuacion.get('servicio_codigo', '99')} - {adecuacion.get('servicio_nombre', 'Adecuación Locativa')}"
    }

def validar_codigo_barras(codigo):
    """Validar formato de código de barras y detectar tipo"""
    if not codigo or len(codigo.strip()) < 3:
        return False, "Código muy corto"
    
    codigo = codigo.strip()
    
    # Patrones comunes de códigos de barras
    patrones = {
        'EAN-13': lambda c: len(c) == 13 and c.isdigit(),
        'EAN-8': lambda c: len(c) == 8 and c.isdigit(),
        'UPC-A': lambda c: len(c) == 12 and c.isdigit(),
        'Code 128': lambda c: 6 <= len(c) <= 20 and c.isalnum(),
        'Code 39': lambda c: 4 <= len(c) <= 25 and all(c in '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ-. $/+%' for c in c),
        'QR Code': lambda c: 3 <= len(c) <= 100,  # QR puede contener cualquier carácter
        'Empleado ID': lambda c: 4 <= len(c) <= 15 and (c.isalnum() or '-' in c or '_' in c)
    }
    
    for tipo, validador in patrones.items():
        if validador(codigo):
            return True, f"Código válido ({tipo})"
    
    return True, "Formato personalizado"  # Aceptar otros formatos

def obtener_horario_laboral(fecha):
    """Obtener horario laboral según el día de la semana"""
    config = load_config()
    horarios = config.get('horarios_laborales', {})
    
    # Obtener día de la semana (0=lunes, 6=domingo)
    dia_semana = fecha.weekday()
    
    if dia_semana <= 3:  # Lunes a jueves (0-3)
        return horarios.get('lunes_a_jueves', {})
    elif dia_semana == 4:  # Viernes (4)
        return horarios.get('viernes', {})
    elif dia_semana == 5:  # Sábado (5)
        return horarios.get('sabado', {})
    else:  # Domingo (6)
        return None  # No hay horario laboral los domingos

def analizar_horario(hora_entrada, hora_salida, fecha):
    """Analizar el cumplimiento del horario laboral"""
    horario = obtener_horario_laboral(fecha)
    
    if not horario:
        return {
            'es_dia_laboral': False,
            'mensaje': 'Día no laboral',
            'estado': 'no_laboral'
        }
    
    # Convertir horarios de texto a objetos time
    hora_entrada_esperada = datetime.strptime(horario['hora_entrada'], '%H:%M').time()
    hora_salida_esperada = datetime.strptime(horario['hora_salida'], '%H:%M').time()
    tolerancia_entrada = horario.get('tolerancia_entrada', 15)
    tolerancia_salida = horario.get('tolerancia_salida', 15)
    horas_normales = horario.get('horas_normales', 8)
    
    # Calcular tolerancias
    entrada_con_tolerancia = datetime.combine(fecha, hora_entrada_esperada) + pd.Timedelta(minutes=tolerancia_entrada)
    entrada_con_tolerancia = entrada_con_tolerancia.time()
    
    salida_con_tolerancia = datetime.combine(fecha, hora_salida_esperada) - pd.Timedelta(minutes=tolerancia_salida)
    salida_con_tolerancia = salida_con_tolerancia.time()
    
    resultado = {
        'es_dia_laboral': True,
        'hora_entrada_esperada': hora_entrada_esperada,
        'hora_salida_esperada': hora_salida_esperada,
        'horas_normales': horas_normales,
        'tolerancia_entrada': tolerancia_entrada,
        'tolerancia_salida': tolerancia_salida
    }
    
    # Todos los empleados llegan a las 7:00 AM
    if hora_entrada:
        resultado['entrada_estado'] = 'puntual'
        resultado['entrada_mensaje'] = 'Horario normal'
    
    # Analizar salida
    if hora_salida:
        if hora_salida >= salida_con_tolerancia:
            resultado['salida_estado'] = 'completa'
            resultado['salida_mensaje'] = 'Jornada completa'
        else:
            # Calcular minutos de salida temprana
            salida_datetime = datetime.combine(fecha, hora_salida)
            esperada_datetime = datetime.combine(fecha, hora_salida_esperada)
            temprano = (esperada_datetime - salida_datetime).total_seconds() / 60
            resultado['salida_estado'] = 'temprana'
            resultado['salida_mensaje'] = f'Salida temprana ({int(temprano)} min)'
            resultado['minutos_temprano'] = int(temprano)
        
        # Calcular horas trabajadas y extras
        if hora_entrada:
            horas_trabajadas = calcular_horas(hora_entrada, hora_salida)
            resultado['horas_trabajadas'] = horas_trabajadas
            
            if horas_trabajadas > horas_normales:
                horas_extra = horas_trabajadas - horas_normales
                resultado['horas_extra'] = horas_extra
                resultado['tiene_horas_extra'] = True
            else:
                resultado['tiene_horas_extra'] = False
    
    return resultado

def load_data():
    """Cargar datos desde archivo CSV"""
    # Columnas del nuevo formato
    columnas_nuevas = [
        'fecha', 'empleado', 'cedula', 'hora_entrada', 'codigo_actividad', 
        'codigo_op', 'descripcion_proceso', 'hora_salida', 'horas_trabajadas', 
        'servicio'
    ]
    
    if os.path.exists(DATA_FILE):
        try:
            df = pd.read_csv(DATA_FILE)
            df['fecha'] = pd.to_datetime(df['fecha']).dt.date
            
            # LIMPIAR REGISTROS PROBLEMÁTICOS
            # Filtrar registros que tengan al menos cedula y hora_entrada válidos
            registros_validos = ~df['cedula'].isna() & ~df['hora_entrada'].isna()
            df_limpio = df[registros_validos].copy()
            
            # Manejar conversión de horas con valores vacíos
            if 'hora_entrada' in df_limpio.columns:
                df_limpio['hora_entrada'] = pd.to_datetime(df_limpio['hora_entrada'], format='%H:%M:%S', errors='coerce').dt.time
            if 'hora_salida' in df_limpio.columns:
                df_limpio['hora_salida'] = pd.to_datetime(df_limpio['hora_salida'], format='%H:%M:%S', errors='coerce').dt.time
            
            # Agregar nuevas columnas si no existen
            for columna in columnas_nuevas:
                if columna not in df_limpio.columns:
                    df_limpio[columna] = ''
            
            # Filtrar registros donde la conversión de hora_entrada falló
            df_limpio = df_limpio[~df_limpio['hora_entrada'].isna()]
            
            print(f"📊 [DATA CLEANING] Registros originales: {len(df)}, Registros válidos: {len(df_limpio)}")
            
            return df_limpio[columnas_nuevas]  # Asegurar el orden correcto
        except Exception as e:
            st.warning(f"Error cargando datos existentes: {e}")
            return pd.DataFrame(columns=columnas_nuevas)
    
    return pd.DataFrame(columns=columnas_nuevas)

def save_data(df):
    """Guardar datos en archivo CSV"""
    df.to_csv(DATA_FILE, index=False)

def calcular_descuento_breaks(hora_entrada, hora_salida):
    """
    Calcula el tiempo de descuento por desayuno y almuerzo según el rango trabajado.
    
    Horarios de breaks:
    - Desayuno: 9:00 AM a 9:10 AM (10 minutos = 0.167 horas)
    - Almuerzo: 12:30 PM a 1:00 PM (30 minutos = 0.5 horas)
    
    Retorna el total de horas a descontar
    """
    # Convertir strings a time si es necesario
    if isinstance(hora_entrada, str):
        try:
            hora_entrada = datetime.strptime(hora_entrada, '%H:%M:%S').time()
        except:
            return 0.0
    
    if isinstance(hora_salida, str):
        try:
            hora_salida = datetime.strptime(hora_salida, '%H:%M:%S').time()
        except:
            return 0.0
    
    if not isinstance(hora_entrada, time) or not isinstance(hora_salida, time):
        return 0.0
    
    # Horarios de breaks
    desayuno_inicio = time(9, 0)   # 9:00 AM
    desayuno_fin = time(9, 10)      # 9:10 AM
    almuerzo_inicio = time(12, 30)  # 12:30 PM
    almuerzo_fin = time(13, 0)      # 1:00 PM
    
    descuento_total = 0.0
    
    # Verificar si el rango incluye el desayuno (9:00 - 9:10)
    if hora_entrada <= desayuno_inicio and hora_salida >= desayuno_fin:
        # Todo el desayuno está dentro del rango
        descuento_total += 10 / 60  # 10 minutos = 0.167 horas
        print(f"   🍞 Desayuno descontado: 10 minutos")
    elif hora_entrada < desayuno_fin and hora_salida > desayuno_inicio:
        # Parte del desayuno está dentro del rango
        inicio_efectivo = max(hora_entrada, desayuno_inicio)
        fin_efectivo = min(hora_salida, desayuno_fin)
        
        inicio_dt = datetime.combine(date.today(), inicio_efectivo)
        fin_dt = datetime.combine(date.today(), fin_efectivo)
        minutos = (fin_dt - inicio_dt).total_seconds() / 60
        
        descuento_total += minutos / 60
        print(f"   🍞 Desayuno parcial descontado: {minutos:.0f} minutos")
    
    # Verificar si el rango incluye el almuerzo (12:30 - 1:00)
    if hora_entrada <= almuerzo_inicio and hora_salida >= almuerzo_fin:
        # Todo el almuerzo está dentro del rango
        descuento_total += 30 / 60  # 30 minutos = 0.5 horas
        print(f"   🍽️  Almuerzo descontado: 30 minutos")
    elif hora_entrada < almuerzo_fin and hora_salida > almuerzo_inicio:
        # Parte del almuerzo está dentro del rango
        inicio_efectivo = max(hora_entrada, almuerzo_inicio)
        fin_efectivo = min(hora_salida, almuerzo_fin)
        
        inicio_dt = datetime.combine(date.today(), inicio_efectivo)
        fin_dt = datetime.combine(date.today(), fin_efectivo)
        minutos = (fin_dt - inicio_dt).total_seconds() / 60
        
        descuento_total += minutos / 60
        print(f"   🍽️  Almuerzo parcial descontado: {minutos:.0f} minutos")
    
    return descuento_total

def calcular_horas(hora_entrada, hora_salida, descontar_breaks=True):
    """
    Calcular horas trabajadas entre entrada y salida.
    Por defecto descuenta automáticamente desayuno y almuerzo.
    """
    # Validar que ambos valores sean objetos time válidos
    if pd.isna(hora_entrada) or pd.isna(hora_salida):
        return 0.0
    
    # Convertir strings a time si es necesario
    if isinstance(hora_entrada, str):
        try:
            hora_entrada = datetime.strptime(hora_entrada, '%H:%M:%S').time()
        except:
            return 0.0
    
    if isinstance(hora_salida, str):
        try:
            hora_salida = datetime.strptime(hora_salida, '%H:%M:%S').time()
        except:
            return 0.0
    
    # Verificar que son objetos time
    if not isinstance(hora_entrada, time) or not isinstance(hora_salida, time):
        return 0.0
    
    entrada = datetime.combine(date.today(), hora_entrada)
    salida = datetime.combine(date.today(), hora_salida)
    
    if salida < entrada:
        salida = datetime.combine(date.today(), hora_salida) + pd.Timedelta(days=1)
    
    diferencia = salida - entrada
    horas_brutas = diferencia.total_seconds() / 3600
    
    # Aplicar descuentos de breaks si está habilitado
    if descontar_breaks:
        descuento = calcular_descuento_breaks(hora_entrada, hora_salida)
        horas_netas = horas_brutas - descuento
        
        if descuento > 0:
            print(f"   ⏱️  Horas brutas: {horas_brutas:.3f} | Descuento: {descuento:.3f} | Netas: {horas_netas:.3f}")
        
        return max(0, horas_netas)  # No permitir valores negativos
    
    return horas_brutas

def calcular_horas_conteo_diario(empleado_cedula, fecha_registro, hora_registro, hora_forzada=None):
    """
    Nueva lógica de conteos basada en verificación directa desde Google Sheets:
    - Consulta el Sheet al momento de leer la cédula
    - Si no hay registros del día: calcula desde 7:00 AM
    - Si ya hay registros del día: calcula desde la última hora_exacta del día
    - hora_forzada: Si se proporciona, usa esta hora en lugar de la actual (para adecuación locativa)
    """
    hora_inicio_dia = time(7, 0)  # 7:00 AM por defecto
    
    # Usar hora forzada si está disponible (adecuación locativa), sino hora actual
    if hora_forzada:
        hora_actual_exacta = hora_forzada
    else:
        hora_actual_exacta = obtener_hora_colombia_time()  # Hora exacta actual en Colombia
    
    cedula_str = str(empleado_cedula).strip()
    
    print(f"\n{'='*60}")
    print(f"🔍 [ANÁLISIS AUTOMÁTICO] Analizando cédula {cedula_str} para fecha {fecha_registro}")
    print(f"{'='*60}")
    
    # NUEVA LÓGICA: Verificar directamente en Google Sheets
    registros_del_dia, es_primer_registro_del_dia, ultima_hora_exacta = verificar_registros_del_dia_en_sheets(
        cedula_str, 
        fecha_registro
    )
    
    if es_primer_registro_del_dia:
        # PRIMER REGISTRO DEL DÍA - Contar desde 7:00 AM
        tiempo_trabajado = calcular_horas(hora_inicio_dia, hora_actual_exacta)
        
        print(f"✨ [PRIMER REGISTRO DEL DÍA]")
        print(f"   📅 Fecha: {fecha_registro}")
        print(f"   🕐 Hora inicio: 7:00 AM")
        print(f"   🕐 Hora actual: {hora_actual_exacta}")
        print(f"   ⏱️  Tiempo trabajado: {tiempo_trabajado:.3f} horas")
        print(f"{'='*60}\n")
        
        return {
            'es_primer_registro': True,  # True para primer registro del día
            'es_primer_registro_del_dia': True,
            'hora_inicio_conteo': hora_inicio_dia,
            'hora_fin_conteo': hora_actual_exacta,
            'tiempo_trabajado': round(tiempo_trabajado, 3),
            'registro_anterior_actualizado': None,
            'hora_exacta_registro': hora_actual_exacta.strftime('%H:%M:%S')
        }
    
    else:
        # REGISTRO ADICIONAL DEL DÍA - Contar desde última hora_exacta del día
        if ultima_hora_exacta:
            try:
                ultima_hora_obj = datetime.strptime(ultima_hora_exacta, '%H:%M:%S').time()
            except:
                # Si falla el parsing, usar 7 AM como fallback
                ultima_hora_obj = hora_inicio_dia
        else:
            ultima_hora_obj = hora_inicio_dia
        
        tiempo_trabajado = calcular_horas(ultima_hora_obj, hora_actual_exacta)
        
        print(f"🔄 [REGISTRO ADICIONAL DEL DÍA]")
        print(f"   📅 Fecha: {fecha_registro}")
        print(f"   📊 Registros previos hoy: {len(registros_del_dia)}")
        print(f"   🕐 Última hora del día: {ultima_hora_exacta}")
        print(f"   🕐 Hora actual: {hora_actual_exacta}")
        print(f"   ⏱️  Tiempo adicional: {tiempo_trabajado:.3f} horas")
        print(f"{'='*60}\n")
        
        return {
            'es_primer_registro': False,
            'es_primer_registro_del_dia': False,
            'hora_inicio_conteo': ultima_hora_obj,
            'hora_fin_conteo': hora_actual_exacta,
            'tiempo_trabajado': round(tiempo_trabajado, 3),
            'registro_anterior_actualizado': None,
            'hora_exacta_registro': hora_actual_exacta.strftime('%H:%M:%S')
        }

def obtener_resumen_dia_empleado(empleado_cedula, fecha_registro):
    """Obtener resumen completo del día para un empleado"""
    df = load_data()
    cedula_str = str(empleado_cedula).strip()
    
    registros_del_dia = df[
        (df['cedula'].astype(str).str.strip() == cedula_str) & 
        (df['fecha'].astype(str) == str(fecha_registro))
    ].sort_values('hora_entrada')
    
    if len(registros_del_dia) == 0:
        return {
            'total_registros': 0,
            'tiempo_total_trabajado': 0,
            'primer_registro': None,
            'ultimo_registro': None,
            'registros_detalle': []
        }
    
    # Calcular tiempo total trabajado
    tiempo_total = 0
    registros_detalle = []
    
    for i, (idx, registro) in enumerate(registros_del_dia.iterrows()):
        horas_trabajadas = registro.get('horas_trabajadas', 0)
        if pd.isna(horas_trabajadas):
            horas_trabajadas = 0
        else:
            horas_trabajadas = float(horas_trabajadas)
            
        tiempo_total += horas_trabajadas
        
        registros_detalle.append({
            'numero': i + 1,
            'op': registro.get('op', 'N/A'),
            'hora_entrada': registro.get('hora_entrada', 'N/A'),
            'hora_salida': registro.get('hora_salida', 'N/A'),
            'hora_exacta': registro.get('hora_exacta', 'N/A'),
            'horas_trabajadas': horas_trabajadas,
            'estado': 'Cerrado' if registro.get('hora_salida') else 'Abierto'
        })
    
    return {
        'total_registros': len(registros_del_dia),
        'tiempo_total_trabajado': round(tiempo_total, 2),
        'primer_registro': registros_del_dia.iloc[0],
        'ultimo_registro': registros_del_dia.iloc[-1],
        'registros_detalle': registros_detalle
    }

def buscar_empleado_por_codigo(codigo_barras):
    """Buscar empleado por código de barras en configuración local"""
    config = load_config()
    codigos = config.get('codigos_barras', {})
    return codigos.get(codigo_barras, None)

def diagnosticar_conexion_sheets():
    """Diagnosticar el estado de la conexión con Google Sheets"""
    config = load_config()
    gs_config = config.get('google_sheets', {})
    
    diagnosticos = []
    
    # Verificar configuración habilitada
    if not gs_config.get('enabled', False):
        diagnosticos.append("❌ Google Sheets está deshabilitado en config.json")
        return diagnosticos
    else:
        diagnosticos.append("✅ Google Sheets habilitado en configuración")
    
    # Verificar archivo de credenciales
    credentials_file = gs_config.get('credentials_file', '')
    
    # Si la ruta es relativa, convertir a absoluta
    if credentials_file and not os.path.isabs(credentials_file):
        credentials_file = os.path.join(SCRIPT_DIR, credentials_file)
    
    # Si no hay archivo en config, usar el predeterminado
    if not credentials_file:
        credentials_file = CREDENTIALS_FILE
    
    if not os.path.exists(credentials_file):
        diagnosticos.append(f"❌ Archivo de credenciales no encontrado: {credentials_file}")
        diagnosticos.append("💡 Crea el archivo credentials.json siguiendo CONFIGURACION_GOOGLE_SHEETS.md")
        return diagnosticos
    else:
        diagnosticos.append(f"✅ Archivo de credenciales encontrado: {credentials_file}")
    
    # Verificar ID de spreadsheet
    spreadsheet_id = gs_config.get('spreadsheet_id', '')
    if not spreadsheet_id:
        diagnosticos.append("❌ ID de Google Sheets no configurado")
        return diagnosticos
    else:
        diagnosticos.append(f"✅ ID de Google Sheets configurado: {spreadsheet_id}")
    
    # Intentar conexión real
    try:
        scope = [
            'https://spreadsheets.google.com/feeds',
            'https://www.googleapis.com/auth/drive'
        ]
        
        credentials = Credentials.from_service_account_file(credentials_file, scopes=scope)
        gc = gspread.authorize(credentials)
        
        spreadsheet = gc.open_by_key(spreadsheet_id)
        diagnosticos.append("✅ Conexión exitosa con Google Sheets")
        
        # Verificar hojas específicas
        try:
            worksheet_empleados = spreadsheet.worksheet(gs_config.get('worksheet_empleados', 'Datos_colab'))
            diagnosticos.append(f"✅ Hoja '{gs_config.get('worksheet_empleados', 'Datos_colab')}' encontrada")
        except:
            diagnosticos.append(f"❌ Hoja '{gs_config.get('worksheet_empleados', 'Datos_colab')}' no encontrada")
        
        try:
            worksheet_servicios = spreadsheet.worksheet(gs_config.get('worksheet_servicios', 'Servicio'))
            diagnosticos.append(f"✅ Hoja '{gs_config.get('worksheet_servicios', 'Servicio')}' encontrada")
        except:
            diagnosticos.append(f"❌ Hoja '{gs_config.get('worksheet_servicios', 'Servicio')}' no encontrada")
        
    except Exception as e:
        diagnosticos.append(f"❌ Error de conexión: {str(e)}")
        
        # Diagnósticos específicos de errores comunes
        error_str = str(e).lower()
        if "permission denied" in error_str or "forbidden" in error_str:
            diagnosticos.append("💡 Problema de permisos: Comparte el Google Sheet con el client_email de las credenciales")
        elif "not found" in error_str:
            diagnosticos.append("💡 Google Sheet no encontrado: Verifica el spreadsheet_id")
        elif "api not enabled" in error_str:
            diagnosticos.append("💡 API no habilitada: Habilita Google Sheets API y Google Drive API")
    
    return diagnosticos

def conectar_google_sheets():
    """Conectar a Google Sheets usando las credenciales configuradas (local o Streamlit Cloud)"""
    config = load_config()
    gs_config = config.get('google_sheets', {})
    
    if not gs_config.get('enabled', False):
        return None, "Google Sheets no está habilitado"
    
    try:
        scope = [
            'https://spreadsheets.google.com/feeds',
            'https://www.googleapis.com/auth/drive'
        ]
        
        credentials = None
        
        # OPCIÓN 1: Intentar usar Streamlit Secrets (para Streamlit Cloud)
        try:
            if hasattr(st, 'secrets') and 'gcp_service_account' in st.secrets:
                from google.oauth2.service_account import Credentials
                credentials_dict = dict(st.secrets["gcp_service_account"])
                credentials = Credentials.from_service_account_info(credentials_dict, scopes=scope)
        except Exception as e:
            pass  # Si falla, intentar con archivo local
        
        # OPCIÓN 2: Usar archivo local credentials.json
        if credentials is None:
            credentials_file = gs_config.get('credentials_file', '')
            
            # Si la ruta es relativa, convertir a absoluta
            if credentials_file and not os.path.isabs(credentials_file):
                credentials_file = os.path.join(SCRIPT_DIR, credentials_file)
            
            # Si no hay archivo en config, usar el predeterminado
            if not credentials_file:
                credentials_file = CREDENTIALS_FILE
            
            if not os.path.exists(credentials_file):
                return None, f"Archivo de credenciales no encontrado: {credentials_file}"
            
            credentials = Credentials.from_service_account_file(credentials_file, scopes=scope)
        
        gc = gspread.authorize(credentials)
        
        spreadsheet_id = gs_config.get('spreadsheet_id', '')
        if not spreadsheet_id:
            return None, "ID de Google Sheets no configurado"
        
        spreadsheet = gc.open_by_key(spreadsheet_id)
        return spreadsheet, "Conexión exitosa"
        
    except Exception as e:
        return None, f"Error al conectar: {str(e)}"

def obtener_ultimo_registro_sheets(cedula):
    """Buscar el último registro de un empleado en la hoja 'Registros' de Google Sheets y extraer hora_exacta"""
    spreadsheet, mensaje = conectar_google_sheets()
    
    if spreadsheet is None:
        return None, mensaje
    
    try:
        config = load_config()
        worksheet_name = config.get('google_sheets', {}).get('worksheet_registros', 'Registros')
        worksheet = spreadsheet.worksheet(worksheet_name)
        
        # Obtener todos los registros
        try:
            all_values = worksheet.get_all_values()
            if len(all_values) < 2:
                return None, "La hoja 'Registros' está vacía"
            
            headers = all_values[0]
            rows = all_values[1:]
            
            # Encontrar el índice de las columnas relevantes
            try:
                cedula_idx = headers.index('Cédula')
            except ValueError:
                try:
                    cedula_idx = headers.index('cedula')
                except ValueError:
                    return None, "No se encontró la columna 'Cédula' en la hoja 'Registros'"
            
            try:
                hora_exacta_idx = headers.index('hora_exacta')
            except ValueError:
                return None, "No se encontró la columna 'hora_exacta' en la hoja 'Registros'"
            
            # Buscar registros de esta cédula (del más reciente al más antiguo)
            ultimo_registro = None
            for row in reversed(rows):  # Recorrer desde el final
                if len(row) > cedula_idx:
                    cedula_en_fila = str(row[cedula_idx]).strip()
                    if cedula_en_fila == str(cedula).strip():
                        # Encontramos el último registro de esta cédula
                        if len(row) > hora_exacta_idx and row[hora_exacta_idx]:
                            hora_exacta = row[hora_exacta_idx].strip()
                            if hora_exacta:
                                # Construir objeto con información del registro
                                ultimo_registro = {
                                    'hora_exacta': hora_exacta,
                                    'fila_completa': row,
                                    'headers': headers
                                }
                                break
            
            if ultimo_registro:
                return ultimo_registro, "Último registro encontrado"
            else:
                return None, "No se encontraron registros previos para esta cédula"
                
        except Exception as e:
            return None, f"Error al leer la hoja 'Registros': {str(e)}"
            
    except Exception as e:
        return None, f"Error al buscar último registro: {str(e)}"

def verificar_registros_del_dia_en_sheets(cedula, fecha_actual):
    """
    Verifica en Google Sheets si la cédula ya tiene registros en el día actual.
    Retorna:
    - registros_del_dia: lista de registros del día actual
    - es_primer_registro_del_dia: True si no hay registros del día, False si ya hay
    - ultima_hora_exacta_del_dia: la hora_exacta del último registro del día (si existe)
    """
    spreadsheet, mensaje = conectar_google_sheets()
    
    if spreadsheet is None:
        print(f"⚠️ [VERIFICACIÓN SHEETS] No se pudo conectar: {mensaje}")
        return [], True, None  # Asumir primer registro si no hay conexión
    
    try:
        config = load_config()
        worksheet_name = config.get('google_sheets', {}).get('worksheet_registros', 'Registros')
        worksheet = spreadsheet.worksheet(worksheet_name)
        
        all_values = worksheet.get_all_values()
        if len(all_values) < 2:
            print(f"✅ [VERIFICACIÓN SHEETS] Hoja vacía - PRIMER REGISTRO DEL DÍA")
            return [], True, None
        
        headers = all_values[0]
        rows = all_values[1:]
        
        # Encontrar índices de columnas
        cedula_idx = None
        fecha_idx = None
        hora_exacta_idx = None
        
        for idx, header in enumerate(headers):
            header_lower = header.lower().strip()
            if 'cédula' in header_lower or 'cedula' in header_lower:
                cedula_idx = idx
            elif 'fecha' in header_lower:
                fecha_idx = idx
            elif 'hora_exacta' in header_lower:
                hora_exacta_idx = idx
        
        if cedula_idx is None or fecha_idx is None:
            print(f"⚠️ [VERIFICACIÓN SHEETS] Columnas no encontradas")
            return [], True, None
        
        # Formatear fecha actual para comparación
        if hasattr(fecha_actual, 'strftime'):
            fecha_str = fecha_actual.strftime('%d/%m/%Y')
        else:
            fecha_str = str(fecha_actual)
        
        cedula_str = str(cedula).strip()
        registros_del_dia = []
        
        # Buscar todos los registros de esta cédula en esta fecha
        for row in rows:
            if len(row) > max(cedula_idx, fecha_idx):
                cedula_en_fila = str(row[cedula_idx]).strip()
                fecha_en_fila = str(row[fecha_idx]).strip()
                
                # Comparar cédula y fecha
                if cedula_en_fila == cedula_str and fecha_en_fila == fecha_str:
                    registro_info = {
                        'fila': row,
                        'hora_exacta': row[hora_exacta_idx].strip() if hora_exacta_idx and len(row) > hora_exacta_idx else None
                    }
                    registros_del_dia.append(registro_info)
        
        # Determinar si es primer registro del día
        es_primer_registro = len(registros_del_dia) == 0
        ultima_hora_exacta = None
        
        if not es_primer_registro and registros_del_dia[-1]['hora_exacta']:
            ultima_hora_exacta = registros_del_dia[-1]['hora_exacta']
        
        if es_primer_registro:
            print(f"✅ [VERIFICACIÓN SHEETS] Cédula {cedula_str} - PRIMER REGISTRO DEL DÍA {fecha_str}")
        else:
            print(f"🔄 [VERIFICACIÓN SHEETS] Cédula {cedula_str} - Ya tiene {len(registros_del_dia)} registro(s) del día {fecha_str}")
            print(f"🔄 [VERIFICACIÓN SHEETS] Última hora exacta del día: {ultima_hora_exacta}")
        
        return registros_del_dia, es_primer_registro, ultima_hora_exacta
        
    except Exception as e:
        print(f"⚠️ [ERROR VERIFICACIÓN SHEETS] {str(e)}")
        return [], True, None  # En caso de error, asumir primer registro

def buscar_colaborador_en_datos_colab(codigo_barras):
    """Buscar colaborador en la hoja 'Datos_colab' de Google Sheets"""
    spreadsheet, mensaje = conectar_google_sheets()
    
    if spreadsheet is None:
        return None, mensaje
    
    try:
        config = load_config()
        worksheet_name = config.get('google_sheets', {}).get('worksheet_empleados', 'Datos_colab')
        worksheet = spreadsheet.worksheet(worksheet_name)
        
        # Obtener datos usando los encabezados esperados para evitar duplicados
        try:
            records = worksheet.get_all_records(expected_headers=['cedula', 'nombre'])
        except:
            # Si falla con encabezados específicos, intentar con todos los datos
            try:
                records = worksheet.get_all_records()
            except:
                # Si todo falla, obtener datos raw y procesarlos manualmente
                all_values = worksheet.get_all_values()
                if len(all_values) < 2:
                    return None, "La hoja Datos_colab está vacía"
                
                headers = all_values[0]
                rows = all_values[1:]
                
                records = []
                for row in rows:
                    record = {}
                    for i, header in enumerate(headers):
                        if i < len(row) and header.strip():  # Solo usar encabezados no vacíos
                            record[header.strip()] = row[i] if i < len(row) else ''
                    if record:  # Solo agregar registros no vacíos
                        records.append(record)
        
        for record in records:
            # Buscar en la columna 'cedula' (código de barras)
            if 'cedula' in record and str(record['cedula']) == str(codigo_barras):
                cedula_encontrada = record['cedula']
                nombre_encontrado = record.get('nombre', '')
                
                if nombre_encontrado:
                    return nombre_encontrado, "Colaborador encontrado en Google Sheets"
        
        return None, "Colaborador no encontrado en Google Sheets"
        
    except Exception as e:
        return None, f"Error al buscar en Google Sheets: {str(e)}"

def buscar_servicio_por_codigo(codigo_barras):
    """Buscar servicio por código de barras en la hoja 'Servicio'"""
    spreadsheet, mensaje = conectar_google_sheets()
    
    if spreadsheet is None:
        return None, None, f"Error de conexión: {mensaje}"
    
    try:
        worksheet = spreadsheet.worksheet('Servicio')
        
        # Obtener datos de manera más robusta para evitar encabezados duplicados
        try:
            # Intentar con encabezados esperados específicos
            records = worksheet.get_all_records(expected_headers=['codigo', 'actividad'])
        except:
            try:
                # Si falla, obtener datos raw y procesarlos manualmente
                all_values = worksheet.get_all_values()
                if len(all_values) < 2:
                    return None, None, "La hoja 'Servicio' está vacía o solo tiene encabezados"
                
                headers = all_values[0]
                rows = all_values[1:]
                
                records = []
                for row in rows:
                    record = {}
                    for i, header in enumerate(headers):
                        if i < len(row) and header.strip():  # Solo usar encabezados no vacíos
                            record[header.strip()] = row[i] if i < len(row) else ''
                    if record:  # Solo agregar registros no vacíos
                        records.append(record)
            except Exception as e:
                return None, None, f"Error al leer datos de la hoja 'Servicio': {str(e)}"
        
        # Verificar que hay datos
        if not records:
            return None, None, "La hoja 'Servicio' está vacía"
        
        # Obtener las claves y limpiarlas de espacios
        encabezados_originales = list(records[0].keys())
        encabezados_limpios = [str(k).strip().lower() for k in encabezados_originales]
        
        # Buscar los encabezados de forma más flexible
        columna_codigo = None
        columna_actividad = None
        
        for i, encabezado in enumerate(encabezados_limpios):
            if encabezado == 'codigo':
                columna_codigo = encabezados_originales[i]
            elif encabezado == 'actividad':
                columna_actividad = encabezados_originales[i]
        
        # Verificar que encontramos ambas columnas
        if columna_codigo is None:
            return None, None, f"Columna 'codigo' no encontrada. Encabezados disponibles: {encabezados_originales}"
        
        if columna_actividad is None:
            return None, None, f"Columna 'actividad' no encontrada. Encabezados disponibles: {encabezados_originales}"
        
        # Buscar el código
        for record in records:
            if columna_codigo in record:
                codigo_en_hoja = str(record[columna_codigo]).strip()
                codigo_buscado = str(codigo_barras).strip()
                
                if codigo_en_hoja == codigo_buscado:
                    codigo_encontrado = record[columna_codigo]
                    actividad_encontrada = str(record.get(columna_actividad, '')).strip()
                    
                    if actividad_encontrada:
                        return codigo_encontrado, actividad_encontrada, "Servicio encontrado"
                    else:
                        return None, None, f"Código encontrado pero actividad vacía para código '{codigo_barras}'"
        
        return None, None, f"Código '{codigo_barras}' no encontrado en {len(records)} registros"
        
    except Exception as e:
        return None, None, f"Error al buscar servicio: {str(e)}"

def buscar_op_por_codigo(codigo_barras):
    """Buscar OP por código de barras en la hoja 'OPS' y traer toda la información"""
    spreadsheet, mensaje = conectar_google_sheets()
    
    if spreadsheet is None:
        return None, f"Error de conexión: {mensaje}"
    
    try:
        worksheet = spreadsheet.worksheet('OPS')
        
        # Obtener datos de manera robusta para evitar encabezados duplicados
        try:
            records = worksheet.get_all_records(expected_headers=['orden', 'referencia', 'Cantidades', 'cliente', 'item'])
        except:
            try:
                all_values = worksheet.get_all_values()
                if len(all_values) < 2:
                    return None, "La hoja 'OPS' está vacía"
                
                headers = all_values[0]
                rows = all_values[1:]
                
                records = []
                for row in rows:
                    record = {}
                    for i, header in enumerate(headers):
                        if i < len(row) and header.strip():
                            record[header.strip()] = row[i] if i < len(row) else ''
                    if record:
                        records.append(record)
            except Exception as e:
                return None, f"Error al leer hoja 'OPS': {str(e)}"
        
        for record in records:
            # Buscar en la columna 'orden' (código de barras)
            if 'orden' in record and str(record['orden']) == str(codigo_barras):
                # Si se encontró la OP, extraer toda la información
                op_info = {
                    'orden': str(record.get('orden', '')),
                    'referencia': str(record.get('referencia', '')),
                    'cantidades': str(record.get('Cantidades', '')),
                    'cliente': str(record.get('cliente', '')),
                    'item': str(record.get('item', ''))
                }
                
                return op_info, "OP encontrada con información completa"
        
        return None, "OP no encontrada en la hoja OPS"
        
    except Exception as e:
        return None, f"Error al buscar OP: {str(e)}"

def verificar_estructura_servicio():
    """Función de debug para verificar la estructura de la hoja Servicio"""
    spreadsheet, mensaje = conectar_google_sheets()
    
    if spreadsheet is None:
        return f"Error de conexión: {mensaje}"
    
    try:
        worksheet = spreadsheet.worksheet('Servicio')
        
        # Obtener encabezados (primera fila)
        encabezados = worksheet.row_values(1)
        
        # Obtener algunos registros de ejemplo
        all_values = worksheet.get_all_values()
        
        info_debug = []
        info_debug.append(f"📊 Total de filas: {len(all_values)}")
        info_debug.append(f"📋 Encabezados raw: {encabezados}")
        
        # Mostrar cada encabezado con su representación
        for i, enc in enumerate(encabezados):
            info_debug.append(f"  Columna {i}: '{enc}' (len={len(enc)}, repr={repr(enc)})")
        
        if len(all_values) > 1:
            info_debug.append(f"📄 Primera fila de datos: {all_values[1]}")
            if len(all_values) > 2:
                info_debug.append(f"📄 Segunda fila de datos: {all_values[2]}")
        
        # Obtener records para ver la estructura
        records = worksheet.get_all_records()
        if records:
            info_debug.append(f"📝 Primer registro: {records[0]}")
            info_debug.append(f"🔑 Claves: {list(records[0].keys())}")
            
            # Verificar específicamente las claves que buscamos
            claves_disponibles = list(records[0].keys())
            for clave in claves_disponibles:
                clave_limpia = str(clave).strip().lower()
                info_debug.append(f"  '{clave}' → '{clave_limpia}'")
        
        return "\n".join(info_debug)
        
    except Exception as e:
        return f"Error al verificar estructura: {str(e)}"

def obtener_servicios():
    """Obtener lista de servicios de la hoja 'Servicio' de Google Sheets"""
    spreadsheet, mensaje = conectar_google_sheets()
    
    if spreadsheet is None:
        return [], mensaje
    
    try:
        worksheet = spreadsheet.worksheet('Servicio')
        
        # Obtener datos de manera robusta para evitar encabezados duplicados
        try:
            records = worksheet.get_all_records(expected_headers=['codigo', 'actividad'])
        except:
            try:
                all_values = worksheet.get_all_values()
                if len(all_values) < 2:
                    return [], "La hoja 'Servicio' está vacía"
                
                headers = all_values[0]
                rows = all_values[1:]
                
                records = []
                for row in rows:
                    record = {}
                    for i, header in enumerate(headers):
                        if i < len(row) and header.strip():
                            record[header.strip()] = row[i] if i < len(row) else ''
                    if record:
                        records.append(record)
            except Exception as e:
                return [], f"Error al leer hoja 'Servicio': {str(e)}"
        
        if not records:
            return [], "La hoja 'Servicio' está vacía"
        
        # Obtener las claves y limpiarlas de espacios
        encabezados_originales = list(records[0].keys())
        encabezados_limpios = [str(k).strip().lower() for k in encabezados_originales]
        
        # Buscar los encabezados de forma más flexible
        columna_codigo = None
        columna_actividad = None
        
        for i, encabezado in enumerate(encabezados_limpios):
            if encabezado == 'codigo':
                columna_codigo = encabezados_originales[i]
            elif encabezado == 'actividad':
                columna_actividad = encabezados_originales[i]
        
        if not columna_codigo or not columna_actividad:
            return [], f"Encabezados no encontrados. Disponibles: {encabezados_originales}"
        
        servicios = []
        for record in records:
            codigo = str(record.get(columna_codigo, '')).strip()
            actividad = str(record.get(columna_actividad, '')).strip()
            
            if codigo and actividad:
                servicios.append({
                    'numero': codigo,
                    'nomservicio': actividad,
                    'display': f"{codigo} - {actividad}"
                })
        
        return servicios, "Servicios cargados exitosamente"
        
    except Exception as e:
        return [], f"Error al obtener servicios: {str(e)}"

def obtener_hora_inicio_dia():
    """Obtiene la hora de inicio del día laboral (7:00 AM)"""
    return time(7, 0, 0)

def calcular_horas_desde_inicio_dia(hora_actual, fecha_actual):
    """Calcula las horas transcurridas desde las 7:00 AM del día actual"""
    hora_inicio = obtener_hora_inicio_dia()
    
    # Crear datetime completos para el cálculo
    datetime_inicio = datetime.combine(fecha_actual, hora_inicio)
    datetime_actual = datetime.combine(fecha_actual, hora_actual)
    
    # Si la hora actual es antes de las 7:00 AM, consideramos que es del día anterior
    if hora_actual < hora_inicio:
        datetime_actual = datetime.combine(fecha_actual, hora_actual)
        datetime_inicio = datetime.combine(fecha_actual, hora_inicio)
        return 0  # No hay horas si es antes del inicio
    
    diferencia = datetime_actual - datetime_inicio
    return diferencia.total_seconds() / 3600

def obtener_ultimo_registro_del_dia(empleado, fecha_actual, df):
    """Obtiene el último registro del empleado en el día actual"""
    registros_del_dia = df[
        (df['empleado'] == empleado) & 
        (df['fecha'] == fecha_actual)
    ].sort_values('hora_entrada', ascending=False)
    
    if not registros_del_dia.empty:
        return registros_del_dia.iloc[0]
    return None

def obtener_ultimo_registro_por_cedula(cedula, fecha_actual, df):
    """Obtiene el último registro basado en la cédula del colaborador en el día actual"""
    if df.empty:
        return None
    
    # Debug: imprimir información para diagnosticar
    print(f"DEBUG: Buscando registros para cédula: {cedula}, fecha: {fecha_actual}")
    print(f"DEBUG: Total registros en df: {len(df)}")
    
    if 'cedula' in df.columns:
        print(f"DEBUG: Cédulas únicas en df: {df['cedula'].unique()}")
    if 'fecha' in df.columns:
        print(f"DEBUG: Fechas únicas en df: {df['fecha'].unique()}")
    
    # Filtrar registros más robustamente
    try:
        # Asegurar que las columnas existen
        if 'cedula' not in df.columns or 'fecha' not in df.columns:
            print(f"DEBUG: Columnas faltantes. Disponibles: {df.columns.tolist()}")
            return None
        
        # Convertir fecha si es necesario
        if hasattr(fecha_actual, 'date'):
            fecha_buscar = fecha_actual.date()
        else:
            fecha_buscar = fecha_actual
            
        # Filtrar por cédula y fecha
        registros_por_cedula = df[df['cedula'].astype(str).str.strip() == str(cedula).strip()]
        print(f"DEBUG: Registros con cédula {cedula}: {len(registros_por_cedula)}")
        
        if not registros_por_cedula.empty:
            # Filtrar por fecha
            registros_del_dia = registros_por_cedula[
                registros_por_cedula['fecha'].astype(str) == str(fecha_buscar)
            ]
            print(f"DEBUG: Registros del día {fecha_buscar}: {len(registros_del_dia)}")
            
            if not registros_del_dia.empty and 'hora_entrada' in registros_del_dia.columns:
                ultimo_registro = registros_del_dia.sort_values('hora_entrada', ascending=False).iloc[0]
                print(f"DEBUG: Último registro encontrado: {ultimo_registro.to_dict()}")
                return ultimo_registro
        
        print("DEBUG: No se encontraron registros anteriores")
        return None
        
    except Exception as e:
        print(f"DEBUG: Error en obtener_ultimo_registro_por_cedula: {str(e)}")
        return None

def registrar_actividad_continua(empleado, codigo_barras, servicio_info=None):
    """Registrar actividad continua desde las 7:00 AM o último registro"""
    df = load_data()
    fecha_actual = obtener_fecha_colombia()
    hora_actual = obtener_hora_colombia_time()
    
    # Obtener el último registro del día
    ultimo_registro = obtener_ultimo_registro_del_dia(empleado, fecha_actual, df)
    
    if ultimo_registro is not None:
        # Ya hay registros del día, calcular desde el último registro
        if pd.isna(ultimo_registro['hora_salida']) or ultimo_registro['hora_salida'] == '':
            # El último registro no tiene salida, cerrarlo primero
            idx = ultimo_registro.name
            df.loc[idx, 'hora_salida'] = hora_actual.strftime('%H:%M:%S') if hasattr(hora_actual, 'strftime') else str(hora_actual)
            
            # Calcular horas de la actividad anterior
            hora_entrada_anterior = ultimo_registro['hora_entrada']
            if isinstance(hora_entrada_anterior, str):
                hora_entrada_anterior = datetime.strptime(hora_entrada_anterior, '%H:%M:%S').time()
            
            horas_actividad_anterior = calcular_horas(hora_entrada_anterior, hora_actual)
            df.loc[idx, 'horas_trabajadas'] = round(horas_actividad_anterior, 2)
        
        # Hora de inicio para el nuevo registro es la misma hora actual
        hora_inicio_nueva_actividad = hora_actual
    else:
        # Es el primer registro del día, iniciar desde las 7:00 AM
        hora_inicio_nueva_actividad = obtener_hora_inicio_dia()
        
        # Si es después de las 7:00 AM, calculamos las horas desde las 7:00 AM
        if hora_actual > hora_inicio_nueva_actividad:
            horas_desde_inicio = calcular_horas_desde_inicio_dia(hora_actual, fecha_actual)
        else:
            # Si es antes de las 7:00 AM, usar la hora actual como inicio
            hora_inicio_nueva_actividad = hora_actual
            horas_desde_inicio = 0
    
    # Crear nuevo registro para la nueva actividad
    # Analizar horario laboral
    analisis = analizar_horario(hora_actual, None, fecha_actual)
    
    nuevo_registro = {
        'fecha': fecha_actual,
        'empleado': empleado,
        'hora_entrada': hora_inicio_nueva_actividad,
        'hora_salida': '',
        'horas_trabajadas': '',
        'servicio': servicio_info if servicio_info else ''
    }
    
    df = pd.concat([df, pd.DataFrame([nuevo_registro])], ignore_index=True)
    save_data(df)
    
    # Calcular horas totales del día hasta ahora
    horas_totales_dia = calcular_horas_desde_inicio_dia(hora_actual, fecha_actual)
    
    return 'actividad_registrada', horas_totales_dia, analisis

def obtener_resumen_actividades_dia(empleado, fecha_actual):
    """Obtiene un resumen de todas las actividades del empleado en el día"""
    df = load_data()
    registros_dia = df[
        (df['empleado'] == empleado) & 
        (df['fecha'] == fecha_actual)
    ].sort_values('hora_entrada')
    
    if registros_dia.empty:
        return {
            'total_actividades': 0,
            'horas_total_trabajadas': 0,
            'actividades': [],
            'inicio_dia': '07:00',
            'estado_actual': 'sin_registros'
        }
    
    actividades = []
    total_horas = 0
    
    for _, registro in registros_dia.iterrows():
        actividad = {
            'servicio': registro.get('servicio', 'Sin especificar'),
            'hora_inicio': registro['hora_entrada'],
            'hora_fin': registro.get('hora_salida', 'En curso'),
            'horas': registro.get('horas_trabajadas', 0) if pd.notna(registro.get('horas_trabajadas')) else 0
        }
        
        if actividad['horas'] and actividad['horas'] != '':
            total_horas += float(actividad['horas'])
        
        actividades.append(actividad)
    
    # Verificar si hay actividad en curso
    ultima_actividad = registros_dia.iloc[-1]
    estado_actual = 'terminada' if pd.notna(ultima_actividad['hora_salida']) and ultima_actividad['hora_salida'] != '' else 'en_curso'
    
    return {
        'total_actividades': len(actividades),
        'horas_total_trabajadas': total_horas,
        'actividades': actividades,
        'inicio_dia': '07:00',
        'estado_actual': estado_actual
    }

def registrar_entrada_salida(empleado, codigo_barras, servicio_info=None):
    """Función de compatibilidad - redirige a la nueva función de actividades continuas"""
    return registrar_actividad_continua(empleado, codigo_barras, servicio_info)

def pantalla_inicio():
    """Pantalla inicial de la aplicación con diseño Tekpro estilo tarjeta"""
    
    col1, col2, col3 = st.columns([1, 2, 1])
    
    with col2:
        st.markdown("""
        <div style='background: white; border-radius: 25px; overflow: hidden; box-shadow: 0 20px 60px rgba(62, 174, 165, 0.2); margin-top: 50px;'>
            <div style='height: 180px; background: linear-gradient(135deg, #2D8B84 0%, #3EAEA5 25%, #5BC4BC 50%, #7DD4CE 75%, #A8E6E1 100%); display: flex; flex-direction: column; align-items: center; justify-content: center; position: relative;'>
                <div style='font-family: Poppins, sans-serif; font-size: 24px; font-weight: 700; color: white; letter-spacing: 6px; text-shadow: 0 4px 20px rgba(0,0,0,0.8); margin-bottom: 10px; background: #3EAEA5; padding: 8px 25px; border-radius: 8px;'>TEKPRO</div>
                <div style='font-family: Poppins, sans-serif; font-size: 52px; font-weight: 700; color: white; letter-spacing: 12px; text-shadow: 0 6px 25px rgba(0,0,0,0.8); background: #2D8B84; padding: 15px 40px; border-radius: 12px;'>CHRONOTRACK</div>
            </div>
            <div style='padding: 40px; text-align: center; background: white;'>
                <div style='font-family: Poppins, sans-serif; font-size: 18px; font-weight: 600; color: #2D8B84; margin: 0 0 10px 0; letter-spacing: 3px;'>ChronoTrack</div>
                <h1 style='font-family: Poppins, sans-serif; font-size: 56px; font-weight: 700; color: #3EAEA5; margin: 0 0 40px 0; letter-spacing: 4px;'>CHRONOTRACK</h1>
                <div style='margin: 30px 0;'>
                    <div style='width: 100px; height: 100px; margin: 0 auto; border-radius: 50%; background: linear-gradient(135deg, #2D8B84 0%, #3EAEA5 50%, #5BC4BC 100%); display: flex; align-items: center; justify-content: center; font-size: 48px;'>⏱</div>
                </div>
                <p style='color: #6c757d; font-size: 14px; margin: 20px 0; font-family: Poppins, sans-serif;'>Escanea tu código de barras para comenzar</p>
            </div>
        </div>
        """, unsafe_allow_html=True)
        
        st.markdown("<br>", unsafe_allow_html=True)
        
        if st.button("INICIAR REGISTRO", type="primary", use_container_width=True):
            st.session_state.screen = 'registro_colaborador'
            st.rerun()
        
    
    # Botón oculto para acceso de administrador
    st.markdown("<br><br>", unsafe_allow_html=True)
    col1, col2, col3, col4, col5 = st.columns([1, 1, 1, 1, 1])
    with col3:
        if st.button("⚙", help="Acceso de administrador", key="admin_access"):
            st.session_state.screen = 'admin_login'
            st.rerun()

def pantalla_registro_colaborador():
    """Pantalla de registro paso a paso con códigos de barras - Diseño Tekpro estilo tarjeta"""
    
    # Header estilo tarjeta Tekpro
    st.markdown("""
    <div style='
        background: white;
        border-radius: 20px;
        overflow: hidden;
        box-shadow: 0 10px 40px rgba(62, 174, 165, 0.15);
        margin-bottom: 30px;
    '>
        <!-- Patrón de bloques pixelados superior (más pequeño) -->
        <div style='
            height: 100px;
            background: linear-gradient(135deg, #2D8B84 0%, #3EAEA5 25%, #5BC4BC 50%, #7DD4CE 75%);
            position: relative;
            overflow: hidden;
        '>
            <!-- Bloques pixelados simplificados -->
            <div style='position: absolute; top: 0; left: 0; width: 100%; height: 100%;'>
                <svg width="100%" height="100" xmlns="http://www.w3.org/2000/svg">
                    <rect x="0" y="0" width="50" height="50" fill="#2D8B84" opacity="0.7"/>
                    <rect x="50" y="0" width="50" height="50" fill="#3EAEA5" opacity="0.5"/>
                    <rect x="100" y="0" width="50" height="50" fill="#5BC4BC" opacity="0.6"/>
                    <rect x="150" y="0" width="50" height="50" fill="#7DD4CE" opacity="0.4"/>
                    
                    <rect x="0" y="50" width="50" height="50" fill="#3EAEA5" opacity="0.6"/>
                    <rect x="50" y="50" width="50" height="50" fill="#5BC4BC" opacity="0.7"/>
                    <rect x="100" y="50" width="50" height="50" fill="#7DD4CE" opacity="0.5"/>
                    <rect x="150" y="50" width="50" height="50" fill="white" opacity="0.2"/>
                </svg>
            </div>
            
            <!-- Logo y título -->
            <div style='position: relative; z-index: 1; text-align: center; padding-top: 25px;'>
                <h1 style='
                    font-family: Poppins, sans-serif;
                    font-size: 32px;
                    font-weight: 500;
                    color: white;
                    margin: 0;
                    letter-spacing: 1px;
                '>ChronoTrack</h1>
                <p style='
                    font-family: Poppins, sans-serif;
                    font-size: 14px;
                    color: rgba(255,255,255,0.9);
                    margin: 5px 0 0 0;
                '>Registro de Actividades</p>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    if st.button("← Volver", type="secondary"):
        st.session_state.screen = 'inicio'
        st.rerun()
    
    st.markdown("<br>", unsafe_allow_html=True)
    
    # Inicializar variables de sesión para el flujo paso a paso
    if 'step' not in st.session_state:
        st.session_state.step = 1
    if 'empleado_data' not in st.session_state:
        st.session_state.empleado_data = {}
    
    # Paso 1: Código de barras de la cédula
    if st.session_state.step == 1:
        mostrar_paso_cedula()
    # Paso 2: Código de barras de la actividad  
    elif st.session_state.step == 2:
        mostrar_paso_actividad()
    # Paso 3: Código de barras de la OP
    elif st.session_state.step == 3:
        mostrar_paso_op()
    # Paso 4: Confirmación y guardado (eliminamos descripción)
    elif st.session_state.step == 4:
        mostrar_confirmacion_guardado()

def mostrar_paso_cedula():
    """Paso 1: Escanear cédula del colaborador - Diseño Tekpro"""
    st.markdown("""
    <div class='codigo-barras-container pixel-pattern'>
        <h3 style='text-align: center; margin-bottom: 15px; color: #3EAEA5; font-weight: 600;'>
            Paso 1: Identificación del Colaborador
        </h3>
        <p style='text-align: center; color: #6c757d; font-size: 15px;'>
            Escanea el código de barras de tu cédula para comenzar
        </p>
    </div>
    """, unsafe_allow_html=True)
    
    # Instrucciones para lector USB mejoradas con diseño Tekpro
    st.markdown("""
    <div class='scanner-instructions'>
        <h4>Lector de Código de Barras</h4>
        <p>Haz clic en el campo y escanea tu identificación</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Detectar si es un dispositivo móvil y mostrar opción de cámara
    st.markdown("""
    <script>
    const isMobile = /Android|webOS|iPhone|iPad|iPod|BlackBerry|IEMobile|Opera Mini/i.test(navigator.userAgent);
    if (isMobile) {
        window.parent.postMessage({type: 'streamlit:setComponentValue', value: true}, '*');
    }
    </script>
    """, unsafe_allow_html=True)
    
    # Botón para abrir cámara en móviles
    col1, col2 = st.columns([1, 1])
    
    with col1:
        if st.button("📷 Escanear con Cámara", use_container_width=True, type="primary"):
            st.session_state['mostrar_camara'] = True
    
    with col2:
        if st.button("⌨️ Ingresar Manual", use_container_width=True):
            st.session_state['mostrar_camara'] = False
    
    # Inicializar estado
    if 'mostrar_camara' not in st.session_state:
        st.session_state['mostrar_camara'] = False
    
    codigo_barras = None
    
    # Mostrar interfaz de cámara o campo de texto según la selección
    if st.session_state.get('mostrar_camara', False):
        st.markdown("""
        <div style='background: linear-gradient(135deg, #3EAEA5 0%, #5BC4BC 100%); padding: 20px; border-radius: 15px; margin: 20px 0;'>
            <h3 style='color: white; text-align: center; margin-bottom: 15px;'>📷 Escáner de Código de Barras</h3>
            <p style='color: white; text-align: center;'>Coloca el código de barras frente a la cámara</p>
        </div>
        """, unsafe_allow_html=True)
        
        # Componente HTML5 para acceso a cámara y escaneo de códigos de barras
        components.html("""
        <!DOCTYPE html>
        <html>
        <head>
            <script src="https://unpkg.com/@zxing/library@latest"></script>
            <style>
                body { margin: 0; padding: 20px; background: #f0f2f6; font-family: Arial, sans-serif; }
                #video-container { position: relative; max-width: 100%; margin: 0 auto; }
                #video { width: 100%; max-height: 400px; border-radius: 10px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }
                #result { 
                    margin-top: 20px; 
                    padding: 15px; 
                    background: #28a745; 
                    color: white; 
                    border-radius: 10px; 
                    font-size: 18px;
                    text-align: center;
                    display: none;
                }
                #loading { text-align: center; color: #666; padding: 20px; }
                .error { background: #dc3545 !important; }
            </style>
        </head>
        <body>
            <div id="loading">Iniciando cámara...</div>
            <div id="video-container" style="display:none;">
                <video id="video" playsinline></video>
            </div>
            <div id="result"></div>
            
            <script>
                const codeReader = new ZXing.BrowserMultiFormatReader();
                const videoElement = document.getElementById('video');
                const resultElement = document.getElementById('result');
                const loadingElement = document.getElementById('loading');
                const videoContainer = document.getElementById('video-container');
                
                // Iniciar escaneo
                codeReader.decodeFromVideoDevice(null, videoElement, (result, err) => {
                    if (result) {
                        const codigo = result.text;
                        resultElement.textContent = '✅ Código detectado: ' + codigo;
                        resultElement.style.display = 'block';
                        resultElement.classList.remove('error');
                        
                        // Enviar código a Streamlit
                        window.parent.postMessage({
                            type: 'streamlit:setComponentValue',
                            value: codigo
                        }, '*');
                        
                        // Detener escaneo después de detectar
                        setTimeout(() => {
                            codeReader.reset();
                        }, 2000);
                    }
                    
                    if (err && !(err instanceof ZXing.NotFoundException)) {
                        console.error(err);
                    }
                }).then(() => {
                    loadingElement.style.display = 'none';
                    videoContainer.style.display = 'block';
                }).catch(err => {
                    loadingElement.textContent = '❌ Error al acceder a la cámara. Por favor, da permisos de cámara.';
                    loadingElement.style.color = '#dc3545';
                    console.error(err);
                });
            </script>
        </body>
        </html>
        """, height=550)
        
        # Campo oculto para recibir el código escaneado
        codigo_barras = st.text_input(
            "Código escaneado:",
            key="codigo_camara",
            label_visibility="collapsed"
        )
    else:
        # Campo optimizado para lectores USB
        st.markdown('<div class="barcode-scanner-field">', unsafe_allow_html=True)
        codigo_barras = st.text_input(
            "🔍 Código de barras:",
            placeholder="● ● ● Campo listo para escanear ● ● ●",
            key="codigo_input",
            help="✅ Optimizado para lectores USB\n🔍 El código aparecerá automáticamente\n⚡ Procesamiento instantáneo",
            label_visibility="collapsed"
        )
        st.markdown('</div>', unsafe_allow_html=True)
    
    # Auto-enfoque para escáner
    components.html("""
    <script>
    function focusInput() {
        // Buscar el input por múltiples criterios
        var selectors = [
            'input[data-testid="textInput-codigo_input"]',
            'input[placeholder*="Campo listo para escanear"]',
            'input[aria-label*="Código de barras"]',
            'input[key="codigo_input"]'
        ];
        
        for (var i = 0; i < selectors.length; i++) {
            var input = parent.document.querySelector(selectors[i]);
            if (input) {
                input.focus();
                input.select();
                break;
            }
        }
    }
    
    // Intentar enfocar inmediatamente y luego repetir
    focusInput();
    setTimeout(focusInput, 100);
    setTimeout(focusInput, 500);
    </script>
    """, height=0)
    
    # JavaScript para auto-focus y detección de escaneo
    st.markdown("""
    <script>
    let lastScanTime = 0;
    let scanBuffer = '';
    
    function focusBarcodeField() {
        const inputs = window.parent.document.querySelectorAll('input[aria-label*="Código de barras"]');
        if (inputs.length > 0) {
            inputs[0].focus();
            inputs[0].select();
            return inputs[0];
        }
        return null;
    }
    
    function detectScan() {
        const input = focusBarcodeField();
        if (input && !input.hasEventListener) {
            input.addEventListener('input', function(e) {
                const currentTime = Date.now();
                const timeDiff = currentTime - lastScanTime;
                
                // Si el tiempo entre caracteres es muy corto, probablemente es un escáner
                if (timeDiff < 50 && e.target.value.length > 3) {
                    // Marcar visualmente que se detectó un escaneo
                    e.target.style.backgroundColor = '#d4edda';
                    e.target.style.borderColor = '#28a745';
                }
                
                lastScanTime = currentTime;
            });
            
            input.hasEventListener = true;
        }
    }
    
    // Auto-focus inicial
    setTimeout(focusBarcodeField, 500);
    setTimeout(detectScan, 600);
    
    // Mantener el campo activo
    setInterval(() => {
        focusBarcodeField();
        detectScan();
    }, 2000);
    </script>
    """, unsafe_allow_html=True)
    
    # Selector de servicio en un expander
    servicio_seleccionado = None
        
    if codigo_barras:
        # Validar formato del código de barras
        es_valido, tipo_codigo = validar_codigo_barras(codigo_barras)
        
        if es_valido:
            st.success(f"**Código detectado:** {tipo_codigo}")
            
            with st.spinner("Buscando colaborador..."):
                empleado, mensaje_gs = buscar_colaborador_en_datos_colab(codigo_barras)
            
            if empleado is None:
                empleado = buscar_empleado_por_codigo(codigo_barras)
                if empleado:
                    mensaje_gs = "Colaborador encontrado en configuración local"
            
            if empleado:
                # 🔍 BUSCAR ÚLTIMO REGISTRO EN GOOGLE SHEETS
                with st.spinner("📊 Analizando último registro en Google Sheets..."):
                    ultimo_registro_sheets, mensaje_registro = obtener_ultimo_registro_sheets(codigo_barras)
                
                hora_exacta_anterior = None
                diferencia_horas = None
                
                if ultimo_registro_sheets:
                    hora_exacta_anterior = ultimo_registro_sheets['hora_exacta']
                    
                    # Calcular diferencia entre hora actual y hora_exacta anterior
                    try:
                        hora_actual = obtener_hora_colombia()
                        # Parsear hora_exacta anterior (formato esperado: HH:MM:SS)
                        if ':' in hora_exacta_anterior:
                            partes = hora_exacta_anterior.split(':')
                            if len(partes) >= 2:
                                hora_anterior_time = time(int(partes[0]), int(partes[1]), int(partes[2]) if len(partes) > 2 else 0)
                                # Convertir a datetime para calcular diferencia
                                hora_anterior_dt = datetime.combine(date.today(), hora_anterior_time)
                                
                                # Calcular diferencia en horas
                                diferencia = hora_actual - hora_anterior_dt
                                diferencia_horas = diferencia.total_seconds() / 3600
                                
                                # Si la diferencia es negativa, significa que el registro anterior fue ayer
                                if diferencia_horas < 0:
                                    diferencia_horas += 24  # Agregar 24 horas
                    except Exception as e:
                        st.warning(f"No se pudo calcular la diferencia de tiempo: {str(e)}")
                
                st.markdown(f"""
                <div class='success-message'>
                    <strong>Colaborador encontrado:</strong> {empleado}<br>
                    <small>{mensaje_gs}</small>
                </div>
                """, unsafe_allow_html=True)
                
                # 🏠 VERIFICAR SI ES HORARIO DE ADECUACIÓN LOCATIVA (4:20 PM - 4:30 PM)
                es_adecuacion, info_adecuacion = es_horario_adecuacion_locativa()
                if es_adecuacion and info_adecuacion:
                    st.markdown(f"""
                    <div style='background: linear-gradient(135deg, #E8F5E9 0%, #C8E6C9 100%); 
                                padding: 15px; border-radius: 12px; border-left: 5px solid #4CAF50; margin: 10px 0;'>
                        <strong>🏠 Horario de Adecuación Locativa</strong><br>
                        <span style='font-size: 16px; color: #2E7D32;'>
                            Tu registro se guardará como: <strong>{info_adecuacion['hora_registro']}</strong>
                        </span><br>
                        <span style='font-size: 14px; color: #388E3C;'>
                            Actividad: <strong>{info_adecuacion['servicio_nombre']}</strong>
                        </span><br>
                        <small style='color: #666;'>Los registros entre 4:20 PM y 4:30 PM se asignan automáticamente a esta actividad</small>
                    </div>
                    """, unsafe_allow_html=True)
                
                # Mostrar información del último registro si existe
                if hora_exacta_anterior and diferencia_horas is not None:
                    st.markdown(f"""
                    <div class='info-message'>
                        <strong>Último Registro Encontrado</strong><br>
                        <strong>Hora exacta anterior:</strong> {hora_exacta_anterior}<br>
                        <strong>Tiempo transcurrido:</strong> {diferencia_horas:.2f} horas ({int(diferencia_horas * 60)} minutos)<br>
                        <small>{mensaje_registro}</small>
                    </div>
                    """, unsafe_allow_html=True)
                    
                    # GUARDAR LA DIFERENCIA EN SESSION_STATE PARA USAR EN EL GUARDADO
                    st.session_state.tiempo_calculado = round(diferencia_horas, 2)
                else:
                    st.info(f"{mensaje_registro}")
                    st.session_state.tiempo_calculado = 0  # Sin registro previo, tiempo es 0
                
                # Incluir información del servicio si está seleccionado
                servicio_info = servicio_seleccionado['display'] if servicio_seleccionado else None
                tipo_registro, horas_totales, analisis = registrar_entrada_salida(empleado, codigo_barras, servicio_info)
                
                if tipo_registro == 'actividad_registrada':
                    # Determinar estilo según el análisis
                    if not analisis.get('es_dia_laboral'):
                        mensaje_class = 'info-message'
                        estado_texto = 'Actividad registrada (Día no laboral)'
                    else:
                        mensaje_class = 'success-message'
                        estado_texto = 'Nueva actividad registrada'
                    
                    # Obtener información adicional sobre el día
                    df_actual = load_data()
                    registros_hoy = df_actual[
                        (df_actual['empleado'] == empleado) & 
                        (df_actual['fecha'] == obtener_fecha_colombia())
                    ]
                    
                    actividad_html = f"""
                    <div class='{mensaje_class}'>
                        <strong>{estado_texto}</strong><br>
                        Empleado: {empleado}<br>
                        Fecha: {obtener_fecha_colombia().strftime('%d/%m/%Y')}<br>
                        Hora registro: {obtener_hora_colombia().strftime('%H:%M:%S')}<br>
                        <strong>Inicio del día: 07:00 AM</strong><br>
                        <strong>Horas transcurridas hoy: {horas_totales:.2f}h</strong><br>
                        <strong>Total actividades: {len(registros_hoy)}</strong>
                    """
                    
                    if servicio_seleccionado:
                        actividad_html += f"<br>🔧 Servicio: {servicio_seleccionado['display']}"
                    
                    # Agregar información del horario laboral
                    if analisis.get('es_dia_laboral'):
                        actividad_html += f"<br>📅 Horario esperado: {analisis['hora_entrada_esperada'].strftime('%H:%M')}"
                        if analisis.get('entrada_mensaje'):
                            actividad_html += f"<br><small>ℹ️ {analisis['entrada_mensaje']}</small>"
                    
                    # Mostrar actividades anteriores del día
                    if len(registros_hoy) > 0:
                        actividad_html += "<br><br><strong>📊 Resumen del día:</strong>"
                        for i, registro in enumerate(registros_hoy.iterrows()):
                            reg = registro[1]
                            if pd.notna(reg['horas_trabajadas']) and reg['horas_trabajadas'] != '':
                                actividad_html += f"<br>• Actividad {i+1}: {reg['horas_trabajadas']:.1f}h"
                                if reg['servicio']:
                                    actividad_html += f" - {reg['servicio']}"
                    
                    actividad_html += "</div>"
                    
                    st.markdown(actividad_html, unsafe_allow_html=True)
                
                
                
            else:
                st.markdown(f"""
                <div class='error-message'>
                    <strong>❌ Colaborador no encontrado</strong><br>
                    Código: {codigo_barras}<br>
                    {mensaje_gs}
                </div>
                """, unsafe_allow_html=True)
                
                st.info("💡 **Sugerencias:**\n- Verifica que el código sea correcto\n- Contacta al administrador para registrar tu código")
        else:
            # Código no válido
            st.error(f"❌ **Formato de código inválido:** {tipo_codigo}")
            st.warning("⚠️ **Posibles soluciones:**\n- Asegúrate de escanear completamente el código\n- Verifica que el lector esté configurado correctamente\n- Intenta escanear nuevamente")
        
    # Procesar código de barras si se ingresó
    if codigo_barras:
        # Validar y buscar empleado
        empleado, mensaje_gs = buscar_colaborador_en_datos_colab(codigo_barras)
        
        if empleado:
            # Guardar datos del empleado y avanzar al siguiente paso
            st.session_state.empleado_data = {
                'cedula': codigo_barras,
                'nombre': empleado
            }
            st.session_state.step = 2
            st.rerun()
        else:
            st.error(f"❌ Colaborador no encontrado: {codigo_barras}")

def mostrar_paso_actividad():
    """Paso 2: Escanear código de actividad"""
    empleado_data = st.session_state.empleado_data
    
    # Mostrar saludo personalizado
    st.markdown(f"""
    <div class='success-message'>
        <h3 style='margin: 0; color: #155724;'>👋 ¡Hola, {empleado_data['nombre']}!</h3>
        <p style='margin: 5px 0 0 0;'>Continuemos con el registro de tu actividad</p>
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown("""
    <div class='codigo-barras-container'>
        <h3 style='text-align: center; margin-bottom: 20px; color: #1f77b4;'>
            🔧 Paso 2: Código de Actividad
        </h3>
        <p style='text-align: center; color: #666;'>Escanea el código de barras de la actividad a realizar</p>
    </div>
    """, unsafe_allow_html=True)
    
    codigo_actividad = st.text_input(
        "🔧 Código de actividad:",
        placeholder="● ● ● Escanea código de actividad ● ● ●",
        key="codigo_actividad_input"
    )
    
    # Auto-enfoque para escáner
    components.html("""
    <script>
    function focusInput() {
        var selectors = [
            'input[data-testid="textInput-codigo_actividad_input"]',
            'input[placeholder*="Escanea código de actividad"]',
            'input[aria-label*="Código de actividad"]'
        ];
        
        for (var i = 0; i < selectors.length; i++) {
            var input = parent.document.querySelector(selectors[i]);
            if (input) {
                input.focus();
                input.select();
                break;
            }
        }
    }
    
    focusInput();
    setTimeout(focusInput, 100);
    setTimeout(focusInput, 500);
    </script>
    """, height=0)
    
    # Mostrar información del servicio encontrado
    if 'servicio_info' in st.session_state.empleado_data:
        servicio_info = st.session_state.empleado_data['servicio_info']
        st.markdown(f"""
        <div class='success-message'>
            <h4 style='margin: 0;'>✅ Servicio Encontrado</h4>
            <p><strong>Código:</strong> {servicio_info['numero']}</p>
            <p><strong>Servicio:</strong> {servicio_info['nomservicio']}</p>
        </div>
        """, unsafe_allow_html=True)
    
    col1, col2 = st.columns(2)
    with col1:
        if st.button("← Anterior", type="secondary"):
            st.session_state.step = 1
            st.rerun()
    
    if codigo_actividad:
        with st.spinner("🔍 Buscando servicio..."):
            numero, nomservicio, mensaje = buscar_servicio_por_codigo(codigo_actividad)
        
        if numero and nomservicio:
            # Guardar información del servicio encontrado
            st.session_state.empleado_data['codigo_actividad'] = codigo_actividad
            st.session_state.empleado_data['servicio_info'] = {
                'numero': numero,
                'nomservicio': nomservicio
            }
            
            # Verificar si el código del servicio está entre 15 y 31
            try:
                codigo_num = int(numero)
                if 15 <= codigo_num <= 31:
                    # Servicios del 15 al 31: OP automática "0000"
                    st.session_state.empleado_data['codigo_op'] = '0000'
                    st.session_state.empleado_data['op_info'] = {
                        'orden': '0000',
                        'referencia': 'N/A',
                        'cantidades': 'N/A',
                        'cliente': 'N/A',
                        'item': 'Servicio directo'
                    }
                    
                    # Mostrar mensaje especial y saltar directo al paso 4
                    st.success(f"✅ Servicio encontrado: {numero} - {nomservicio}")
                    st.info("🔧 Servicio directo - OP automática: 0000")
                    
                    # Esperar un momento y saltar al paso 4 (confirmación)
                    import time
                    time.sleep(1)
                    st.session_state.step = 4
                    st.rerun()
                else:
                    # Servicios normales: continuar con el paso de OP
                    st.success(f"✅ Servicio encontrado: {numero} - {nomservicio}")
                    
                    # Esperar un momento para mostrar el mensaje y luego avanzar al paso 3
                    import time
                    time.sleep(1)
                    st.session_state.step = 3
                    st.rerun()
            except ValueError:
                # Si no es un número, continuar normalmente
                st.success(f"✅ Servicio encontrado: {numero} - {nomservicio}")
                
                # Esperar un momento para mostrar el mensaje y luego avanzar
                import time
                time.sleep(1)
                st.session_state.step = 3
                st.rerun()
        else:
            st.error(f"❌ {mensaje}")
            st.warning("💡 Verifica que el código de actividad esté registrado en la hoja 'Servicio'")
            
            # Mostrar información adicional de debug
            with st.expander("🔧 Información de Debug"):
                st.write(f"**Código buscado:** '{codigo_actividad}'")
                st.write("**Verificando estructura de la hoja 'Servicio'...**")
                
                # Llamar función de verificación
                resultado_verificacion = verificar_estructura_servicio()
                st.text(resultado_verificacion)
            
            # Mostrar opción para continuar sin servicio
            if st.button("Continuar sin servicio →", type="secondary"):
                st.session_state.empleado_data['codigo_actividad'] = codigo_actividad
                st.session_state.empleado_data['servicio_info'] = {
                    'numero': codigo_actividad,
                    'nomservicio': 'Servicio no encontrado'
                }
                st.session_state.step = 3
                st.rerun()

def mostrar_paso_op():
    """Paso 3: Escanear código de OP (Orden de Producción)"""
    empleado_data = st.session_state.empleado_data
    
    st.markdown(f"""
    <div class='info-message'>
        <h4>📋 Colaborador: {empleado_data['nombre']}</h4>
        <p>🔧 Actividad: {empleado_data.get('codigo_actividad', 'N/A')}</p>
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown("""
    <div class='codigo-barras-container'>
        <h3 style='text-align: center; margin-bottom: 20px; color: #1f77b4;'>
            📋 Paso 3: Código de OP
        </h3>
        <p style='text-align: center; color: #666;'>Escanea el código de barras de la Orden de Producción</p>
    </div>
    """, unsafe_allow_html=True)
    
    codigo_op = st.text_input(
        "📋 Código de OP:",
        placeholder="● ● ● Escanea código de OP ● ● ●",
        key="codigo_op_input"
    )
    
    # Auto-enfoque para escáner
    components.html("""
    <script>
    function focusInput() {
        var selectors = [
            'input[data-testid="textInput-codigo_op_input"]',
            'input[placeholder*="Escanea código de OP"]',
            'input[aria-label*="Código de OP"]'
        ];
        
        for (var i = 0; i < selectors.length; i++) {
            var input = parent.document.querySelector(selectors[i]);
            if (input) {
                input.focus();
                input.select();
                break;
            }
        }
    }
    
    focusInput();
    setTimeout(focusInput, 100);
    setTimeout(focusInput, 500);
    </script>
    """, height=0)
    
    # Mostrar información de la OP encontrada
    if 'op_info' in st.session_state.empleado_data:
        op_info = st.session_state.empleado_data['op_info']
        st.markdown(f"""
        <div class='success-message'>
            <h4 style='margin: 0;'>✅ OP Encontrada</h4>
            <p><strong>📋 Orden:</strong> {op_info['orden']}</p>
            <p><strong>🏷️ Referencia:</strong> {op_info['referencia']}</p>
            <p><strong>📦 Cantidades:</strong> {op_info['cantidades']}</p>
            <p><strong>👤 Cliente:</strong> {op_info['cliente']}</p>
            <p><strong>📝 Item:</strong> {op_info['item']}</p>
        </div>
        """, unsafe_allow_html=True)
    
    col1, col2 = st.columns(2)
    with col1:
        if st.button("← Anterior", type="secondary"):
            st.session_state.step = 2
            st.rerun()
    
    if codigo_op:
        with st.spinner("🔍 Buscando información de la OP..."):
            op_info, mensaje = buscar_op_por_codigo(codigo_op)
        
        if op_info:
            # Guardar información completa de la OP
            st.session_state.empleado_data['codigo_op'] = codigo_op
            st.session_state.empleado_data['op_info'] = op_info
            
            # Mostrar mensaje de éxito y avanzar
            st.success(f"✅ OP encontrada: {op_info['orden']} - {op_info['cliente']}")
            
            # Esperar un momento para mostrar el mensaje y luego avanzar
            import time
            time.sleep(1)
            st.session_state.step = 4
            st.rerun()
        else:
            st.error(f"❌ {mensaje}")
            st.warning("💡 Verifica que el código de OP esté registrado en la hoja 'OPS'")
            
            # Mostrar opción para continuar sin OP
            if st.button("Continuar sin OP →", type="secondary"):
                st.session_state.empleado_data['codigo_op'] = codigo_op
                st.session_state.empleado_data['op_info'] = {
                    'orden': codigo_op,
                    'referencia': 'N/A',
                    'cantidades': 'N/A',
                    'cliente': 'No encontrado',
                    'item': 'N/A'
                }
                st.session_state.step = 4
                st.rerun()

def mostrar_confirmacion_guardado():
    """Paso 5: Confirmación y guardado final"""
    empleado_data = st.session_state.empleado_data
    
    st.markdown("""
    <div class='success-message'>
        <h3 style='text-align: center; margin-bottom: 20px;'>✅ Confirmación de Registro</h3>
    </div>
    """, unsafe_allow_html=True)
    
    # Obtener información completa del servicio y OP
    servicio_info = empleado_data.get('servicio_info', {})
    op_info = empleado_data.get('op_info', {})
    

    
    # Calcular tiempo de la actividad anterior basado en cédula
    df = load_data()
    fecha_actual = obtener_fecha_colombia()
    hora_actual = obtener_hora_colombia_time()
    
    # Debug: mostrar información del DataFrame
    print(f"DEBUG Confirmación: DataFrame shape: {df.shape}")
    print(f"DEBUG Confirmación: Columnas disponibles: {df.columns.tolist()}")
    print(f"DEBUG Confirmación: Buscando cédula: {empleado_data['cedula']}")
    print(f"DEBUG Confirmación: Fecha actual: {fecha_actual}")
    
    ultimo_registro = obtener_ultimo_registro_por_cedula(empleado_data['cedula'], fecha_actual, df)
    
    tiempo_actividad_anterior = 0
    if ultimo_registro is not None:
        hora_entrada_anterior = ultimo_registro['hora_entrada']
        if isinstance(hora_entrada_anterior, str):
            hora_entrada_anterior = datetime.strptime(hora_entrada_anterior, '%H:%M:%S').time()
        tiempo_actividad_anterior = calcular_horas(hora_entrada_anterior, hora_actual)
    
    # Calcular los valores para mostrar en la confirmación
    orden_completa = op_info.get('orden', '')
    orden_mostrar = orden_completa if orden_completa else 'N/A'  # La orden es el OP tal como está
    
    # El item es el número después del último guión o viene de la OP
    item_desde_op = op_info.get('item', '')
    if item_desde_op:
        item_mostrar = item_desde_op
    elif '-' in orden_completa:
        partes = orden_completa.split('-')
        item_mostrar = partes[-1] if len(partes) > 1 else "1"
    else:
        item_mostrar = "1"
    

    # Mostrar información de la OP y tiempo
    if ultimo_registro is not None:
        # Formatear el tiempo de manera más legible
        horas = int(tiempo_actividad_anterior)
        minutos = int((tiempo_actividad_anterior - horas) * 60)
        if horas > 0:
            tiempo_texto = f"{horas}h {minutos}min"
        else:
            tiempo_texto = f"{minutos} minutos"
    else:
        tiempo_texto = "Primera actividad del día"
    
    # Debug temporal: mostrar información en la interfaz
    with st.expander("🔍 Debug - ¿Por qué aparece 'Primer registro del día'?"):
        st.write(f"**DataFrame shape:** {df.shape}")
        st.write(f"**Columnas disponibles:** {df.columns.tolist()}")
        st.write(f"**Cédula buscada:** {empleado_data['cedula']}")
        st.write(f"**Fecha actual:** {fecha_actual}")
        st.write(f"**Último registro encontrado:** {ultimo_registro is not None}")
        
        if not df.empty:
            st.write("**Primeros registros del DataFrame:**")
            st.dataframe(df.head())
            
            if 'cedula' in df.columns:
                st.write(f"**Cédulas únicas en datos:** {df['cedula'].unique()}")
            if 'fecha' in df.columns:
                st.write(f"**Fechas únicas en datos:** {df['fecha'].unique()}")
        else:
            st.write("**El DataFrame está vacío**")
    
    st.markdown(f"""
    <div class='info-message'>
        <p><strong>📋 Orden de Producción:</strong> {orden_mostrar}</p>
        <p><strong>🏭 Cliente:</strong> {op_info.get('cliente', 'N/A')}</p>
        <p><strong>📦 Referencia:</strong> {op_info.get('referencia', 'N/A')}</p>
        <p><strong>🔢 Item:</strong> {item_mostrar}</p>
        <p><strong>⏰ Tiempo de la Actividad Completada:</strong> {tiempo_texto}</p>
        <p><strong> Fecha:</strong> {obtener_hora_colombia().strftime('%d/%m/%Y %H:%M:%S')}</p>
    </div>
    """, unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        if st.button("← Modificar", type="secondary"):
            st.session_state.step = 3
            st.rerun()
    
    with col2:
        if st.button("🚫 Cancelar", type="secondary"):
            # Limpiar datos y volver al inicio
            st.session_state.step = 1
            st.session_state.empleado_data = {}
            st.rerun()
    
    with col3:
        if st.button("💾 Guardar Registro", type="primary"):
            # Guardar el registro completo
            guardar_registro_completo(empleado_data)
            # Mostrar confirmación y reiniciar
            st.success("✅ ¡Registro guardado exitosamente!")
            
            
            # Reiniciar para nuevo registro
            st.session_state.step = 1
            st.session_state.empleado_data = {}
            
            # Auto-reload después de 3 segundos
            st.markdown("""
            <script>
            setTimeout(function() {
                window.location.reload();
            }, 3000);
            </script>
            """, unsafe_allow_html=True)

def finalizar_actividad_actual(empleado, hora_finalizacion=None):
    """Finalizar la actividad actual de un empleado"""
    if hora_finalizacion is None:
        hora_finalizacion = obtener_hora_colombia_time()
    
    df = load_data()
    fecha_actual = obtener_fecha_colombia()
    
    # Buscar el último registro sin salida
    ultimo_registro = obtener_ultimo_registro_del_dia(empleado, fecha_actual, df)

def finalizar_actividad_por_cedula(cedula, hora_finalizacion=None):
    """Finalizar la actividad actual de un empleado basado en su cédula"""
    if hora_finalizacion is None:
        hora_finalizacion = obtener_hora_colombia_time()
    
    df = load_data()
    fecha_actual = obtener_fecha_colombia()
    
    # Buscar el último registro sin salida basado en cédula
    ultimo_registro = obtener_ultimo_registro_por_cedula(cedula, fecha_actual, df)
    
    if ultimo_registro is not None and (pd.isna(ultimo_registro['hora_salida']) or ultimo_registro['hora_salida'] == ''):
        # Finalizar el registro
        idx = ultimo_registro.name
        df.loc[idx, 'hora_salida'] = hora_finalizacion.strftime('%H:%M:%S') if hasattr(hora_finalizacion, 'strftime') else str(hora_finalizacion)
        
        # Calcular horas trabajadas
        hora_entrada = ultimo_registro['hora_entrada']
        if isinstance(hora_entrada, str):
            hora_entrada = datetime.strptime(hora_entrada, '%H:%M:%S').time()
        
        horas_trabajadas = calcular_horas(hora_entrada, hora_finalizacion)
        df.loc[idx, 'horas_trabajadas'] = round(horas_trabajadas, 2)
        
        # Guardar cambios
        save_data(df)
        
        return True, horas_trabajadas
    
    return False, 0

def guardar_registro_completo(empleado_data):
    """Guardar registro usando la nueva lógica de conteos diarios"""
    df = load_data()
    fecha_actual = obtener_fecha_colombia()
    hora_actual = obtener_hora_colombia_time()
    empleado = empleado_data['nombre']
    cedula = empleado_data['cedula']
    
    # ============================================
    # VERIFICAR ADECUACIÓN LOCATIVA (4:20 PM - 4:30 PM)
    # ============================================
    es_adecuacion, info_adecuacion = es_horario_adecuacion_locativa()
    
    if es_adecuacion and info_adecuacion:
        # Aplicar hora fija de 16:30 y servicio de Adecuación Locativa
        hora_actual = datetime.strptime(info_adecuacion['hora_registro'], '%H:%M').time()
        
        # Sobrescribir el servicio con Adecuación Locativa
        servicio_adecuacion = obtener_servicio_adecuacion_locativa()
        empleado_data['servicio_info'] = {
            'numero': servicio_adecuacion['numero'],
            'nomservicio': servicio_adecuacion['nomservicio']
        }
        empleado_data['codigo_actividad'] = servicio_adecuacion['numero']
        
        # Establecer OP a 0000 para adecuación locativa
        empleado_data['op_info'] = {
            'orden': '0000',
            'referencia': 'N/A',
            'cantidades': 'N/A',
            'cliente': 'N/A',
            'item': 'Adecuación Locativa'
        }
        empleado_data['codigo_op'] = '0000'
        
        st.info(f"🏠 **Adecuación Locativa aplicada:** Hora ajustada a {info_adecuacion['hora_registro']} - {info_adecuacion['servicio_nombre']}")
    
    # Obtener información completa del servicio y OP
    servicio_info = empleado_data.get('servicio_info', {})
    op_info = empleado_data.get('op_info', {})
    servicio_display = f"{servicio_info.get('numero', '')} - {servicio_info.get('nomservicio', '')}" if servicio_info else ''
    
    # NUEVA LÓGICA DE CONTEOS DIARIOS
    st.info("🕐 Aplicando nueva lógica de conteos diarios...")
    
    # IMPORTANTE: Recargar el DataFrame más actual antes del cálculo
    df_actualizado = load_data()
    
    # Calcular usando la nueva lógica con DataFrame actualizado
    # Si es adecuación locativa, pasar la hora forzada
    hora_para_calculo = hora_actual if es_adecuacion else None
    conteo_resultado = calcular_horas_conteo_diario(cedula, fecha_actual, hora_actual, hora_para_calculo)
    
    # Debug: Mostrar información de la nueva lógica
    with st.expander("🕐 Debug - Nueva Lógica de Conteos"):
        if conteo_resultado['es_primer_registro']:
            st.success("✨ **PRIMER REGISTRO ABSOLUTO DE ESTA CÉDULA**")
            st.write(f"🕐 **Desde las 7:00 AM hasta:** {conteo_resultado['hora_fin_conteo']}")
            st.write(f"⏱️ **Tiempo trabajado:** {conteo_resultado['tiempo_trabajado']} horas")
        else:
            tipo_registro = "PRIMER REGISTRO DEL DÍA" if conteo_resultado.get('es_primer_registro_del_dia', False) else "REGISTRO ADICIONAL"
            st.info(f"🔄 **{tipo_registro} - BASADO EN ÚLTIMA HORA**")
            st.write("**Tiempo calculado:**")
            st.write(f"  - Desde última hora registrada: {conteo_resultado['hora_inicio_conteo']}")
            st.write(f"  - Hasta hora actual: {conteo_resultado['hora_fin_conteo']}")
            st.write(f"  - Tiempo trabajado: {conteo_resultado['tiempo_trabajado']} horas")
    
    # Determinar hora de inicio para el nuevo registro
    hora_inicio_nueva_actividad = conteo_resultado['hora_inicio_conteo']
    
    # PASO 2: Crear el nuevo registro según especificaciones
    # La Orden es simplemente el valor de la orden tal como está en "OPS"
    orden_completa = op_info.get('orden', '')
    
    # El item viene directamente de la OP o se extrae del formato anterior
    item_desde_op = op_info.get('item', '')
    if item_desde_op:
        item_formateado = item_desde_op
    elif '-' in orden_completa:
        partes = orden_completa.split('-')
        item_formateado = partes[-1] if len(partes) > 1 else "1"
    else:
        item_formateado = "1"
    
    # Formatear hora_fin_conteo como string si es objeto time
    hora_fin_str = ''
    if conteo_resultado['hora_fin_conteo']:
        if hasattr(conteo_resultado['hora_fin_conteo'], 'strftime'):
            hora_fin_str = conteo_resultado['hora_fin_conteo'].strftime('%H:%M:%S')
        else:
            hora_fin_str = str(conteo_resultado['hora_fin_conteo'])
    
    nuevo_registro = {
        'fecha': fecha_actual,
        'cedula': cedula,
        'empleado': empleado,
        'hora_entrada': hora_inicio_nueva_actividad,
        'codigo_actividad': empleado_data.get('codigo_actividad', ''),
        'op': orden_completa,
        'codigo_producto': str(op_info.get('referencia', '')).strip(),
        'cantidades': str(op_info.get('cantidades', '')).strip(),
        'nombre_cliente': str(op_info.get('cliente', '')).strip(),
        'descripcion_op': str(item_formateado).strip(),
        'descripcion_proceso': 'Produccion',
        'hora_salida': hora_fin_str,  # Siempre guardar la hora de fin del conteo
        'horas_trabajadas': conteo_resultado['tiempo_trabajado'],  # Tiempo trabajado calculado
        'hora_exacta': conteo_resultado['hora_exacta_registro'],  # Hora exacta del registro calculada
        'mes': fecha_actual.strftime('%m'),
        'año': fecha_actual.strftime('%Y'),
        'semana': str(fecha_actual.isocalendar()[1]),
        'referencia': str(op_info.get('referencia', '')).strip(),
        'servicio': f"{str(servicio_info.get('numero', '')).strip()} - {str(servicio_info.get('nomservicio', '')).strip()}" if servicio_info and servicio_info.get('numero') and servicio_info.get('nomservicio') else ''
    }
    
    # PASO 5: Guardar en archivo local INMEDIATAMENTE
    df = pd.concat([df, pd.DataFrame([nuevo_registro])], ignore_index=True)
    save_data(df)
    
    # PASO 6: Guardar registro en Google Sheets
    config = load_config()
    gs_enabled = config.get('google_sheets', {}).get('enabled', False)
    
    # Guardar registro actual en Google Sheets (tanto primer registro como siguientes)
    if gs_enabled:
        try:
            mensaje_guardado = "🔄 Guardando primer registro del día en Google Sheets..." if conteo_resultado['es_primer_registro'] else "🔄 Guardando nueva actividad en Google Sheets..."
            st.info(mensaje_guardado)
            registro_actual_para_sheets = {
                'fecha': fecha_actual,
                'cedula': cedula,
                'empleado': empleado,
                'codigo_actividad': empleado_data.get('codigo_actividad', ''),
                'op': orden_completa,
                'codigo_producto': str(op_info.get('referencia', '')).strip(),
                'cantidades': str(op_info.get('cantidades', '')).strip(),
                'nombre_cliente': str(op_info.get('cliente', '')).strip(),
                'descripcion_op': str(item_formateado).strip(),
                'descripcion_proceso': 'Produccion',
                'hora_entrada': conteo_resultado['hora_inicio_conteo'],
                'hora_salida': conteo_resultado['hora_fin_conteo'],
                'tiempo_horas': conteo_resultado['tiempo_trabajado'],
                'hora_exacta': conteo_resultado['hora_exacta_registro'],  # Hora exacta del registro calculada
                'mes': fecha_actual.strftime('%m'),
                'año': fecha_actual.strftime('%Y'),
                'semana': str(fecha_actual.isocalendar()[1]),
                'referencia': str(op_info.get('referencia', '')).strip(),
                'servicio': f"{str(servicio_info.get('numero', '')).strip()} - {str(servicio_info.get('nomservicio', '')).strip()}" if servicio_info and servicio_info.get('numero') and servicio_info.get('nomservicio') else ''
            }
            
            guardar_en_google_sheets_simple(registro_actual_para_sheets)
            if conteo_resultado['es_primer_registro']:
                st.success(f"✅ Primer registro del día guardado - Tiempo: {conteo_resultado['tiempo_trabajado']:.2f} horas")
            else:
                st.success(f"✅ Nueva actividad guardada - Tiempo inicial: {conteo_resultado['tiempo_trabajado']:.3f} horas")
        except Exception as e:
            tipo_registro = "primer registro" if conteo_resultado['es_primer_registro'] else "nueva actividad"
            st.error(f"❌ Error guardando {tipo_registro}: {str(e)}")
            st.info("📝 Datos guardados solo localmente")
    
    # PASO 7: Mostrar resumen del día y nueva actividad
    resumen_dia = obtener_resumen_dia_empleado(cedula, fecha_actual)
    
    with st.expander("� Resumen del día actual"):
        st.write(f"**Total de registros del día:** {resumen_dia['total_registros']}")
        
        # Mostrar tiempo total trabajado (horas reales)
        tiempo_total = resumen_dia['tiempo_total_trabajado']
        
        if tiempo_total == 0:
            st.info("⏰ **Tiempo total trabajado:** 0.00 horas (primera actividad del día)")
        else:
            st.info(f"⏰ **Tiempo total trabajado:** {tiempo_total:.2f} horas (tiempo real registrado)")
        
        if resumen_dia['registros_detalle']:
            st.write("**Detalle de registros:**")
            for reg in resumen_dia['registros_detalle']:
                estado_emoji = "✅" if reg['estado'] == 'Cerrado' else "🔄"
                hora_exacta = f" [Registrado: {reg['hora_exacta']}]" if reg['hora_exacta'] != 'N/A' else ""
                st.write(f"  {estado_emoji} {reg['numero']}. {reg['op']} - {reg['horas_trabajadas']:.2f}h ({reg['estado']}){hora_exacta}")
    
    # Información sobre la nueva actividad
    if conteo_resultado['es_primer_registro']:
        st.success("✅ **PRIMER REGISTRO ABSOLUTO DE ESTA CÉDULA**")
        st.info(f"🕐 Tiempo desde las 7:00 AM hasta {conteo_resultado['hora_exacta_registro']} = {conteo_resultado['tiempo_trabajado']:.3f} horas")
        st.info("📋 Registro guardado en Google Sheets")
    else:
        if conteo_resultado.get('es_primer_registro_del_dia', False):
            st.success("✅ **PRIMER REGISTRO DEL DÍA - BASADO EN ÚLTIMA HORA**")
        else:
            st.success("✅ **REGISTRO ADICIONAL - BASADO EN ÚLTIMA HORA**")
        st.info(f"🕐 Tiempo desde {conteo_resultado['hora_inicio_conteo']} hasta {conteo_resultado['hora_exacta_registro']} = {conteo_resultado['tiempo_trabajado']:.3f} horas")
        st.info("� Registro guardado en Google Sheets")
    
    # Información sobre el sistema de horas exactas
    with st.expander("ℹ️ Cómo funciona el nuevo sistema basado en última hora"):
        st.markdown("""
        **🕐 Sistema Basado en Última Hora Registrada:**
        
        1. **El sistema usa la última hora registrada de cada cédula**
        
        2. **Primer registro de una cédula:**
           - Cuenta desde las 7:00 AM hasta la **hora exacta** del registro
           - Se registra el tiempo REAL trabajado 
           - Se guarda inmediatamente en Google Sheets
           
        3. **Registros siguientes:**
           - Busca la **última hora registrada** de esa cédula (hora_exacta > hora_salida > hora_entrada)
           - Calcula tiempo desde esa última hora hasta la hora exacta actual
           - Cada registro refleja el tiempo real trabajado desde la última actividad
           
        4. **Ejemplo basado en tu CSV:**
           - Cédula 1152707808 último registro: 07:17:09 (12 nov)
           - Nuevo registro hoy a las 16:30:00
           - Tiempo calculado: desde 07:17:09 hasta 16:30:00 = ~9.21 horas
           
        5. **Ventajas:**
           - **Continuidad real**: No se pierden horas entre días
           - **Precisión total**: Usa las horas exactas registradas
           - **Flexibilidad**: Se adapta al horario real de cada empleado
           
        6. **Ventajas:**
           - ✅ Precisión al segundo
           - ✅ No hay estimaciones ni redondeos
           - ✅ Auditoría completa con hora exacta
           - ✅ Tiempo real trabajado por actividad
        """)
    
    st.balloons()

def guardar_en_google_sheets_simple(registro):
    """Función ultra-básica para guardar en Google Sheets con máxima confiabilidad"""
    try:
        spreadsheet, mensaje = conectar_google_sheets()
        if not spreadsheet:
            raise Exception(f"No se pudo conectar: {mensaje}")
        
        # Obtener la hoja 'Registros' existente (no crear nueva)
        try:
            worksheet = spreadsheet.worksheet('Registros')
        except:
            raise Exception("No se pudo acceder a la hoja 'Registros'. Verifica que exista en el Google Sheet.")
        
        # Preparar datos según la estructura exacta del Sheet
        fecha_obj = registro.get('fecha')
        if hasattr(fecha_obj, 'strftime'):
            fecha_str = fecha_obj.strftime('%d/%m/%Y')
        else:
            fecha_str = str(fecha_obj)
        
        # Extraer información del servicio si existe
        servicio = registro.get('servicio', '')
        codigo_servicio = ''
        actividad_servicio = ''
        
        if servicio and ' - ' in servicio:
            partes_servicio = servicio.split(' - ', 1)
            codigo_servicio = partes_servicio[0].strip()
            actividad_servicio = partes_servicio[1].strip()
        
        # Formatear horas de entrada y salida
        hora_entrada = registro.get('hora_entrada', '')
        hora_salida = registro.get('hora_salida', '')
        
        if isinstance(hora_entrada, time):
            hora_entrada = hora_entrada.strftime('%H:%M:%S')
        elif not isinstance(hora_entrada, str):
            hora_entrada = str(hora_entrada)
            
        if isinstance(hora_salida, time):
            hora_salida = hora_salida.strftime('%H:%M:%S')
        elif not isinstance(hora_salida, str):
            hora_salida = str(hora_salida)
        
        # 🆕 USAR EL TIEMPO CALCULADO directamente desde el registro
        tiempo_horas_calculado = registro.get('tiempo_horas', 0)
        
        # Asegurar que el tiempo se guarde correctamente incluso para primer registro del día
        print(f"💾 [GUARDADO SHEETS] Tiempo a guardar: {tiempo_horas_calculado} horas")
        
        # Fila de datos según estructura exacta del Sheet:
        # Fecha | Cédula | Nombre | Orden | Cliente | Código | Actividad | Item | Tiempo [Hr] | Observaciones | Proceso | Mes | Año | Semana | REFERENCIA | hora_exacta
        fila_datos = [
            fecha_str,  # Fecha
            str(registro.get('cedula', '')),  # Cédula
            str(registro.get('empleado', '')),  # Nombre
            str(registro.get('op', '')),  # Orden
            str(registro.get('nombre_cliente', '')),  # Cliente
            codigo_servicio,  # Código
            actividad_servicio,  # Actividad
            str(registro.get('descripcion_op', '')),  # Item
            str(tiempo_horas_calculado),  # 🆕 Tiempo [Hr] - USAR EL CALCULADO DESDE LA DIFERENCIA
            '',  # Observaciones (vacío)
            str(registro.get('descripcion_proceso', 'Produccion')),  # Proceso
            str(registro.get('mes', '')),  # Mes
            str(registro.get('año', '')),  # Año
            str(registro.get('semana', '')),  # Semana
            str(registro.get('referencia', str(registro.get('codigo_producto', '')))),  # REFERENCIA
            str(registro.get('hora_exacta', '')),  # hora_exacta (última columna)
        ]
        
        # Agregar la fila
        worksheet.append_row(fila_datos)
        return True
        
    except Exception as e:
        # Si falla, intentar con método aún más básico
        try:
            import gspread
            from google.oauth2.service_account import Credentials
            
            config = load_config()
            gs_config = config.get('google_sheets', {})
            credentials_file = gs_config.get('credentials_file', '')
            
            if credentials_file and os.path.exists(credentials_file):
                scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
                creds = Credentials.from_service_account_file(credentials_file, scopes=scope)
                client = gspread.authorize(creds)
                
                spreadsheet_url = gs_config.get('spreadsheet_url', '')
                if spreadsheet_url:
                    spreadsheet = client.open_by_url(spreadsheet_url)
                    worksheet = spreadsheet.worksheet('Registros')
                    
                    # Solo datos esenciales
                    fila_minima = [
                        str(registro.get('fecha', '')),
                        str(registro.get('empleado', '')),
                        str(registro.get('op', '')),
                        str(registro.get('tiempo_horas', 0))
                    ]
                    
                    worksheet.append_row(fila_minima)
                    return True
        except:
            pass
        
        raise Exception(f"Error guardando en Google Sheets: {str(e)}")

def guardar_en_google_sheets(registro):
    """Guardar registro en Google Sheets en la hoja 'Registros' existente"""
    spreadsheet, mensaje = conectar_google_sheets()
    if spreadsheet:
        try:
            # Obtener el nombre de la hoja de registros de la configuración
            config = load_config()
            worksheet_name = config.get('google_sheets', {}).get('worksheet_registros', 'Registros')
            
            # Acceder a la hoja de registros
            worksheet = spreadsheet.worksheet(worksheet_name)
            
            # Preparar datos según la estructura existente
            fecha_obj = registro['fecha']
            
            # Obtener la información del servicio para usar código y actividad literal
            servicio_info = registro.get('servicio_info', {})
            
            # 🆕 USAR EL TIEMPO CALCULADO DESDE SESSION_STATE (diferencia con último registro)
            tiempo_horas_calculado = registro.get('tiempo_horas', 0)
            
            # Si existe un tiempo calculado en session_state, usarlo (viene de la diferencia con hora_exacta anterior)
            if hasattr(st, 'session_state') and hasattr(st.session_state, 'tiempo_calculado'):
                tiempo_horas_calculado = st.session_state.tiempo_calculado
            
            # Mapear los datos a los encabezados exactos como en la imagen
            # Orden: Fecha | Cédula | Nombre | Orden | Cliente | Código | Actividad | Item | Tiempo [Hr] | Observaciones | Proceso | Usuario de Ingreso | Mes | Año | Semana | REFERENCIA
            fila_registro = [
                fecha_obj.strftime('%d/%m/%Y'),  # Fecha
                str(registro.get('cedula', '')),  # Cédula (de Datos_colab por cedula)
                str(registro.get('empleado', '')),  # Nombre (de Datos_colab por nombre)
                str(registro.get('op', '')),  # Orden (de OPS por orden)
                str(registro.get('nombre_cliente', '')),  # Cliente (de OPS por cliente)
                str(servicio_info.get('numero', '')),  # Código (literal de Servicio)
                str(servicio_info.get('nomservicio', '')),  # Actividad (literal de Servicio)
                str(registro.get('op_info', {}).get('item', '')),  # Item (descripción de la OP)
                str(tiempo_horas_calculado),  # 🆕 Tiempo [Hr] - USAR EL CALCULADO DESDE LA DIFERENCIA
                '',  # Observaciones (vacío)
                'Producción',  # Proceso (con acento como en la imagen)
                str(registro.get('mes', fecha_obj.strftime('%m'))),  # Mes
                str(registro.get('año', fecha_obj.strftime('%Y'))),  # Año
                str(registro.get('semana', str(fecha_obj.isocalendar()[1]))),  # Semana
                str(registro.get('op_info', {}).get('referencia', ''))  # REFERENCIA (de OPS por referencia)
            ]
            
            # Agregar la fila a la hoja existente
            worksheet.append_row(fila_registro)
            
        except Exception as e:
            raise Exception(f"Error guardando en Google Sheets: {str(e)}")

def pantalla_login_admin():
    """Pantalla de login para administrador"""
    st.markdown("<h2 style='text-align: center; color: #dc3545;'>🔒 Acceso de Administrador</h2>", unsafe_allow_html=True)
    
    # Botón para regresar
    if st.button("← Volver al Inicio", type="secondary"):
        st.session_state.screen = 'inicio'
        st.session_state.login_attempts = 0
        st.rerun()
    
    st.markdown("<br>", unsafe_allow_html=True)
    
    # Verificar si está bloqueado
    if esta_bloqueado():
        st.error("🚫 **Acceso temporalmente bloqueado**")
        st.warning("Has excedido el número máximo de intentos de login. Contacta al administrador del sistema.")
        
        config = load_config()
        admin_config = config.get('admin', {})
        max_attempts = admin_config.get('max_attempts', 3)
        
        st.info(f"Intentos realizados: {st.session_state.login_attempts}/{max_attempts}")
        
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            if st.button("🔄 Reiniciar Contador", type="primary", use_container_width=True):
                st.session_state.login_attempts = 0
                st.success("Contador de intentos reiniciado")
                st.rerun()
        
        return
    
    # Formulario de login
    col1, col2, col3 = st.columns([1, 2, 1])
    
    with col2:
        st.markdown("""
        <div style='text-align: center; background-color: #fff3cd; padding: 20px; border-radius: 10px; border-left: 5px solid #ffc107; margin-bottom: 20px;'>
            <h4 style='color: #856404; margin-bottom: 15px;'>🔐 Autenticación Requerida</h4>
            <p style='color: #856404; margin: 0;'>Ingresa la contraseña de administrador para continuar</p>
        </div>
        """, unsafe_allow_html=True)
        
        # Campo de contraseña
        password = st.text_input(
            "Contraseña:",
            type="password",
            placeholder="Ingresa la contraseña de administrador...",
            key="admin_password",
            help="Contraseña requerida para acceder al panel de administración"
        )
        
        st.markdown("<br>", unsafe_allow_html=True)
        
        # Mostrar intentos restantes
        config = load_config()
        admin_config = config.get('admin', {})
        max_attempts = admin_config.get('max_attempts', 3)
        intentos_restantes = max_attempts - st.session_state.login_attempts
        
        if st.session_state.login_attempts > 0:
            if intentos_restantes > 0:
                st.warning(f"⚠️ Intentos restantes: {intentos_restantes}")
            else:
                st.error("❌ Sin intentos restantes")
        
        # Botón de login
        if st.button("🚀 Ingresar al Panel", type="primary", use_container_width=True):
            if password:
                if verificar_contraseña_admin(password):
                    # Login exitoso
                    st.session_state.admin_authenticated = True
                    st.session_state.admin_mode = True
                    st.session_state.screen = 'admin'
                    st.session_state.login_attempts = 0
                    st.success("✅ Acceso autorizado. Redirigiendo...")
                    st.rerun()
                else:
                    # Contraseña incorrecta
                    st.session_state.login_attempts += 1
                    intentos_restantes = max_attempts - st.session_state.login_attempts
                    
                    if intentos_restantes > 0:
                        st.error(f"❌ Contraseña incorrecta. Intentos restantes: {intentos_restantes}")
                    else:
                        st.error("🚫 Máximo de intentos excedido. Acceso bloqueado.")
                    
                    st.rerun()
            else:
                st.warning("⚠️ Por favor ingresa la contraseña")
    
    # Información adicional
    st.markdown("<br><br>", unsafe_allow_html=True)
    
    with st.expander("ℹ️ Información de seguridad"):
        st.write(f"""
        **🔒 Política de Seguridad:**
        - Máximo {max_attempts} intentos de login
        - Acceso bloqueado temporalmente después de exceder los intentos
        - La contraseña es requerida para acceder a:
          - Gestión de empleados y códigos
          - Configuración del sistema
          - Visualización de registros
          - Configuración de Google Sheets
        
        **💡 ¿Olvidaste la contraseña?**
        - Contraseña por defecto: `admin123`
        - Puedes cambiarla en el panel de configuración
        - Contacta al administrador del sistema si no tienes acceso
        """)

def pantalla_admin():
    """Panel de administrador"""
    # Verificación de seguridad adicional
    if not st.session_state.admin_authenticated:
        st.error("🚫 Acceso no autorizado")
        st.warning("Debes iniciar sesión como administrador para acceder a este panel")
        if st.button("🔒 Ir al Login"):
            st.session_state.screen = 'admin_login'
            st.rerun()
        return
    
    st.markdown("<h2 style='text-align: center; color: #dc3545;'>⚙️ Panel de Administrador</h2>", unsafe_allow_html=True)
    
    # Mostrar información del usuario logueado
    st.success("🔓 **Sesión de administrador activa**")
    
    st.markdown("<br>", unsafe_allow_html=True)
    
    col1, col2 = st.columns([1, 4])
    with col1:
        if st.button("← Salir", type="secondary"):
            st.session_state.admin_mode = False
            st.session_state.admin_authenticated = False
            st.session_state.screen = 'inicio'
            st.success("🚪 Sesión de administrador cerrada")
            st.rerun()
    
    tab1, tab2, tab3, tab4, tab5 = st.tabs(["📊 Dashboard", "👥 Empleados", "📋 Registros", "📈 Reportes OP", "⚙️ Configuración"])
    
    with tab1:
        mostrar_dashboard()
    
    with tab2:
        gestionar_empleados()
    
    with tab3:
        ver_registros()
    
    with tab4:
        mostrar_reportes_op()
    
    with tab5:
        configurar_sistema()

def mostrar_dashboard():
    """Mostrar dashboard con estadísticas"""
    st.subheader("📊 Resumen de Hoy")
    
    df = load_data()
    fecha_hoy = obtener_fecha_colombia()
    registros_hoy = df[df['fecha'] == fecha_hoy]
    
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        entradas_hoy = len(registros_hoy[registros_hoy['hora_entrada'].notna()])
        st.metric("Entradas Hoy", entradas_hoy)
    
    with col2:
        salidas_hoy = len(registros_hoy[registros_hoy['hora_salida'].notna() & (registros_hoy['hora_salida'] != '')])
        st.metric("Salidas Hoy", salidas_hoy)
    
    with col3:
        empleados_activos = len(registros_hoy[(registros_hoy['hora_entrada'].notna()) & 
                                            (pd.isna(registros_hoy['hora_salida']) | (registros_hoy['hora_salida'] == ''))])
        st.metric("Empleados Activos", empleados_activos)
    
    with col4:
        if not registros_hoy.empty:
            horas_promedio = registros_hoy['horas_trabajadas'].fillna(0).mean()
            st.metric("Horas Promedio", f"{horas_promedio:.1f}h")
        else:
            st.metric("Horas Promedio", "0h")
    
    # Estado del sistema
    st.subheader("🔧 Estado del Sistema")
    col1, col2 = st.columns(2)
    
    with col1:
        st.info("""
        **📱 Lector de Código de Barras:**
        - ✅ Compatibilidad USB HID habilitada
        - 🔄 Auto-focus activo en pantalla de registro
        - ⚡ Procesamiento en tiempo real
        - 🎯 Validación automática de formatos
        """)
    
    with col2:
        # Verificar configuración de Google Sheets
        config = load_config()
        gs_config = config.get('google_sheets', {})
        gs_status = "🟢 Conectado" if gs_config.get('enabled') else "🔴 Desconectado"
        
        st.info(f"""
        **🔗 Integración Google Sheets:**
        - Estado: {gs_status}
        - Servicios: {'✅ Disponibles' if gs_config.get('enabled') else '❌ No disponibles'}
        - Datos colaboradores: {'✅ Sincronizados' if gs_config.get('enabled') else '❌ Local únicamente'}
        """)
    
    # Información de horarios laborales
    st.subheader("⏰ Horarios Laborales Configurados")
    
    # Obtener horario del día actual
    fecha_actual = obtener_fecha_colombia()
    horario_hoy = obtener_horario_laboral(fecha_actual)
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        horarios_config = config.get('horarios_laborales', {})
        lunes_jueves = horarios_config.get('lunes_a_jueves', {})
        
        st.info(f"""
        **📅 Lunes a Jueves:**
        - Entrada: {lunes_jueves.get('hora_entrada', '07:00')}
        - Salida: {lunes_jueves.get('hora_salida', '16:30')}
        - Horas: {lunes_jueves.get('horas_normales', 8.5)}h
        - Tolerancia: {lunes_jueves.get('tolerancia_entrada', 15)} min
        """)
    
    with col2:
        viernes = horarios_config.get('viernes', {})
        
        st.info(f"""
        **🎉 Viernes:**
        - Entrada: {viernes.get('hora_entrada', '07:00')}
        - Salida: {viernes.get('hora_salida', '15:30')}
        - Horas: {viernes.get('horas_normales', 7.5)}h
        - Tolerancia: {viernes.get('tolerancia_entrada', 15)} min
        """)
    
    with col3:
        # Mostrar horario de hoy
        if horario_hoy:
            dia_nombres = {
                0: 'Lunes', 1: 'Martes', 2: 'Miércoles', 3: 'Jueves',
                4: 'Viernes', 5: 'Sábado', 6: 'Domingo'
            }
            dia_hoy = dia_nombres[fecha_actual.weekday()]
            
            st.success(f"""
            **📍 Horario de Hoy ({dia_hoy}):**
            - Entrada: {horario_hoy.get('hora_entrada', 'N/A')}
            - Salida: {horario_hoy.get('hora_salida', 'N/A')}
            - Horas: {horario_hoy.get('horas_normales', 0)}h
            - Tolerancia: {horario_hoy.get('tolerancia_entrada', 0)} min
            """)
        else:
            st.warning("""
            **📅 Hoy (Domingo):**
            - Día no laboral
            - Sin horario establecido
            """)
    
    # Resumen de Horas por Empleado
    st.subheader("📊 Resumen de Horas por Empleado")
    
    if not registros_hoy.empty:
        # Crear resumen por empleado
        cedulas_hoy = registros_hoy['cedula'].unique()
        empleados_resumen = []
        
        for cedula in cedulas_hoy:
            if pd.notna(cedula) and cedula != '':
                resumen = obtener_resumen_dia_empleado(cedula, fecha_hoy)
                empleado_nombre = registros_hoy[registros_hoy['cedula'] == cedula]['empleado'].iloc[0] if not registros_hoy[registros_hoy['cedula'] == cedula].empty else str(cedula)
                
                empleados_resumen.append({
                    'Empleado': empleado_nombre,
                    'Cédula': cedula,
                    'Total Registros': resumen['total_registros'],
                    'Horas Trabajadas': f"{resumen['tiempo_total_trabajado']:.2f}h"
                })
        
        if empleados_resumen:
            df_resumen = pd.DataFrame(empleados_resumen)
            st.dataframe(df_resumen, use_container_width=True, hide_index=True)
        else:
            st.info("No hay registros de empleados para mostrar")
    else:
        st.info("No hay registros para el día de hoy")
    st.subheader("📋 Registros Recientes")
    if not registros_hoy.empty:
        st.dataframe(registros_hoy.sort_values('hora_entrada', ascending=False), use_container_width=True)
    else:
        st.info("No hay registros para hoy")

def gestionar_empleados():
    """Gestionar empleados y códigos de barras"""
    st.subheader("👥 Gestión de Empleados")
    
    config = load_config()
    
    with st.expander("➕ Agregar Nuevo Empleado"):
        col1, col2 = st.columns(2)
        
        with col1:
            nuevo_nombre = st.text_input("Nombre del empleado:")
        
        with col2:
            nuevo_codigo = st.text_input("Código de barras:")
        
        if st.button("Agregar Empleado"):
            if nuevo_nombre and nuevo_codigo:
                if 'empleados' not in config:
                    config['empleados'] = []
                if 'codigos_barras' not in config:
                    config['codigos_barras'] = {}
                
                config['empleados'].append(nuevo_nombre)
                config['codigos_barras'][nuevo_codigo] = nuevo_nombre
                save_config(config)
                st.success(f"Empleado {nuevo_nombre} agregado con código {nuevo_codigo}")
                st.rerun()
            else:
                st.error("Por favor completa todos los campos")
    
    st.subheader("Lista de Empleados")
    empleados = config.get('empleados', [])
    codigos = config.get('codigos_barras', {})
    
    if empleados:
        for empleado in empleados:
            codigo_empleado = None
            for codigo, nombre in codigos.items():
                if nombre == empleado:
                    codigo_empleado = codigo
                    break
            
            col1, col2, col3 = st.columns([3, 2, 1])
            with col1:
                st.write(f"👤 {empleado}")
            with col2:
                st.write(f"🏷️ {codigo_empleado if codigo_empleado else 'Sin código'}")
            with col3:
                if st.button(f"🗑️", key=f"del_{empleado}"):
                    config['empleados'].remove(empleado)
                    if codigo_empleado:
                        del config['codigos_barras'][codigo_empleado]
                    save_config(config)
                    st.rerun()
    else:
        st.info("No hay empleados registrados")

def ver_registros():
    """Ver y filtrar registros"""
    st.subheader("📋 Registros de Horas")
    
    df = load_data()
    
    if df.empty:
        st.info("No hay registros disponibles")
        return
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        empleados_unicos = ['Todos'] + sorted(df['empleado'].unique().tolist())
        empleado_filtro = st.selectbox("Filtrar por empleado:", empleados_unicos)
    
    with col2:
        fecha_inicio = st.date_input("Fecha inicio:", value=date.today() - pd.Timedelta(days=7))
    
    with col3:
        fecha_fin = st.date_input("Fecha fin:", value=date.today())
    
    df_filtrado = df.copy()
    
    if empleado_filtro != 'Todos':
        df_filtrado = df_filtrado[df_filtrado['empleado'] == empleado_filtro]
    
    df_filtrado = df_filtrado[
        (df_filtrado['fecha'] >= fecha_inicio) & 
        (df_filtrado['fecha'] <= fecha_fin)
    ]
    
    if not df_filtrado.empty:
        st.dataframe(df_filtrado.sort_values('fecha', ascending=False), use_container_width=True)
        
        st.subheader("📊 Estadísticas del Período")
        col1, col2, col3 = st.columns(3)
        
        with col1:
            total_registros = len(df_filtrado)
            st.metric("Total Registros", total_registros)
        
        with col2:
            horas_totales = df_filtrado['horas_trabajadas'].fillna(0).sum()
            st.metric("Horas Totales", f"{horas_totales:.1f}h")
        
        with col3:
            empleados_activos = df_filtrado['empleado'].nunique()
            st.metric("Empleados Activos", empleados_activos)
        
        if st.button("📥 Exportar a CSV"):
            csv = df_filtrado.to_csv(index=False)
            st.download_button(
                label="Descargar CSV",
                data=csv,
                file_name=f"registros_{fecha_inicio}_{fecha_fin}.csv",
                mime="text/csv"
            )
    else:
        st.info("No se encontraron registros con los filtros aplicados")

def obtener_horas_por_op(df_filtrado=None, fecha_inicio=None, fecha_fin=None):
    """Obtener horas trabajadas agrupadas por Orden de Producción"""
    if df_filtrado is None:
        df = load_data()
        if fecha_inicio and fecha_fin:
            df = df[(df['fecha'] >= fecha_inicio) & (df['fecha'] <= fecha_fin)]
    else:
        df = df_filtrado
    
    if df.empty:
        return pd.DataFrame()
    
    # Verificar que existan las columnas necesarias
    if 'op' not in df.columns:
        return pd.DataFrame()
    
    if 'horas_trabajadas' not in df.columns:
        return pd.DataFrame()
    
    # Filtrar solo registros con OP y horas trabajadas válidas
    df_con_op = df[
        (df['op'].notna()) & 
        (df['op'] != '') & 
        (df['op'] != '0') &
        (df['horas_trabajadas'].notna()) & 
        (df['horas_trabajadas'] != '') &
        (df['horas_trabajadas'] != 0)
    ].copy()
    
    if df_con_op.empty:
        return pd.DataFrame()
    
    # Convertir horas_trabajadas a numérico
    df_con_op['horas_trabajadas'] = pd.to_numeric(df_con_op['horas_trabajadas'], errors='coerce')
    
    # Agrupar por OP y sumar horas
    reporte_op = df_con_op.groupby('op').agg({
        'horas_trabajadas': 'sum',
        'nombre_cliente': 'first',  # Tomar el primer cliente (debería ser el mismo para toda la OP)
        'codigo_producto': 'first',  # Tomar la primera referencia
        'descripcion_op': 'first',  # Tomar la primera descripción
        'empleado': 'nunique',  # Contar empleados únicos
        'fecha': ['min', 'max']  # Fechas de inicio y fin
    }).reset_index()
    
    # Aplanar columnas multinivel
    reporte_op.columns = ['OP', 'Horas_Totales', 'Cliente', 'Referencia', 'Item', 'Empleados_Unicos', 'Fecha_Inicio', 'Fecha_Fin']
    
    # Redondear horas y ordenar por horas totales
    reporte_op['Horas_Totales'] = reporte_op['Horas_Totales'].round(2)
    reporte_op = reporte_op.sort_values('Horas_Totales', ascending=False)
    
    return reporte_op

def obtener_detalle_op(op_seleccionada, fecha_inicio=None, fecha_fin=None):
    """Obtener detalle de horas por empleado para una OP específica"""
    df = load_data()
    
    if df.empty:
        return pd.DataFrame()
    
    # Verificar que existan las columnas necesarias
    if 'op' not in df.columns or 'horas_trabajadas' not in df.columns:
        return pd.DataFrame()
    
    if fecha_inicio and fecha_fin:
        df = df[(df['fecha'] >= fecha_inicio) & (df['fecha'] <= fecha_fin)]
    
    # Filtrar por OP específica
    df_op = df[
        (df['op'] == op_seleccionada) &
        (df['horas_trabajadas'].notna()) & 
        (df['horas_trabajadas'] != '') &
        (df['horas_trabajadas'] != 0)
    ].copy()
    
    if df_op.empty:
        return pd.DataFrame()
    
    # Convertir horas_trabajadas a numérico
    df_op['horas_trabajadas'] = pd.to_numeric(df_op['horas_trabajadas'], errors='coerce')
    
    # Agrupar por empleado y fecha para ver detalle
    detalle = df_op.groupby(['empleado', 'fecha']).agg({
        'horas_trabajadas': 'sum',
        'codigo_actividad': lambda x: ', '.join(x.unique()),  # Actividades realizadas
        'hora_entrada': 'first',
        'hora_salida': 'last'
    }).reset_index()
    
    # Redondear horas
    detalle['horas_trabajadas'] = detalle['horas_trabajadas'].round(2)
    detalle = detalle.sort_values(['empleado', 'fecha'])
    
    return detalle

def mostrar_reportes_op():
    """Mostrar reportes de horas trabajadas por OP"""
    st.subheader("📈 Reportes por Orden de Producción")
    
    # Filtros de fecha
    col1, col2, col3 = st.columns(3)
    
    with col1:
        fecha_inicio = st.date_input(
            "Fecha inicio:",
            value=date.today() - pd.Timedelta(days=30),
            key="reporte_fecha_inicio"
        )
    
    with col2:
        fecha_fin = st.date_input(
            "Fecha fin:",
            value=date.today(),
            key="reporte_fecha_fin"
        )
    
    with col3:
        if st.button("🔄 Actualizar Reporte"):
            st.rerun()
    
    # Obtener reporte de horas por OP
    with st.spinner("Generando reporte por OP..."):
        reporte_op = obtener_horas_por_op(fecha_inicio=fecha_inicio, fecha_fin=fecha_fin)
    
    if reporte_op.empty:
        st.warning("📋 No se encontraron registros con OP válidas en el período seleccionado")
        
        # Mostrar información de debug
        with st.expander("🔍 Información de debug"):
            df = load_data()
            df_periodo = df[(df['fecha'] >= fecha_inicio) & (df['fecha'] <= fecha_fin)]
            
            st.write(f"**Total registros en período:** {len(df_periodo)}")
            
            if not df_periodo.empty:
                ops_disponibles = df_periodo['op'].value_counts()
                st.write("**OPs en los datos:**")
                st.dataframe(ops_disponibles)
                
                horas_disponibles = df_periodo['horas_trabajadas'].describe()
                st.write("**Estadísticas de horas:**")
                st.write(horas_disponibles)
        
        return
    
    # Mostrar métricas resumen
    st.subheader("📊 Resumen del Período")
    
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        total_ops = len(reporte_op)
        st.metric("OPs Trabajadas", total_ops)
    
    with col2:
        total_horas = reporte_op['Horas_Totales'].sum()
        st.metric("Horas Totales", f"{total_horas:.1f}h")
    
    with col3:
        total_empleados = reporte_op['Empleados_Unicos'].sum()
        st.metric("Empleados Involucrados", total_empleados)
    
    with col4:
        promedio_horas = reporte_op['Horas_Totales'].mean()
        st.metric("Promedio por OP", f"{promedio_horas:.1f}h")
    
    # Mostrar tabla principal de OPs
    st.subheader("📋 Horas por Orden de Producción")
    
    # Formatear la tabla para mejor visualización
    reporte_display = reporte_op.copy()
    reporte_display['Fecha_Inicio'] = pd.to_datetime(reporte_display['Fecha_Inicio']).dt.strftime('%d/%m/%Y')
    reporte_display['Fecha_Fin'] = pd.to_datetime(reporte_display['Fecha_Fin']).dt.strftime('%d/%m/%Y')
    
    # Renombrar columnas para mejor presentación
    reporte_display = reporte_display.rename(columns={
        'OP': '🏭 Orden de Producción',
        'Horas_Totales': '⏰ Horas Totales',
        'Cliente': '👤 Cliente',
        'Referencia': '📦 Referencia',
        'Item': '📝 Item',
        'Empleados_Unicos': '👥 Empleados',
        'Fecha_Inicio': '📅 Inicio',
        'Fecha_Fin': '📅 Fin'
    })
    
    st.dataframe(
        reporte_display,
        use_container_width=True,
        column_config={
            "⏰ Horas Totales": st.column_config.NumberColumn(
                "⏰ Horas Totales",
                help="Total de horas trabajadas en esta OP",
                format="%.2f h"
            ),
            "👥 Empleados": st.column_config.NumberColumn(
                "👥 Empleados",
                help="Número de empleados únicos que trabajaron en esta OP"
            )
        }
    )
    
    # Gráfico de barras de las top OPs
    st.subheader("📊 Top 10 OPs por Horas Trabajadas")
    
    if len(reporte_op) > 0:
        top_10_ops = reporte_op.head(10)
        
        # Crear gráfico de barras
        import matplotlib.pyplot as plt
        
        fig, ax = plt.subplots(figsize=(12, 6))
        bars = ax.barh(top_10_ops['OP'], top_10_ops['Horas_Totales'], color='#1f77b4')
        
        # Agregar valores en las barras
        for i, (bar, horas) in enumerate(zip(bars, top_10_ops['Horas_Totales'])):
            ax.text(bar.get_width() + 0.1, bar.get_y() + bar.get_height()/2, 
                   f'{horas:.1f}h', va='center', fontsize=10)
        
        ax.set_xlabel('Horas Trabajadas')
        ax.set_ylabel('Orden de Producción')
        ax.set_title('Top 10 OPs por Horas Trabajadas')
        ax.grid(axis='x', alpha=0.3)
        
        plt.tight_layout()
        st.pyplot(fig)
    
    # Sección de detalle por OP
    st.subheader("🔍 Detalle por OP Específica")
    
    if not reporte_op.empty:
        op_seleccionada = st.selectbox(
            "Selecciona una OP para ver el detalle:",
            options=reporte_op['OP'].tolist(),
            format_func=lambda x: f"{x} - {reporte_op[reporte_op['OP']==x]['Horas_Totales'].iloc[0]:.1f}h - {reporte_op[reporte_op['OP']==x]['Cliente'].iloc[0]}"
        )
        
        if op_seleccionada:
            # Mostrar información de la OP seleccionada
            op_info = reporte_op[reporte_op['OP'] == op_seleccionada].iloc[0]
            
            col1, col2 = st.columns(2)
            
            with col1:
                st.markdown(f"""
                <div class='info-message'>
                    <h4>📋 Información de la OP: {op_seleccionada}</h4>
                    <p><strong>👤 Cliente:</strong> {op_info['Cliente']}</p>
                    <p><strong>📦 Referencia:</strong> {op_info['Referencia']}</p>
                    <p><strong>📝 Item:</strong> {op_info['Item']}</p>
                </div>
                """, unsafe_allow_html=True)
            
            with col2:
                st.markdown(f"""
                <div class='success-message'>
                    <h4>📊 Estadísticas</h4>
                    <p><strong>⏰ Horas Totales:</strong> {op_info['Horas_Totales']:.2f} horas</p>
                    <p><strong>👥 Empleados:</strong> {op_info['Empleados_Unicos']} personas</p>
                    <p><strong>📅 Período:</strong> {op_info['Fecha_Inicio']} - {op_info['Fecha_Fin']}</p>
                </div>
                """, unsafe_allow_html=True)
            
            # Obtener detalle de la OP
            detalle_op = obtener_detalle_op(op_seleccionada, fecha_inicio, fecha_fin)
            
            if not detalle_op.empty:
                st.subheader(f"👥 Detalle de Empleados - OP {op_seleccionada}")
                
                # Formatear detalle para mejor presentación
                detalle_display = detalle_op.copy()
                detalle_display['fecha'] = pd.to_datetime(detalle_display['fecha']).dt.strftime('%d/%m/%Y')
                
                detalle_display = detalle_display.rename(columns={
                    'empleado': '👤 Empleado',
                    'fecha': '📅 Fecha',
                    'horas_trabajadas': '⏰ Horas',
                    'codigo_actividad': '🔧 Actividades',
                    'hora_entrada': '🕐 Entrada',
                    'hora_salida': '🕕 Salida'
                })
                
                st.dataframe(
                    detalle_display,
                    use_container_width=True,
                    column_config={
                        "⏰ Horas": st.column_config.NumberColumn(
                            "⏰ Horas",
                            help="Horas trabajadas por el empleado en esta fecha",
                            format="%.2f h"
                        )
                    }
                )
                
                # Resumen por empleado
                st.subheader("📊 Resumen por Empleado")
                resumen_empleados = detalle_op.groupby('empleado')['horas_trabajadas'].sum().reset_index()
                resumen_empleados = resumen_empleados.sort_values('horas_trabajadas', ascending=False)
                resumen_empleados['horas_trabajadas'] = resumen_empleados['horas_trabajadas'].round(2)
                
                col1, col2 = st.columns(2)
                
                with col1:
                    st.dataframe(
                        resumen_empleados.rename(columns={
                            'empleado': '👤 Empleado',
                            'horas_trabajadas': '⏰ Total Horas'
                        }),
                        use_container_width=True
                    )
                
                with col2:
                    # Gráfico de pastel de distribución por empleado
                    fig, ax = plt.subplots(figsize=(8, 6))
                    ax.pie(resumen_empleados['horas_trabajadas'], 
                          labels=resumen_empleados['empleado'], 
                          autopct='%1.1f%%', 
                          startangle=90)
                    ax.set_title(f'Distribución de Horas por Empleado\nOP: {op_seleccionada}')
                    plt.tight_layout()
                    st.pyplot(fig)
    
    # Botón de exportación
    st.subheader("📥 Exportar Datos")
    
    col1, col2 = st.columns(2)
    
    with col1:
        if st.button("📊 Exportar Reporte Completo"):
            csv_completo = reporte_op.to_csv(index=False)
            st.download_button(
                label="💾 Descargar CSV - Reporte por OPs",
                data=csv_completo,
                file_name=f"reporte_ops_{fecha_inicio}_{fecha_fin}.csv",
                mime="text/csv"
            )
    
    with col2:
        if 'op_seleccionada' in locals() and not detalle_op.empty:
            if st.button("📋 Exportar Detalle de OP"):
                csv_detalle = detalle_op.to_csv(index=False)
                st.download_button(
                    label=f"💾 Descargar CSV - OP {op_seleccionada}",
                    data=csv_detalle,
                    file_name=f"detalle_op_{op_seleccionada}_{fecha_inicio}_{fecha_fin}.csv",
                    mime="text/csv"
                )

def configurar_sistema():
    """Configuración del sistema"""
    st.subheader("⚙️ Configuración del Sistema")
    
    config = load_config()
    
    # Sección de seguridad
    st.subheader("🔒 Configuración de Seguridad")
    
    with st.expander("🔑 Cambiar Contraseña de Administrador", expanded=False):
        admin_config = config.get('admin', {})
        
        col1, col2 = st.columns(2)
        
        with col1:
            current_password = st.text_input(
                "Contraseña actual:",
                type="password",
                key="current_pass",
                help="Ingresa la contraseña actual para verificar tu identidad"
            )
        
        with col2:
            new_password = st.text_input(
                "Nueva contraseña:",
                type="password",
                key="new_pass",
                help="Ingresa la nueva contraseña (mínimo 6 caracteres)"
            )
        
        confirm_password = st.text_input(
            "Confirmar nueva contraseña:",
            type="password",
            key="confirm_pass",
            help="Confirma la nueva contraseña"
        )
        
        if st.button("🔄 Cambiar Contraseña"):
            if not all([current_password, new_password, confirm_password]):
                st.error("❌ Todos los campos son obligatorios")
            elif not verificar_contraseña_admin(current_password):
                st.error("❌ La contraseña actual es incorrecta")
            elif len(new_password) < 6:
                st.error("❌ La nueva contraseña debe tener al menos 6 caracteres")
            elif new_password != confirm_password:
                st.error("❌ Las contraseñas nuevas no coinciden")
            else:
                # Cambiar contraseña
                config['admin']['password'] = new_password
                save_config(config)
                st.success("✅ Contraseña cambiada exitosamente")
                
                
                # Limpiar campos
                st.session_state.current_pass = ""
                st.session_state.new_pass = ""
                st.session_state.confirm_pass = ""
    
    # Configuración de intentos de login
    with st.expander("⚙️ Configuración de Intentos de Login", expanded=False):
        admin_config = config.get('admin', {})
        
        max_attempts = st.number_input(
            "Máximo intentos de login:",
            min_value=1,
            max_value=10,
            value=admin_config.get('max_attempts', 3),
            help="Número máximo de intentos antes de bloquear el acceso"
        )
        
        if st.button("💾 Guardar Configuración de Seguridad"):
            config['admin']['max_attempts'] = max_attempts
            save_config(config)
            st.success("✅ Configuración de seguridad actualizada")
    
    # Configuración de horarios laborales
    st.subheader("⏰ Configuración de Horarios Laborales")
    
    with st.expander("📅 Configurar Horarios de Trabajo", expanded=False):
        horarios_config = config.get('horarios_laborales', {})
        
        st.write("**Horario Lunes a Jueves:**")
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            lj_entrada = st.time_input(
                "Hora entrada L-J:",
                value=datetime.strptime(horarios_config.get('lunes_a_jueves', {}).get('hora_entrada', '07:00'), '%H:%M').time(),
                key="lj_entrada"
            )
        
        with col2:
            lj_salida = st.time_input(
                "Hora salida L-J:",
                value=datetime.strptime(horarios_config.get('lunes_a_jueves', {}).get('hora_salida', '16:30'), '%H:%M').time(),
                key="lj_salida"
            )
        
        with col3:
            lj_tolerancia = st.number_input(
                "Tolerancia L-J (min):",
                min_value=0,
                max_value=60,
                value=horarios_config.get('lunes_a_jueves', {}).get('tolerancia_entrada', 15),
                key="lj_tolerancia"
            )
        
        with col4:
            # Calcular horas automáticamente
            entrada_dt = datetime.combine(date.today(), lj_entrada)
            salida_dt = datetime.combine(date.today(), lj_salida)
            if salida_dt < entrada_dt:
                salida_dt += pd.Timedelta(days=1)
            lj_horas = (salida_dt - entrada_dt).total_seconds() / 3600
            st.metric("Horas L-J:", f"{lj_horas:.1f}h")
        
        st.write("**Horario Viernes:**")
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            v_entrada = st.time_input(
                "Hora entrada Viernes:",
                value=datetime.strptime(horarios_config.get('viernes', {}).get('hora_entrada', '07:00'), '%H:%M').time(),
                key="v_entrada"
            )
        
        with col2:
            v_salida = st.time_input(
                "Hora salida Viernes:",
                value=datetime.strptime(horarios_config.get('viernes', {}).get('hora_salida', '15:30'), '%H:%M').time(),
                key="v_salida"
            )
        
        with col3:
            v_tolerancia = st.number_input(
                "Tolerancia Viernes (min):",
                min_value=0,
                max_value=60,
                value=horarios_config.get('viernes', {}).get('tolerancia_entrada', 15),
                key="v_tolerancia"
            )
        
        with col4:
            # Calcular horas automáticamente
            entrada_dt = datetime.combine(date.today(), v_entrada)
            salida_dt = datetime.combine(date.today(), v_salida)
            if salida_dt < entrada_dt:
                salida_dt += pd.Timedelta(days=1)
            v_horas = (salida_dt - entrada_dt).total_seconds() / 3600
            st.metric("Horas Viernes:", f"{v_horas:.1f}h")
        
        if st.button("💾 Guardar Horarios Laborales"):
            config['horarios_laborales'] = {
                'lunes_a_jueves': {
                    'hora_entrada': lj_entrada.strftime('%H:%M'),
                    'hora_salida': lj_salida.strftime('%H:%M'),
                    'horas_normales': round(lj_horas, 1),
                    'tolerancia_entrada': int(lj_tolerancia),
                    'tolerancia_salida': int(lj_tolerancia)
                },
                'viernes': {
                    'hora_entrada': v_entrada.strftime('%H:%M'),
                    'hora_salida': v_salida.strftime('%H:%M'),
                    'horas_normales': round(v_horas, 1),
                    'tolerancia_entrada': int(v_tolerancia),
                    'tolerancia_salida': int(v_tolerancia)
                }
            }
            save_config(config)
            st.success("✅ Horarios laborales actualizados")
            
    
    st.subheader("🔗 Integración con Google Sheets")
    
    # Mensaje informativo sobre la configuración preestablecida
    st.info("📋 **Google Sheets preconfigurado:** Tu hoja de cálculo ya está configurada. Solo necesitas habilitar la integración y proporcionar las credenciales.")
    
    # Información sobre las hojas requeridas
    with st.expander("📊 Estructura requerida en Google Sheets"):
        st.markdown("""
        **Hojas necesarias en tu Google Sheets:**
        
        1. **Hoja "Datos_colab"** - Para datos de colaboradores:
           - Columnas sugeridas: `codigo` (o `codigo_barras`) y `nombre` (o `empleado`)
           - Ejemplo:
             ```
             codigo    | nombre
             12345     | Juan Pérez
             67890     | María García
             ```
        
        2. **Hoja "Servicio"** - Para tipos de servicios:
           - Columnas sugeridas: `numero` (o `id`) y `descripcion` (o `nombre`)
           - Ejemplo:
             ```
             numero | descripcion
             001    | Mantenimiento General
             002    | Limpieza de Oficinas
             003    | Soporte Técnico
             ```
        
        **Nota:** El sistema es flexible con los nombres de columnas y buscará automáticamente las variantes más comunes.
        """)
    
    gs_config = config.get('google_sheets', {})
    
    gs_enabled = st.checkbox(
        "Habilitar integración con Google Sheets",
        value=gs_config.get('enabled', False)
    )
    
    if gs_enabled:
        spreadsheet_id = st.text_input(
            "ID de Google Sheets:",
            value=gs_config.get('spreadsheet_id', ''),
            help="ID del documento de Google Sheets (extraído de la URL)"
        )
        
        worksheet_name = st.text_input(
            "Nombre de la hoja:",
            value=gs_config.get('worksheet_name', 'Datos_colab'),
            help="Nombre de la hoja donde están los datos de colaboradores"
        )
        
        credentials_file = st.text_input(
            "Archivo de credenciales:",
            value=gs_config.get('credentials_file', ''),
            help="Ruta al archivo JSON de credenciales de Google"
        )
        
        # Instrucciones para obtener credenciales
        with st.expander("📖 ¿Cómo obtener las credenciales de Google?"):
            st.markdown("""
            **Pasos para configurar Google Sheets API:**
            
            1. **Crear un proyecto en Google Cloud Console:**
               - Ve a [Google Cloud Console](https://console.cloud.google.com/)
               - Crea un nuevo proyecto o selecciona uno existente
            
            2. **Habilitar Google Sheets API:**
               - En el menú lateral, ve a "APIs y servicios" > "Biblioteca"
               - Busca "Google Sheets API" y habilítala
            
            3. **Crear credenciales de cuenta de servicio:**
               - Ve a "APIs y servicios" > "Credenciales"
               - Haz clic en "Crear credenciales" > "Cuenta de servicio"
               - Completa el formulario y crea la cuenta
            
            4. **Descargar el archivo JSON:**
               - Haz clic en la cuenta de servicio creada
               - Ve a la pestaña "Claves"
               - Haz clic en "Agregar clave" > "Crear nueva clave"
               - Selecciona "JSON" y descarga el archivo
            
            5. **Compartir la hoja de cálculo:**
               - Copia el email de la cuenta de servicio (del archivo JSON)
               - Comparte tu Google Sheets con ese email (como editor)
            
            6. **Configurar en la aplicación:**
               - Guarda el archivo JSON en tu computadora
               - Escribe la ruta completa del archivo en el campo de arriba
            """)
        
        if st.button("Probar Conexión"):
            temp_config = config.copy()
            temp_config['google_sheets'] = {
                'enabled': True,
                'spreadsheet_id': spreadsheet_id,
                'worksheet_name': worksheet_name,
                'credentials_file': credentials_file
            }
            save_config(temp_config)
            
            spreadsheet, mensaje = conectar_google_sheets()
            if spreadsheet:
                st.success("✅ Conexión exitosa con Google Sheets")
            else:
                st.error(f"❌ Error de conexión: {mensaje}")
    
    if st.button("💾 Guardar Configuración"):
        config['google_sheets'] = {
            'enabled': gs_enabled,
            'spreadsheet_id': spreadsheet_id if gs_enabled else '',
            'worksheet_name': worksheet_name if gs_enabled else 'Datos_colab',
            'credentials_file': credentials_file if gs_enabled else ''
        }
        save_config(config)
        st.success("Configuración guardada correctamente")
    
    st.subheader("ℹ️ Información del Sistema")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.info(f"""
        **Archivos del sistema:**
        - Datos: {DATA_FILE}
        - Configuración: {CONFIG_FILE}
        """)
    
    with col2:
        if os.path.exists(DATA_FILE):
            file_size = os.path.getsize(DATA_FILE)
            st.info(f"""
            **Estado de archivos:**
            - Tamaño datos: {file_size} bytes
            - Configuración: {'✅' if os.path.exists(CONFIG_FILE) else '❌'}
            """)
        
        # Mostrar información del Google Sheets preconfigurado
        gs_config = config.get('google_sheets', {})
        st.info(f"""
        **Google Sheets configurado:**
        - ID: {gs_config.get('spreadsheet_id', 'No configurado')[:20]}...
        - Hoja: {gs_config.get('worksheet_name', 'Datos_colab')}
        - Estado: {'🟢 Habilitado' if gs_config.get('enabled', False) else '🔴 Deshabilitado'}
        """)
    
    # Sección de diagnóstico de Google Sheets
    st.subheader("🔍 Diagnóstico de Google Sheets")
    
    if st.button("🔄 Ejecutar Diagnóstico"):
        with st.spinner("Verificando conexión con Google Sheets..."):
            diagnosticos = diagnosticar_conexion_sheets()
            
            st.write("**Resultado del diagnóstico:**")
            for diagnostico in diagnosticos:
                if "✅" in diagnostico:
                    st.success(diagnostico)
                elif "❌" in diagnostico:
                    st.error(diagnostico)
                elif "💡" in diagnostico:
                    st.info(diagnostico)
                else:
                    st.write(diagnostico)
    
    # Enlaces útiles
    st.subheader("📚 Recursos y Documentación")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("""
        **📖 Manuales disponibles:**
        - `CONFIGURACION_GOOGLE_SHEETS.md` - Configuración completa
        - `MANUAL_ACTIVIDADES_CONTINUAS.md` - Sistema de actividades
        - `MANUAL_HORARIOS_LABORALES.md` - Configuración de horarios
        """)
    
    with col2:
        st.markdown("""
        **🔗 Enlaces útiles:**
        - [Google Cloud Console](https://console.cloud.google.com/)
        - [Tu Google Sheet](https://docs.google.com/spreadsheets/d/1r3M71nQK_SxVFycYvmoeDek9KVKfBjFZuPax-v5oIb0/edit)
        - [Documentación Google Sheets API](https://developers.google.com/sheets/api)
        """)

def main():
    """Función principal de la aplicación"""
    
    if st.session_state.admin_mode and st.session_state.screen == 'admin' and st.session_state.admin_authenticated:
        pantalla_admin()
    elif st.session_state.screen == 'admin_login':
        pantalla_login_admin()
    elif st.session_state.screen == 'inicio':
        pantalla_inicio()
    elif st.session_state.screen == 'registro_colaborador':
        pantalla_registro_colaborador()
    else:
        st.session_state.screen = 'inicio'
        pantalla_inicio()

if __name__ == "__main__":
    main()
