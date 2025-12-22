#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Descarga automática de adjuntos con cualquier término desde Gmail y compresión en ZIP.
- Evita repetir descargas (usa processed_ids.json)
- Añade pausas automáticas para no superar límites
- Comprime todos los archivos descargados en un archivo ZIP
"""

import os, json, time, base64, re, sys, zipfile, traceback, logging, shutil
from datetime import datetime, timedelta
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
import tkinter as tk
from tkinter import simpledialog, messagebox, ttk

APPDATA_DIR = os.path.join(os.getenv('LOCALAPPDATA', '.'), 'GlosasAutomatizador')
os.makedirs(APPDATA_DIR, exist_ok=True)

# Configuración de logging
LOG_FILE = os.path.join(APPDATA_DIR, 'glosas_downloader.log')
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(LOG_FILE, encoding='utf-8'),
        logging.StreamHandler()
    ]
)

def resource_path(relative_path):
    """Obtiene la ruta correcta tanto en .py como en .exe"""
    try:
        base_path = sys._MEIPASS  # Carpeta temporal de PyInstaller
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# Rutas de archivos
CREDENTIALS = resource_path('credentials.json')
TOKEN = os.path.join(APPDATA_DIR, 'token.json')
PROCESSED_FILE = os.path.join(APPDATA_DIR, 'processed_ids.json')
DOWNLOAD_DIR = os.path.join(APPDATA_DIR, 'downloads')


SCOPES = [
    'https://www.googleapis.com/auth/gmail.readonly'
]


# ============================
# VENTANA DE PROGRESO
# ============================
class ProgressWindow:
    """Ventana de progreso visual para mostrar el estado de la descarga"""
    def __init__(self, parent):
        self.window = tk.Toplevel(parent)
        self.window.title("Descargador GLOSAS - Hospital Universitario del Valle")
        self.window.geometry("650x350")
        self.window.resizable(False, False)
        self.window.configure(bg='#f5f5f5')

        # Hacer que la ventana esté siempre al frente
        self.window.attributes('-topmost', True)

        # Centrar ventana
        self.window.update_idletasks()
        x = (self.window.winfo_screenwidth() // 2) - (650 // 2)
        y = (self.window.winfo_screenheight() // 2) - (350 // 2)
        self.window.geometry(f"650x350+{x}+{y}")

        # Header con color
        header_frame = tk.Frame(self.window, bg='#1976D2', height=80)
        header_frame.pack(fill='x')
        header_frame.pack_propagate(False)

        # Título principal
        title_label = tk.Label(
            header_frame,
            text="🏥 Hospital Universitario del Valle",
            font=('Segoe UI', 12, 'bold'),
            fg='white',
            bg='#1976D2'
        )
        title_label.pack(pady=8)

        # Subtítulo
        subtitle_label = tk.Label(
            header_frame,
            text="Descargador Automático de Archivos",
            font=('Segoe UI', 16, 'bold'),
            fg='white',
            bg='#1976D2'
        )
        subtitle_label.pack()

        # Frame para estadísticas con fondo
        stats_frame = tk.Frame(self.window, bg='#f5f5f5')
        stats_frame.pack(pady=15, padx=25, fill='x')

        # Estadísticas con íconos y mejor formato
        self.status_label = tk.Label(
            stats_frame,
            text="⏳ Iniciando procesamiento...",
            font=('Segoe UI', 11),
            bg='#f5f5f5',
            fg='#333',
            justify='left'
        )
        self.status_label.pack(anchor='w', pady=3)

        self.messages_label = tk.Label(
            stats_frame,
            text="📧 Mensajes procesados: 0 / 0",
            font=('Segoe UI', 10),
            bg='#f5f5f5',
            fg='#555',
            justify='left'
        )
        self.messages_label.pack(anchor='w', pady=3)

        self.files_label = tk.Label(
            stats_frame,
            text="📥 Archivos descargados: 0",
            font=('Segoe UI', 11, 'bold'),
            fg='#4CAF50',
            bg='#f5f5f5',
            justify='left'
        )
        self.files_label.pack(anchor='w', pady=3)

        # Separador visual
        separator = tk.Frame(stats_frame, height=1, bg='#ddd')
        separator.pack(fill='x', pady=8)

        self.current_file_label = tk.Label(
            stats_frame,
            text="",
            font=('Segoe UI', 9),
            fg='#666',
            bg='#f5f5f5',
            justify='left',
            wraplength=600
        )
        self.current_file_label.pack(anchor='w', pady=3)

        # Barra de progreso con estilo mejorado
        progress_frame = tk.Frame(self.window, bg='#f5f5f5')
        progress_frame.pack(pady=15, padx=25, fill='x')

        # Estilo personalizado para la barra de progreso
        style = ttk.Style()
        style.theme_use('clam')
        style.configure("Custom.Horizontal.TProgressbar",
                       troughcolor='#e0e0e0',
                       background='#1976D2',
                       thickness=25)

        self.progress_bar = ttk.Progressbar(
            progress_frame,
            mode='determinate',
            length=600,
            style="Custom.Horizontal.TProgressbar"
        )
        self.progress_bar.pack()

        self.progress_label = tk.Label(
            progress_frame,
            text="0%",
            font=('Segoe UI', 11, 'bold'),
            bg='#f5f5f5',
            fg='#1976D2'
        )
        self.progress_label.pack(pady=8)

        # Evitar que se cierre la ventana
        self.window.protocol("WM_DELETE_WINDOW", lambda: None)

    def update_status(self, status):
        """Actualiza el estado general"""
        # Agregar ícono si no tiene
        if not any(icon in status for icon in ['⏳', '✅', '📦', '🔍']):
            status = f"⏳ {status}"
        self.status_label.config(text=status)
        self.window.update()

    def update_progress(self, current, total):
        """Actualiza la barra de progreso"""
        if total > 0:
            percentage = (current / total) * 100
            self.progress_bar['value'] = percentage
            self.progress_label.config(text=f"{percentage:.1f}%")
            self.messages_label.config(text=f"📧 Mensajes procesados: {current} / {total}")
        self.window.update()

    def update_files(self, count):
        """Actualiza el contador de archivos descargados"""
        self.files_label.config(text=f"📥 Archivos descargados: {count}")
        self.window.update()

    def update_current_file(self, filename):
        """Actualiza el archivo actual siendo descargado"""
        if len(filename) > 70:
            filename = filename[:67] + "..."
        self.current_file_label.config(text=f"📄 Descargando: {filename}")
        self.window.update()

    def close(self):
        """Cierra la ventana de progreso"""
        self.window.destroy()


# ============================
# AUTENTICACIÓN
# ============================
def get_credentials():
    try:
        logging.info("Iniciando proceso de autenticación...")
        creds = None
        if os.path.exists(TOKEN):
            try:
                creds = Credentials.from_authorized_user_file(TOKEN, SCOPES)
                logging.info("Token existente cargado correctamente")
            except Exception as e:
                # Si el token está corrupto, lo eliminamos
                logging.warning(f"Token corrupto detectado, eliminando: {e}")
                os.remove(TOKEN)
                creds = None

        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                try:
                    logging.info("Refrescando token expirado...")
                    creds.refresh(Request())
                    logging.info("Token refrescado exitosamente")
                except Exception as e:
                    # Si falla el refresh, reautenticar
                    logging.warning(f"Fallo al refrescar token: {e}")
                    creds = None

            if not creds:
                if not os.path.exists(CREDENTIALS):
                    logging.error(f"No se encontró credentials.json en: {CREDENTIALS}")
                    raise FileNotFoundError(
                        f"No se encontró el archivo credentials.json\n"
                        f"Buscado en: {CREDENTIALS}"
                    )
                logging.info("Iniciando flujo de autenticación OAuth2...")
                flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS, SCOPES)
                creds = flow.run_local_server(port=0)
                logging.info("Autenticación completada exitosamente")

            with open(TOKEN, 'w') as token:
                token.write(creds.to_json())
                logging.info(f"Token guardado en: {TOKEN}")

        return creds
    except Exception as e:
        logging.error(f"Error en autenticación: {str(e)}")
        raise Exception(f"Error al autenticar con Google: {str(e)}")

# ============================
# CARGAR/SALVAR ESTADO
# ============================

def get_date_range(parent_window):
    """Ventana para seleccionar rango de fechas"""
    logging.info("Abriendo ventana de selección de fechas...")

    class DateRangeDialog:
        def __init__(self, parent):
            try:
                logging.info("Iniciando creación de ventana de fechas...")
                self.result = None
                self.top = tk.Toplevel(parent)
                logging.info("Toplevel creado")

                self.top.title("Rango de Fechas - HUV")
                self.top.geometry("500x320")
                self.top.resizable(False, False)
                self.top.configure(bg='#f5f5f5')
                logging.info("Propiedades básicas de ventana establecidas")

                # Hacer que la ventana esté siempre al frente
                self.top.attributes('-topmost', True)
                logging.info("Atributo topmost establecido")

                # Centrar ventana
                self.top.update_idletasks()
                x = (self.top.winfo_screenwidth() // 2) - (500 // 2)
                y = (self.top.winfo_screenheight() // 2) - (320 // 2)
                self.top.geometry(f"500x320+{x}+{y}")
                logging.info(f"Ventana centrada en posición: {x}, {y}")

            except Exception as e:
                logging.error(f"Error al crear ventana de fechas: {e}")
                logging.error(traceback.format_exc())
                raise

            # Header
            header = tk.Frame(self.top, bg='#1976D2', height=70)
            header.pack(fill='x')
            header.pack_propagate(False)

            tk.Label(header,
                    text="📅 Selección de Rango de Fechas",
                    font=('Segoe UI', 14, 'bold'),
                    bg='#1976D2',
                    fg='white').pack(pady=20)
            
            # Frame para fechas con mejor diseño
            content_frame = tk.Frame(self.top, bg='#f5f5f5')
            content_frame.pack(pady=20, padx=30, fill='both', expand=True)

            # Instrucción
            tk.Label(content_frame,
                    text="Ingresa el rango de fechas para la búsqueda:",
                    font=('Segoe UI', 10),
                    bg='#f5f5f5',
                    fg='#555').pack(pady=(0, 15))

            # Frame para inputs de fechas
            dates_frame = tk.Frame(content_frame, bg='#f5f5f5')
            dates_frame.pack(pady=10)

            # Fecha desde
            tk.Label(dates_frame,
                    text="📆 Desde:",
                    font=('Segoe UI', 10, 'bold'),
                    bg='#f5f5f5').grid(row=0, column=0, padx=10, pady=8, sticky='e')
            self.fecha_desde = tk.Entry(dates_frame,
                                       width=18,
                                       font=('Segoe UI', 11),
                                       justify='center')
            self.fecha_desde.grid(row=0, column=1, padx=10, pady=8)
            default_desde = (datetime.now() - timedelta(days=30)).strftime("%d/%m/%Y")
            self.fecha_desde.insert(0, default_desde)

            # Fecha hasta
            tk.Label(dates_frame,
                    text="📆 Hasta:",
                    font=('Segoe UI', 10, 'bold'),
                    bg='#f5f5f5').grid(row=1, column=0, padx=10, pady=8, sticky='e')
            self.fecha_hasta = tk.Entry(dates_frame,
                                       width=18,
                                       font=('Segoe UI', 11),
                                       justify='center')
            self.fecha_hasta.grid(row=1, column=1, padx=10, pady=8)
            default_hasta = datetime.now().strftime("%d/%m/%Y")
            self.fecha_hasta.insert(0, default_hasta)

            # Nota de formato
            tk.Label(content_frame,
                    text="💡 Formato: DD/MM/AAAA (ejemplo: 25/11/2025)",
                    font=('Segoe UI', 9),
                    bg='#f5f5f5',
                    fg='#666').pack(pady=10)

            # Botones con mejor estilo
            btn_frame = tk.Frame(content_frame, bg='#f5f5f5')
            btn_frame.pack(pady=15)

            tk.Button(btn_frame,
                     text="✓ Aceptar",
                     command=self.ok,
                     bg='#4CAF50',
                     fg='white',
                     font=('Segoe UI', 10, 'bold'),
                     width=12,
                     height=2,
                     cursor='hand2').pack(side='left', padx=8)
            tk.Button(btn_frame,
                     text="✗ Cancelar",
                     command=self.cancel,
                     bg='#f44336',
                     fg='white',
                     font=('Segoe UI', 10, 'bold'),
                     width=12,
                     height=2,
                     cursor='hand2').pack(side='left', padx=8)
            
            self.top.protocol("WM_DELETE_WINDOW", self.cancel)
            self.top.transient(parent)

            # FORZAR VISIBILIDAD DE LA VENTANA
            self.top.deiconify()  # Asegurar que no esté minimizada
            self.top.lift()  # Elevar al frente
            self.top.focus_force()  # Forzar el foco
            self.top.update()  # Actualizar la ventana

            self.top.grab_set()
            logging.info("Esperando interacción del usuario con ventana de fechas...")
            parent.wait_window(self.top)
            logging.info("Ventana de fechas cerrada por usuario")

        def ok(self):
            try:
                desde_str = self.fecha_desde.get().strip()
                hasta_str = self.fecha_hasta.get().strip()

                logging.info(f"Usuario seleccionó fechas: {desde_str} - {hasta_str}")

                # Validar formato
                desde = datetime.strptime(desde_str, "%d/%m/%Y")
                hasta = datetime.strptime(hasta_str, "%d/%m/%Y")

                if desde > hasta:
                    logging.warning("Fechas inválidas: 'Desde' mayor que 'Hasta'")
                    messagebox.showerror("Error", "La fecha 'Desde' no puede ser mayor que 'Hasta'")
                    return

                # Formato para Gmail API (YYYY/MM/DD)
                self.result = {
                    'desde': desde.strftime("%Y/%m/%d"),
                    'hasta': hasta.strftime("%Y/%m/%d")
                }
                logging.info(f"Fechas validadas: {self.result['desde']} - {self.result['hasta']}")
                self.top.destroy()
            except ValueError as e:
                logging.warning(f"Formato de fecha inválido: {e}")
                messagebox.showerror("Error", "Formato de fecha inválido. Use DD/MM/AAAA")

        def cancel(self):
            logging.info("Usuario canceló selección de fechas")
            self.top.destroy()
    
    try:
        dialog = DateRangeDialog(parent_window)
        result = dialog.result
        if result:
            logging.info(f"Fechas seleccionadas correctamente: {result}")
        else:
            logging.info("Usuario no seleccionó fechas (canceló)")
        return result
    except Exception as e:
        logging.error(f"Error en get_date_range: {e}")
        logging.error(traceback.format_exc())
        return None

def load_processed_ids():
    """Carga los IDs de correos ya procesados desde el archivo JSON"""
    if os.path.exists(PROCESSED_FILE):
        try:
            with open(PROCESSED_FILE, 'r', encoding='utf-8') as f:
                ids = json.load(f)
                logging.info(f"Cargados {len(ids)} IDs procesados previamente")
                return set(ids)
        except json.JSONDecodeError as e:
            logging.error(f"Archivo processed_ids.json corrupto: {e}")
            # Hacer backup del archivo corrupto
            backup_path = PROCESSED_FILE + f".backup_{int(time.time())}"
            shutil.copy2(PROCESSED_FILE, backup_path)
            logging.info(f"Backup creado en: {backup_path}")
            # Retornar set vacío para empezar de nuevo
            return set()
        except Exception as e:
            logging.error(f"Error al cargar processed_ids: {e}")
            return set()
    logging.info("No hay IDs procesados previamente")
    return set()

def save_processed_ids(ids):
    """Guarda los IDs de correos procesados en el archivo JSON"""
    try:
        with open(PROCESSED_FILE, 'w', encoding='utf-8') as f:
            json.dump(list(ids), f, indent=2)
        logging.info(f"Guardados {len(ids)} IDs procesados")
    except Exception as e:
        logging.error(f"Error al guardar processed_ids: {e}")
        raise

# ============================
# FUNCIONES PRINCIPALES
# ============================
def get_parts(payload):
    parts = []
    if 'parts' in payload:
        for p in payload['parts']:
            parts.extend(get_parts(p))
    else:
        parts.append(payload)
    return parts

def download_attachment(service, msg_id, part, msg_date=None):
    """Descarga un archivo adjunto de un mensaje de Gmail y lo organiza por fecha"""
    try:
        filename = part.get('filename')
        if not filename:
            logging.warning("Part sin nombre de archivo")
            return None

        body = part.get('body', {})
        att_id = body.get('attachmentId')
        if not att_id:
            logging.warning(f"No hay attachmentId para {filename}")
            return None

        logging.info(f"Descargando adjunto: {filename}")

        # Obtener el adjunto de la API
        att = service.users().messages().attachments().get(
            userId='me',
            messageId=msg_id,
            id=att_id
        ).execute()

        data = att.get('data')
        if not data:
            logging.warning(f"No hay datos en el adjunto {filename}")
            return None

        # Decodificar el contenido
        file_data = base64.urlsafe_b64decode(data.encode('UTF-8'))

        # Si no hay fecha del mensaje, usar fecha actual
        if msg_date is None:
            msg_date = datetime.now()

        # Nombres de meses en español
        meses = {
            1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril",
            5: "Mayo", 6: "Junio", 7: "Julio", 8: "Agosto",
            9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"
        }

        # Extraer año, mes y semana del año
        year = msg_date.year
        month_num = msg_date.month
        month_name = f"{year}-{meses[month_num]}"  # 2025-Noviembre
        week_number = msg_date.isocalendar()[1]  # Número de semana del año

        # Crear estructura de carpetas dentro de DOWNLOAD_DIR: {Year-MonthName}/Semana_{N}/
        month_folder = os.path.join(DOWNLOAD_DIR, month_name)
        week_folder = os.path.join(month_folder, f"Semana_{week_number}")
        os.makedirs(week_folder, exist_ok=True)

        # Agregar fecha al nombre del archivo: YYYY-MM-DD_{original_filename}
        date_prefix = msg_date.strftime("%Y-%m-%d")
        new_filename = f"{date_prefix}_{filename}"

        # Generar ruta única si el archivo ya existe
        path = os.path.join(week_folder, new_filename)
        if os.path.exists(path):
            base, ext = os.path.splitext(new_filename)
            path = os.path.join(week_folder, f"{base}_{int(time.time())}{ext}")
            logging.info(f"Archivo duplicado, renombrado a: {os.path.basename(path)}")

        # Guardar archivo
        with open(path, 'wb') as f:
            f.write(file_data)

        file_size = len(file_data) / 1024  # KB
        logging.info(f"Descargado exitosamente: {new_filename} ({file_size:.2f} KB) en {week_folder}")
        return path

    except Exception as e:
        logging.error(f"Error al descargar adjunto {filename}: {e}")
        return None

def create_zip_file(downloaded_files):
    """Crea un archivo ZIP con estructura de carpetas organizadas por mes y semana"""
    if not downloaded_files:
        logging.info("No hay archivos para comprimir")
        return None

    try:
        # Obtener la ruta del escritorio del usuario
        desktop = os.path.join(os.path.expanduser('~'), 'Desktop')

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        zip_filename = f"GLOSAS_{timestamp}.zip"
        zip_path = os.path.join(desktop, zip_filename)

        logging.info(f"Creando archivo ZIP con estructura de carpetas: {zip_filename}")

        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            # Recorrer toda la estructura de carpetas en DOWNLOAD_DIR
            for root_dir, dirs, files in os.walk(DOWNLOAD_DIR):
                for file in files:
                    file_path = os.path.join(root_dir, file)
                    # Calcular la ruta relativa para mantener la estructura de carpetas
                    arcname = os.path.relpath(file_path, DOWNLOAD_DIR)
                    zipf.write(file_path, arcname)
                    logging.info(f"Agregado al ZIP: {arcname}")

        logging.info(f"ZIP creado exitosamente: {zip_path}")

        # Limpiar toda la estructura de carpetas temporales
        logging.info("Limpiando archivos y carpetas temporales...")
        try:
            if os.path.exists(DOWNLOAD_DIR):
                shutil.rmtree(DOWNLOAD_DIR)
                logging.info(f"Directorio temporal eliminado completamente: {DOWNLOAD_DIR}")
        except Exception as e:
            logging.warning(f"No se pudo eliminar directorio temporal: {e}")

        return zip_path

    except Exception as e:
        logging.error(f"Error al crear archivo ZIP: {e}")
        raise Exception(f"Error al crear archivo ZIP: {str(e)}")

def process_emails(remitente, keyword, fecha_desde=None, fecha_hasta=None, parent_window=None):
    """Procesa correos del remitente y descarga archivos con la palabra clave en el nombre"""
    progress_window = None
    try:
        logging.info(f"Iniciando procesamiento de correos de: {remitente}")

        # Autenticación
        creds = get_credentials()
        gmail = build('gmail', 'v1', credentials=creds)

        # Cargar IDs ya procesados
        processed_ids = load_processed_ids()

        # Construir query de búsqueda
        query = f"from:{remitente} has:attachment"

        # Agregar filtro de fechas si está disponible (hacer "hasta" inclusivo)
        if fecha_desde and fecha_hasta:
            # Sumar 1 día a fecha_hasta para hacerla inclusiva
            hasta_date = datetime.strptime(fecha_hasta, "%Y/%m/%d")
            hasta_inclusiva = (hasta_date + timedelta(days=1)).strftime("%Y/%m/%d")
            query += f" after:{fecha_desde} before:{hasta_inclusiva}"
            logging.info(f"Buscando correos desde {fecha_desde} hasta {fecha_hasta} (inclusivo)")

        logging.info(f"Query de búsqueda: {query}")

        # Primera solicitud a la API para contar mensajes
        response = gmail.users().messages().list(userId='me', q=query, maxResults=500).execute()

        # Contar total de mensajes para la barra de progreso
        messages_first_page = response.get('messages', [])
        total_messages_estimate = len(messages_first_page)

        # Crear ventana de progreso
        if parent_window:
            progress_window = ProgressWindow(parent_window)
            progress_window.update_status("Contando mensajes...")
            progress_window.update_progress(0, 100)

        new_processed = set()
        downloaded_files = []
        total_messages = 0
        processed_count = 0

        # Procesar todos los mensajes
        while True:
            messages = response.get('messages', [])
            total_messages += len(messages)
            logging.info(f"Procesando {len(messages)} mensajes...")

            for msg in messages:
                msg_id = msg['id']
                processed_count += 1

                # Actualizar progreso
                if progress_window:
                    progress_window.update_status(f"Procesando mensajes ({processed_count}/{total_messages})...")
                    progress_window.update_progress(processed_count, total_messages)

                # Saltar si ya fue procesado
                if msg_id in processed_ids:
                    logging.debug(f"Mensaje {msg_id} ya procesado, saltando...")
                    continue

                # Obtener detalles del mensaje
                msg_data = gmail.users().messages().get(userId='me', id=msg_id).execute()

                # Extraer fecha del mensaje (en milisegundos desde epoch)
                internal_date = msg_data.get('internalDate')
                if internal_date:
                    msg_date = datetime.fromtimestamp(int(internal_date) / 1000.0)
                else:
                    msg_date = datetime.now()

                parts = get_parts(msg_data.get('payload', {}))

                # Buscar archivos con la palabra clave en el nombre
                for part in parts:
                    filename = part.get('filename', '')
                    if filename and keyword in filename.upper():
                        logging.info(f"Archivo encontrado: {filename}")

                        # Actualizar UI con el archivo actual
                        if progress_window:
                            progress_window.update_current_file(filename)

                        # Descargar con la fecha del mensaje
                        path = download_attachment(gmail, msg_id, part, msg_date)
                        if path:
                            downloaded_files.append(path)
                            logging.info(f"Descargado: {filename}")

                            # Actualizar contador de archivos
                            if progress_window:
                                progress_window.update_files(len(downloaded_files))

                new_processed.add(msg_id)

                # Pausa ligera para evitar límites de la API
                time.sleep(0.3)

            # Guardar progreso después de cada página
            save_processed_ids(processed_ids.union(new_processed))

            # Verificar si hay más páginas
            if 'nextPageToken' in response:
                logging.info("Obteniendo siguiente página de resultados...")
                if progress_window:
                    progress_window.update_status("Obteniendo más mensajes...")

                response = gmail.users().messages().list(
                    userId='me',
                    q=query,
                    pageToken=response['nextPageToken'],
                    maxResults=500
                ).execute()
            else:
                break

        # Cerrar ventana de progreso
        if progress_window:
            progress_window.update_status("Creando archivo ZIP...")
            progress_window.update_progress(100, 100)

        logging.info(f"Procesamiento completado: {total_messages} mensajes revisados, {len(downloaded_files)} archivos descargados")

        # Crear archivo ZIP con estructura de carpetas
        if downloaded_files:
            zip_path = create_zip_file(downloaded_files)
            if progress_window:
                progress_window.close()

            if zip_path:
                mensaje = f"✅ Proceso completado exitosamente\n\n"
                mensaje += f"📦 Archivo: {os.path.basename(zip_path)}\n"
                mensaje += f"📁 Ubicación: ESCRITORIO\n"
                mensaje += f"📂 Estructura: Carpetas organizadas por mes y semana\n\n"
                mensaje += f"Total de archivos descargados: {len(downloaded_files)}\n"
                mensaje += f"Mensajes revisados: {total_messages}"
                messagebox.showinfo("Descarga Completada", mensaje)
                logging.info("Proceso completado exitosamente")
        else:
            if progress_window:
                progress_window.close()

            mensaje = f"No se encontraron archivos nuevos con 'GLOSAS' para descargar.\n\n"
            mensaje += f"Mensajes revisados: {total_messages}\n"
            mensaje += f"Ya procesados previamente: {len(processed_ids)}"
            messagebox.showwarning("Sin Resultados", mensaje)
            logging.info("No se encontraron archivos nuevos")

    except Exception as e:
        if progress_window:
            progress_window.close()
        logging.error(f"Error en process_emails: {str(e)}")
        logging.error(traceback.format_exc())
        raise Exception(f"Error al procesar correos: {str(e)}")

# ============================
# EJECUCIÓN
# ============================


if __name__ == "__main__":
    # Crear una ventana raíz que permanezca durante toda la ejecución
    root = tk.Tk()
    root.title("Descargador de GLOSAS - HUV")
    root.geometry("1x1")  # Hacer la ventana muy pequeña
    root.attributes('-alpha', 0.0)  # Hacer la ventana transparente
    # No usar withdraw() para que las ventanas hijas sean visibles

    try:
        logging.info("="*60)
        logging.info("INICIANDO DESCARGADOR DE GLOSAS - HUV")
        logging.info("="*60)
        logging.info(f"Directorio de datos: {APPDATA_DIR}")
        logging.info(f"Archivo de log: {LOG_FILE}")

        # Mensaje de bienvenida
        messagebox.showinfo(
            "Descargador de GLOSAS - HUV",
            "Bienvenido al descargador automático de archivos GLOSAS\n\n"
            "Desarrollado por Innovación y Desarrollo HUV\n"
            "📧 innovacionydesarrollo@correohuv.gov.co\n\n"
            "⚠️ IMPORTANTE: Este programa solo debe usarse con:\n"
            "   glosasydevoluciones@correohuv.gov.co\n\n"
            "Se creará un archivo ZIP en tu ESCRITORIO con:\n"
            "✓ Carpetas organizadas por mes (nombre del mes)\n"
            "✓ Subcarpetas por semana\n"
            "✓ Archivos con fecha en el nombre\n\n"
            f"Log: {os.path.basename(LOG_FILE)}"
        )
        
        # Obtener remitente
        remitente = simpledialog.askstring(
            "Descargador de GLOSAS",
            "Ingresa el correo del remitente:",
            parent=root
        )

        keyword = simpledialog.askstring(
            "Descargador de GLOSAS",
            "Ingresa la palabra clave:",
            parent=root
        )

        if not remitente or not keyword:
            logging.info("Usuario canceló: no se ingresó correo")
            messagebox.showwarning("Cancelado", "No se ingresó ningún correo")
            root.destroy()
            sys.exit(0)

        remitente = remitente.strip()
        keyword = keyword.strip()
        logging.info(f"Remitente ingresado: {remitente}")
        logging.info(f"Palabra clave ingresada: {keyword}")

        # Obtener fechas
        fechas = get_date_range(root)
        if not fechas:
            logging.info("Usuario canceló: no se seleccionaron fechas")
            messagebox.showwarning("Cancelado", "Operación cancelada por el usuario")
            root.destroy()
            sys.exit(0)

        logging.info(f"Rango de fechas seleccionado: {fechas['desde']} a {fechas['hasta']}")

        # Mensaje de procesamiento con mejor formato
        messagebox.showinfo(
            "🔍 Iniciando Búsqueda - HUV",
            f"📧 Remitente: {remitente}\n\n"
            f"📅 Período de búsqueda:\n"
            f"   • Desde: {fechas['desde']}\n"
            f"   • Hasta: {fechas['hasta']}\n\n"
            "⏳ Esto puede tomar unos momentos...\n\n"
            "🌐 Se abrirá tu navegador para autenticación\n"
            "   de Gmail si es necesario."
        )

        # Procesar emails
        process_emails(remitente, keyword, fechas['desde'], fechas['hasta'], root)

    except Exception as e:
        error_details = traceback.format_exc()
        logging.error("="*60)
        logging.error("ERROR CRÍTICO")
        logging.error("="*60)
        logging.error(error_details)

        messagebox.showerror(
            "Error",
            f"Ocurrió un error:\n\n{str(e)}\n\n"
            f"Revisa el archivo de log para más detalles:\n"
            f"{LOG_FILE}\n\n"
            f"Detalles:\n{error_details[-500:]}"
        )
    finally:
        logging.info("Cerrando aplicación")
        logging.info("="*60)
        root.destroy()
