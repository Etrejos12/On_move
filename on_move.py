import json
import pandas as pd
import os
import tkinter as tk
from tkinter import ttk
import sys

BASE_DIR = os.getcwd()
directorio_salida = os.path.join(BASE_DIR, "output")
directorio_entrada = os.path.join(BASE_DIR, "input")
rutajson = os.path.join(r"I:\2.Informativo\1.Gestion_Diaria\1. Envios Solicitudes\Trejos", "CONCEPTOS.json")

# ---------------------------------------------------------------------------
# Paleta y tipografía
# El amarillo organizacional se usa como acento (franja superior y botón
# principal); el fondo es neutro para que la ventana sea fácil de leer.
# ---------------------------------------------------------------------------
# Paleta organizacional
COLOR_NEGRO = "#000000"
COLOR_BLANCO = "#FFFFFF"
COLOR_AMARILLO = "#FDDA24"
COLOR_VERDE = "#00C389"
COLOR_MORADO = "#9063CD"
COLOR_NARANJA = "#FF7F41"
COLOR_AZUL = "#59CBE8"
COLOR_ROSADO = "#F5B6CD"

# Usos dentro de la interfaz
COLOR_FONDO = COLOR_BLANCO
COLOR_TEXTO = COLOR_NEGRO
COLOR_EXITO = COLOR_VERDE           # franja de "Proceso finalizado"
COLOR_ERROR = COLOR_NARANJA         # franja de errores
COLOR_BORDE_FOCO = COLOR_AZUL       # borde del campo activo
# Tonos derivados (hover, bordes y texto secundario)
COLOR_AMARILLO_HOVER = "#F0CC10"
COLOR_AMARILLO_PRESIONADO = "#E0BC00"
COLOR_TEXTO_SECUNDARIO = "#555555"
COLOR_BORDE = "#D0D0D0"
COLOR_HOVER_SECUNDARIO = "#F2F2F2"

FUENTE = ("Segoe UI", 10)
FUENTE_NEGRITA = ("Segoe UI", 10, "bold")
FUENTE_TITULO = ("Segoe UI", 12, "bold")

# Cargar configuraciones desde el archivo JSON
try:
    json_path = rutajson
    with open(json_path, 'r', encoding='utf-8') as f:
        CONCEPTOS = json.load(f)

except FileNotFoundError:
    print("Error: No se encontró el archivo CONCEPTOS.json.")
    sys.exit()
except json.JSONDecodeError:
    print("Error: El archivo JSON está mal formado.")
    sys.exit()

def obtener_excel(carpeta):
    archivos = [
    f for f in os.listdir(carpeta)
    if f.endswith((".xlsx", ".xls"))
    ]
    if not archivos:
        raise FileNotFoundError(
        "No se encontró ningún Excel en la carpeta."
        )
    return os.path.join(carpeta, archivos[0])


# Función para detectar el encabezado real en el archivo Excel
def detectar_encabezado_real(path):
    df_raw = pd.read_excel(path, header=None, engine="openpyxl")
    for i in range(min(10, len(df_raw))):
        fila = df_raw.iloc[i].astype(str).str.strip().str.lower()
        # Buscamos la columna que contenga exactamente "descripción transacción"
        if any("descripción transacción" == celda for celda in fila):
            df = pd.read_excel(path, header=i, engine="openpyxl")
            print(f"Encabezado detectado en fila {i}")
            return df
    return None

# ---------------------------------------------------------------------------
# Estilo visual (solo apariencia)
# ---------------------------------------------------------------------------
def configurar_estilos():
    estilo = ttk.Style(root)
    estilo.theme_use("clam")

    estilo.configure(".", background=COLOR_FONDO, foreground=COLOR_TEXTO, font=FUENTE)
    estilo.configure("TFrame", background=COLOR_FONDO)
    estilo.configure("TLabel", background=COLOR_FONDO, foreground=COLOR_TEXTO, font=FUENTE)
    estilo.configure("Titulo.TLabel", font=FUENTE_TITULO)
    estilo.configure("Campo.TLabel", font=FUENTE_NEGRITA, foreground=COLOR_TEXTO_SECUNDARIO)
    estilo.configure("Mensaje.TLabel", foreground=COLOR_TEXTO_SECUNDARIO)

    # Campo de texto
    estilo.configure(
        "TEntry",
        fieldbackground=COLOR_FONDO, foreground=COLOR_TEXTO,
        bordercolor=COLOR_BORDE, lightcolor=COLOR_BORDE, darkcolor=COLOR_BORDE,
        insertcolor=COLOR_TEXTO, padding=(8, 6),
    )
    estilo.map(
        "TEntry",
        bordercolor=[("focus", COLOR_BORDE_FOCO)],
        lightcolor=[("focus", COLOR_BORDE_FOCO)],
    )

    # Lista desplegable
    estilo.configure(
        "TCombobox",
        fieldbackground=COLOR_FONDO, background=COLOR_FONDO, foreground=COLOR_TEXTO,
        bordercolor=COLOR_BORDE, lightcolor=COLOR_BORDE, darkcolor=COLOR_BORDE,
        arrowcolor=COLOR_TEXTO_SECUNDARIO, arrowsize=14, padding=(8, 6),
    )
    estilo.map(
        "TCombobox",
        fieldbackground=[("readonly", COLOR_FONDO)],
        selectbackground=[("readonly", COLOR_FONDO)],
        selectforeground=[("readonly", COLOR_TEXTO)],
        bordercolor=[("focus", COLOR_BORDE_FOCO)],
        background=[("active", COLOR_HOVER_SECUNDARIO)],
    )
    root.option_add("*TCombobox*Listbox.font", FUENTE)
    root.option_add("*TCombobox*Listbox.background", COLOR_FONDO)
    root.option_add("*TCombobox*Listbox.selectBackground", COLOR_AMARILLO)
    root.option_add("*TCombobox*Listbox.selectForeground", COLOR_TEXTO)

    # Botón principal (amarillo organizacional)
    estilo.configure(
        "Primario.TButton",
        background=COLOR_AMARILLO, foreground=COLOR_TEXTO, font=FUENTE_NEGRITA,
        bordercolor=COLOR_AMARILLO, lightcolor=COLOR_AMARILLO, darkcolor=COLOR_AMARILLO,
        focuscolor=COLOR_AMARILLO, borderwidth=1, padding=(18, 7),
    )
    estilo.map(
        "Primario.TButton",
        background=[("pressed", COLOR_AMARILLO_PRESIONADO), ("active", COLOR_AMARILLO_HOVER)],
        bordercolor=[("pressed", COLOR_AMARILLO_PRESIONADO), ("active", COLOR_AMARILLO_HOVER)],
    )

    # Botón secundario (neutro)
    estilo.configure(
        "Secundario.TButton",
        background=COLOR_FONDO, foreground=COLOR_TEXTO, font=FUENTE,
        bordercolor=COLOR_BORDE, lightcolor=COLOR_FONDO, darkcolor=COLOR_FONDO,
        focuscolor=COLOR_FONDO, borderwidth=1, padding=(18, 7),
    )
    estilo.map(
        "Secundario.TButton",
        background=[("pressed", COLOR_BORDE), ("active", COLOR_HOVER_SECUNDARIO)],
    )


def crear_ventana_base(titulo, color_acento=COLOR_AMARILLO, ancho_minimo=360):
    """Ventana con franja de color arriba y cuerpo con márgenes uniformes."""
    ventana = tk.Toplevel(root)
    ventana.title(titulo)
    ventana.configure(bg=COLOR_FONDO)
    ventana.resizable(False, False)
    ventana.minsize(ancho_minimo, 1)

    tk.Frame(ventana, bg=color_acento, height=6).pack(fill="x", side="top")
    cuerpo = ttk.Frame(ventana, padding=(24, 18, 24, 20))
    cuerpo.pack(fill="both", expand=True)
    return ventana, cuerpo


def centrar_ventana(ventana):
    """Ajusta la ventana a su contenido y la centra en la pantalla."""
    ventana.update_idletasks()
    ancho = max(ventana.winfo_reqwidth(), ventana.winfo_width())
    alto = max(ventana.winfo_reqheight(), ventana.winfo_height())
    x = (ventana.winfo_screenwidth() - ancho) // 2
    y = (ventana.winfo_screenheight() - alto) // 3
    ventana.geometry(f"+{x}+{y}")
    ventana.lift()
    ventana.focus_force()


def crear_aviso(titulo, mensaje, color_acento):
    """Ventana de aviso con título, mensaje y botón Aceptar."""
    ventana, cuerpo = crear_ventana_base(titulo, color_acento)
    ttk.Label(cuerpo, text=titulo, style="Titulo.TLabel").pack(anchor="w")
    ttk.Label(cuerpo, text=mensaje, style="Mensaje.TLabel", wraplength=320, justify="left").pack(anchor="w", pady=(6, 18))
    boton = ttk.Button(cuerpo, text="Aceptar", style="Primario.TButton", command=ventana.destroy)
    boton.pack(anchor="e")
    boton.focus_set()
    ventana.bind("<Return>", lambda e: ventana.destroy())
    ventana.bind("<Escape>", lambda e: ventana.destroy())
    centrar_ventana(ventana)
    return ventana

# Función para mostrar la ventana de carga
def mostrar_ventana_cargando():
    carga, cuerpo = crear_ventana_base("Cargando", ancho_minimo=320)
    ttk.Label(cuerpo, text="Procesando", style="Titulo.TLabel").pack(anchor="w")
    ttk.Label(cuerpo, text="Cargando, por favor espere...", style="Mensaje.TLabel").pack(anchor="w", pady=(6, 0))
    centrar_ventana(carga)
    carga.update()
    return carga

# Función para mostrar un mensaje de error personalizado
def mostrar_mensaje_error(titulo, mensaje):
    error_ventana = crear_aviso(titulo, mensaje, COLOR_ERROR)
    error_ventana.grab_set()

# Función para mostrar la notificación de finalización
def mostrar_notificacion_finalizada():
    notificacion_ventana = crear_aviso("Proceso Finalizado", "El proceso ha terminado.", COLOR_EXITO)
    notificacion_ventana.grab_set()

# Función para ejecutar el proceso
def procesar_datos(rub, estructura):
    carga = mostrar_ventana_cargando()
    try:
        archivo_excel = obtener_excel(
            directorio_entrada
            )

        # Leer Excel con lógica para estructura3 usando detectar_encabezado_real
        if estructura == "estructura3":
            df_unificado = detectar_encabezado_real(archivo_excel)
            if df_unificado is None:
                raise ValueError("No se pudo detectar automáticamente el encabezado para estructura3.")
        else:
            # Para otras estructuras, leer todas las hojas y concatenar
            hojas = pd.read_excel(archivo_excel, sheet_name=None)
            df_unificado = pd.concat(hojas.values(), ignore_index=True)

        # Obtener nombre columna para filtrar
        columna_filtrado = CONCEPTOS["estructuras"][estructura]["codigo_trns"]

        # Verificar que la columna existe
        if columna_filtrado not in df_unificado.columns:
            raise ValueError(f"La columna '{columna_filtrado}' no se encuentra en el DataFrame.")

        # Limpiar espacios en la columna
        df_unificado[columna_filtrado] = df_unificado[columna_filtrado].astype(str).str.strip()

        # Filtrar y guardar archivos
        for categoria, conceptos in CONCEPTOS.items():
            if isinstance(conceptos, dict):
                # Si es diccionario, obtenemos la lista de conceptos
                lista_conceptos = list(conceptos.values())
                # Aplanamos la lista en caso de que haya sublistas
                conceptos_flat = []
                for item in lista_conceptos:
                    if isinstance(item, list):
                        conceptos_flat.extend(item)
                    else:
                        conceptos_flat.append(item)
                conceptos = conceptos_flat
            elif not isinstance(conceptos, list):
                conceptos = [conceptos]

            df_filtrado = df_unificado[df_unificado[columna_filtrado].isin(conceptos)]

            if not df_filtrado.empty:
                archivo_salida = os.path.join(directorio_salida, f"{categoria}_{rub}.xlsx")
                df_filtrado.to_excel(archivo_salida, index=False)

        # 🔄 NUEVO BLOQUE: Generar archivo con conceptos no identificados
        conceptos_definidos = set()
        for clave, valores in CONCEPTOS.items():
            if clave in ["archivo_excel", "directorio_salida", "estructuras"]:
                continue
            if isinstance(valores, dict):
                conceptos_definidos.update(valores.values())
            elif isinstance(valores, list):
                conceptos_definidos.update(valores)
            else:
                conceptos_definidos.add(valores)

        df_no_identificados = df_unificado[~df_unificado[columna_filtrado].isin(conceptos_definidos)]

        if not df_no_identificados.empty:
            archivo_no_identificados = os.path.join(directorio_salida, f"conceptos_no_identificados_{rub}.xlsx")
            df_no_identificados.to_excel(archivo_no_identificados, index=False)

        mostrar_notificacion_finalizada()

    except FileNotFoundError:
        mostrar_mensaje_error("Error", "No se encontró el archivo Excel especificado.")
    except Exception as e:
        mostrar_mensaje_error("Error", f"Se produjo un error al filtrar los datos: {e}")
    finally:
        carga.destroy()

# Función para crear el cuadro de diálogo personalizado
def crear_dialogo_rub():
    dialogo, cuerpo = crear_ventana_base("Entrada de RUB", ancho_minimo=400)
    cuerpo.columnconfigure(0, weight=1)

    ttk.Label(cuerpo, text="Datos del caso", style="Titulo.TLabel").grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 14))

    ttk.Label(cuerpo, text="Digite el rub del caso", style="Campo.TLabel").grid(row=1, column=0, columnspan=2, sticky="w")
    entrada_rub = ttk.Entry(cuerpo, font=FUENTE)
    entrada_rub.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(4, 12))

    # Dropdown para seleccionar la estructura
    estructuras = list(CONCEPTOS["estructuras"].keys())
    seleccion_estructura = tk.StringVar(value=estructuras[0])  # Valor por defecto
    ttk.Label(cuerpo, text="Estructura", style="Campo.TLabel").grid(row=3, column=0, columnspan=2, sticky="w")
    ttk.Combobox(
        cuerpo, textvariable=seleccion_estructura, values=estructuras,
        state="readonly", font=FUENTE,
    ).grid(row=4, column=0, columnspan=2, sticky="ew", pady=(4, 20))

    def aceptar():
        rub = entrada_rub.get()
        estructura = seleccion_estructura.get()
        if rub:
            dialogo.destroy()
            root.withdraw()
            procesar_datos(rub, estructura)
        else:
            mostrar_mensaje_error("Input requerido", "Debe proporcionar un valor para 'rub'.")

    # Botones alineados a la derecha: Cancelar (secundario) y Aceptar (principal)
    ttk.Button(cuerpo, text="Cancelar", style="Secundario.TButton", command=dialogo.destroy).grid(row=5, column=0, sticky="e", padx=(0, 8))
    ttk.Button(cuerpo, text="Aceptar", style="Primario.TButton", command=aceptar).grid(row=5, column=1, sticky="e")

    # Atajos de teclado: Enter = Aceptar, Esc = Cancelar
    dialogo.bind("<Return>", lambda e: aceptar())
    dialogo.bind("<Escape>", lambda e: dialogo.destroy())

    centrar_ventana(dialogo)
    entrada_rub.focus_set()
    dialogo.grab_set()

# Crear la ventana principal
root = tk.Tk()
root.title("Principal")
root.geometry("600x1000")
root.configure(bg=COLOR_FONDO)
root.withdraw()  # Ocultar la ventana principal al inicio
configurar_estilos()

# Llamar a la función para crear el cuadro de diálogo personalizado
crear_dialogo_rub()

# Iniciar el bucle principal de la interfaz gráfica
root.mainloop()
