import json
import pandas as pd
import os
import tkinter as tk
import sys

BASE_DIR = os.getcwd()
directorio_salida = os.path.join(BASE_DIR, "output")
directorio_entrada = os.path.join(BASE_DIR, "input")
rutajson = os.path.join(r"I:\2.Informativo\1.Gestion_Diaria\1. Envios Solicitudes\Trejos", "CONCEPTOS.json")

# Colores
COLOR_FONDO = "#FDDA24"
COLOR_BOTONES = "#d5f5e3"
COLOR_TEXTO = 'black'

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

# Función para mostrar la ventana de carga
def mostrar_ventana_cargando():
    carga = tk.Toplevel(root)
    carga.title("Cargando")
    carga.geometry("300x100")
    carga.configure(bg=COLOR_FONDO)
    tk.Label(carga, text="Cargando, por favor espere...", padx=20, pady=20, bg=COLOR_FONDO, fg=COLOR_TEXTO).pack()
    carga.update()
    return carga

# Función para mostrar un mensaje de error personalizado
def mostrar_mensaje_error(titulo, mensaje):
    error_ventana = tk.Toplevel(root)
    error_ventana.title(titulo)
    error_ventana.geometry("300x150")
    error_ventana.configure(bg=COLOR_FONDO)
    tk.Label(error_ventana, text=mensaje, padx=20, pady=20, bg=COLOR_FONDO, fg=COLOR_TEXTO).pack()
    tk.Button(error_ventana, text="Aceptar", command=error_ventana.destroy, bg=COLOR_BOTONES, fg=COLOR_FONDO).pack(pady=10)
    error_ventana.grab_set()

# Función para mostrar la notificación de finalización
def mostrar_notificacion_finalizada():
    notificacion_ventana = tk.Toplevel(root)
    notificacion_ventana.title("Proceso Finalizado")
    notificacion_ventana.geometry("300x150")
    notificacion_ventana.configure(bg=COLOR_FONDO)
    tk.Label(notificacion_ventana, text="El proceso ha terminado.", padx=20, pady=20, bg=COLOR_FONDO, fg=COLOR_TEXTO).pack()
    tk.Button(notificacion_ventana, text="Aceptar", command=notificacion_ventana.destroy, bg=COLOR_BOTONES, fg=COLOR_TEXTO).pack(pady=10)
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
    dialogo = tk.Toplevel(root)
    dialogo.title("Entrada de RUB")
    dialogo.geometry("600x250")
    dialogo.configure(bg=COLOR_FONDO)

    tk.Label(dialogo, text="Digite el rub del caso:", padx=20, pady=20, bg=COLOR_FONDO, fg=COLOR_TEXTO).pack()
    entrada_rub = tk.Entry(dialogo)
    entrada_rub.pack(pady=10)

    # Dropdown para seleccionar la estructura
    estructuras = list(CONCEPTOS["estructuras"].keys())
    seleccion_estructura = tk.StringVar(value=estructuras[0])  # Valor por defecto
    tk.OptionMenu(dialogo, seleccion_estructura, *estructuras).pack(pady=10)

    def aceptar():
        rub = entrada_rub.get()
        estructura = seleccion_estructura.get()
        if rub:
            dialogo.destroy()
            root.withdraw()
            procesar_datos(rub, estructura)
        else:
            mostrar_mensaje_error("Input requerido", "Debe proporcionar un valor para 'rub'.")

    tk.Button(dialogo, text="Aceptar", command=aceptar, bg=COLOR_BOTONES, fg=COLOR_TEXTO, width=12, height=1).pack(pady=10)
    tk.Button(dialogo, text="Cancelar", command=dialogo.destroy, bg=COLOR_BOTONES, fg=COLOR_TEXTO, width=12, height=1).pack(pady=5)
    dialogo.grab_set()

# Crear la ventana principal
root = tk.Tk()
root.title("Principal")
root.geometry("600x1000")
root.configure(bg=COLOR_FONDO)
root.withdraw()  # Ocultar la ventana principal al inicio

# Llamar a la función para crear el cuadro de diálogo personalizado
crear_dialogo_rub()

# Iniciar el bucle principal de la interfaz gráfica
root.mainloop()
