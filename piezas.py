import tkinter as tk #interfaz grafica
from tkinter import ttk, messagebox
from PIL import Image, ImageTk
import pyodbc # conexion a sql
import pandas as pd
import os # para abrir el excel automaticamente.
import threading
import keyring  #para esconder la contraseña a la DB

#conexion SQL
server = "serverdb-bi-prd.database.windows.net"
database = "db01-bi-prd"
username = "Operaciones"
#contraseña encriptada.

# Leer contraseña desde keyring
password = keyring.get_password("ocasa_db", username)
if not password:
    raise SystemExit("❌Problemas con la conexion a la DB")

conn = pyodbc.connect(
    'DRIVER={ODBC Driver 17 for SQL Server};'
    f'SERVER={server};'
    f'DATABASE={database};'
    f'UID={username};'
    f'PWD={password};'
)

def ejecutar_busqueda(pieza_input):
    progress.start()
    label_estado.config(text="Buscando datos, por favor espere...")
    try:
        pieza_list = [p.strip() for p in pieza_input.split(',')]
        pieza_sql = "(" + ",".join(f"'{p}'" for p in pieza_list) + ")"

        query = f"""
        SELECT DISTINCT 
            AR.ID_Jornada,
            AR.Referencia_pedido AS 'Equipo',
            AR.ID_Ruta AS 'Ruta_Asignada',
            AR.Cliente_Descripcion,
            AR.Direccion,
            AR.Centro_Codigo AS 'Iata',
            AR.Centro_Descripcion AS 'Sucursal',
            AR.Latitud,
            AR.Longitud,
            FORMAT(CONVERT(date, CAST(PR.Fecha_Programacion AS CHAR(8))), 'dd-MM-yyyy') AS Fecha_Programacion,
            FORMAT(CONVERT(date, CAST(PR.Fecha_Despacho AS CHAR(8))), 'dd-MM-yyyy') AS Fecha_Despacho,
            PR.Ruteador
        FROM
            DW.Fact_Archivo_Ruteo AS AR
        JOIN    
            DW.Fact_Ruteo_Planificacion_Rutas AS PR
                ON AR.ID_Jornada = PR.[N°_Jornada]
        WHERE
            AR.Referencia_pedido IN {pieza_sql}
        ORDER BY 
            Fecha_Programacion DESC;
        """

        df = pd.read_sql(query, conn)
        output_file = "resultados_Pieza_Buscada.xlsx"
        df.to_excel(output_file, index=False)
        try:
            os.startfile(output_file)
        except Exception:
            pass

        label_estado.config(text="✅ Búsqueda completada. Archivo generado.")
    except Exception as e:
        label_estado.config(text="❌ Error al realizar la búsqueda.")
        root.after(0, lambda: messagebox.showerror("Error", str(e)))
    finally:
        progress.stop()

def buscar():
    pieza_input = entry_pieza.get().strip()
    if not pieza_input:
        messagebox.showwarning("Atención", "Por favor, ingrese al menos una pieza.")
        return
    threading.Thread(target=ejecutar_busqueda, args=(pieza_input,), daemon=True).start()


# INTERFAZ GRÁFICA
root = tk.Tk()
root.title("Buscador de Piezas - Ocasa")
root.geometry("450x550")
root.resizable(False, False)

try:
    image_path = os.path.join(os.path.dirname(__file__), "ocasa imagen.jpg")
    image = Image.open(image_path)
    image = image.resize((200, 100))
    img = ImageTk.PhotoImage(image)
    label_img = tk.Label(root, image=img)
    label_img.pack(pady=10)
except Exception:
    label_img = tk.Label(root, text="(No se pudo cargar la imagen)", fg="red")
    label_img.pack(pady=10)

label_titulo = tk.Label(root, text="Buscar Piezas ", font=("Arial", 14, "bold"))
label_titulo.pack(pady=10)

frame_input = tk.Frame(root)
frame_input.pack(pady=10)

label_pieza = tk.Label(frame_input, text="Ingrese numero de Equipo (separadas por coma):", font=("Arial", 10))
label_pieza.pack(anchor="w")

entry_pieza = tk.Entry(frame_input, width=50, font=("Arial", 11))
entry_pieza.pack(pady=5)

btn_buscar = tk.Button(root, text="🔍 Buscar", font=("Arial", 12, "bold"), bg="#2E86C1", fg="white", width=15, command=buscar)
btn_buscar.pack(pady=15)

progress = ttk.Progressbar(root, mode='indeterminate', length=250)
progress.pack(pady=5)

label_estado = tk.Label(root, text="", font=("Arial", 10), fg="blue")
label_estado.pack(pady=10)

label_footer = tk.Label(root, text="Ruteo Centralizado", font=("Arial", 8), fg="gray")
label_footer.pack(side="bottom", pady=10)

root.mainloop()
conn.close()
