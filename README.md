## 🧭 Buscador de Piezas

Aplicación de escritorio desarrollada en **Python + Tkinter** que permite buscar información de piezas o pedidos en la base de datos **Azure SQL Server** de Ocasa.  
Los resultados se exportan automáticamente a **Excel** y se abren al finalizar la búsqueda.

---

### ⚙️ Funcionalidades principales

#### 🔐 Conexión segura a SQL Server
- Se conecta de forma segura a **Azure SQL** mediante `pyodbc`.
- Las credenciales se gestionan con **keyring**, evitando exponer la contraseña en el código fuente.

---

#### 🧭 Búsqueda de piezas
- El usuario ingresa uno o varios números de **equipo/pedido**, separados por coma.
- El sistema consulta dinámicamente las tablas:
  - `DW.Fact_Archivo_Ruteo`
  - `DW.Fact_Ruteo_Planificacion_Rutas`
- Devuelve información detallada:
  - Jornada  
  - Equipo  
  - Ruta asignada  
  - Cliente  
  - Dirección  
  - Centro (IATA y Sucursal)  
  - Latitud / Longitud  
  - Fechas de Programación y Despacho  
  - Ruteador

---

#### 📊 Exportación automática a Excel
- Los resultados se guardan en **`resultados_Pieza_Buscada.xlsx`**.
- El archivo se abre automáticamente al finalizar (`os.startfile()`).

---

#### 🖥️ Interfaz gráfica intuitiva
Desarrollada con **Tkinter** y **Pillow**:

- Campo para ingresar las piezas.  
- Botón **“🔍 Buscar”**.  
- **Barra de progreso animada** mientras se ejecuta la consulta.  
- **Mensajes de estado** informativos (“Buscando…”, “✅ Completado”, “❌ Error”).  
- Logo de **Ocasa** en la parte superior.  

---

#### ⚡ Rendimiento fluido
- La consulta SQL se ejecuta en **un hilo separado** (`threading.Thread`), evitando que la interfaz se congele mientras se realiza la búsqueda.

---

### 🧩 Requisitos

- Python 3.9 o superior  
- Librerías necesarias:
  ```bash
  pip install pandas pyodbc pillow keyring openpyxl
