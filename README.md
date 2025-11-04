⚙️ Funciones principales del código

🔐 Conexión segura a SQL Server

Se conecta a la base de datos en AZURE SQL:

🧭 Búsqueda de piezas

El usuario ingresa uno o varios números de equipo o pedido (separados por coma).

Se ejecuta una consulta SQL dinámica que busca esos números en las tablas:

DW.Fact_Archivo_Ruteo

DW.Fact_Ruteo_Planificacion_Rutas

El resultado contiene datos como:

Jornada, equipo, ruta asignada, cliente, dirección, centro, latitud/longitud, fechas y ruteador.

📊 Exportación automática a Excel

Los resultados se guardan en un archivo llamado resultados_Pieza_Buscada.xlsx.

Se abre automáticamente al finalizar (os.startfile()).

🖥️ Interfaz gráfica (Tkinter)

Permite al usuario interactuar fácilmente:

Campo para ingresar piezas.

Botón “🔍 Buscar”.

Barra de progreso animada.

Mensajes de estado (por ejemplo, “Buscando datos…” o “✅ Búsqueda completada”).

Imagen con el logo de Ocasa.

⚡ Ejecución en hilo separado

Usa threading.Thread para que la búsqueda SQL se ejecute en segundo plano, evitando que la interfaz se congele mientras se consulta la base.

💼 En resumen:

El programa es un buscador gráfico de piezas/pedidos, conectado a la base de datos de Ocasa, que permite:

Buscar múltiples piezas al mismo tiempo.

Obtener información detallada desde SQL Server.

Exportar los resultados automáticamente a Excel.

Mostrar el progreso visualmente y mantener una experiencia fluida para el usuario.
