# 📊 Automatización de Consolidación de Reportes de Ventas

Script en Python que automatiza una tarea manual y repetitiva: **consolidar reportes de ventas dispersos en varios archivos, con formatos inconsistentes, en un único informe listo para usar.**

Es un ejemplo del tipo de problema que resuelven las soluciones de **automatización de procesos (RPA)**: eliminar tareas manuales, repetitivas y propensas a error humano, dejando además un registro (log) de cada ejecución para trazabilidad del proceso.

## 🧩 El problema que resuelve

En muchas empresas con varias sucursales o fuentes de datos, cada una entrega su reporte de ventas en un formato ligeramente distinto (fechas en diferente formato, datos faltantes, nombres de columnas no siempre alineados). Consolidar esto manualmente cada día:

- Toma tiempo valioso del equipo.
- Es propenso a errores (copiar y pegar mal, fórmulas rotas).
- No deja un registro claro de qué se corrigió o descartó.

Este script automatiza todo el proceso de principio a fin.

## ⚙️ Qué hace

1. Busca automáticamente todos los archivos `.csv` dentro de una carpeta de entrada.
2. Normaliza los datos: unifica formatos de fecha distintos (`YYYY-MM-DD` y `DD/MM/YYYY`) y trata valores faltantes.
3. Consolida todos los archivos en un único conjunto de datos.
4. Calcula automáticamente:
   - Total de ventas por producto.
   - Total de ventas por vendedor.
   - Total de ventas por día.
5. Genera un archivo **Excel** con el detalle, los resúmenes y un **gráfico de barras automático**.
6. Registra cada ejecución en un archivo de **log** (fecha, archivos procesados, advertencias, errores) para dejar trazabilidad del proceso, tal como se espera de un flujo de automatización real.

## 🛠️ Tecnologías utilizadas

- **Python 3**
- **pandas** — limpieza, transformación y consolidación de datos
- **openpyxl** — generación del archivo Excel y del gráfico
- **logging** (librería estándar) — trazabilidad de cada ejecución
- **argparse** (librería estándar) — ejecución configurable desde línea de comandos

## 📂 Estructura del proyecto

```
automatizacion-reportes-ventas/
├── data/                        # Archivos CSV de entrada (ejemplo incluido)
│   ├── ventas_tienda_centro.csv
│   └── ventas_tienda_norte.csv
├── src/
│   └── consolidar_reportes.py   # Script principal
├── output/                      # Reportes generados y logs (se crea automáticamente)
├── requirements.txt
└── README.md
```

## 🚀 Instalación y uso

```bash
# 1. Clonar el repositorio
git clone https://github.com/XimenaChala/automatizacion-reportes-ventas.git
cd automatizacion-reportes-ventas

# 2. Instalar dependencias
pip install -r requirements.txt

# 3. Ejecutar con los datos de ejemplo
python src/consolidar_reportes.py --input-dir data --output-dir output
```

También puedes usar tus propios archivos CSV: solo colócalos en la carpeta que indiques con `--input-dir` (deben tener las columnas `fecha, producto, cantidad, precio_unitario, vendedor`).

Al finalizar, encontrarás en la carpeta `output/`:
- `reporte_consolidado_<fecha_hora>.xlsx` — el informe final con 4 hojas (Detalle, Por Producto, Por Vendedor, Por Día) y un gráfico.
- `proceso.log` — el registro detallado de la ejecución.

## 🔭 Posibles mejoras futuras

- Integrar el script como una tarea programada (scheduler) para que corra automáticamente cada noche.
- Enviar el reporte por correo automáticamente al finalizar (ej. con `smtplib` o un conector a Outlook/Gmail).
- Adaptar la lógica de limpieza como un flujo dentro de una herramienta RPA (UiPath, Power Automate) para integrarlo con sistemas empresariales (ERP, SAP, Excel compartido en la nube).

## 👤 Autor

**Ximena del Pilar Zambrano Chala**
Estudiante de Ingeniería de Sistemas — Corporación Universitaria del Huila
📧 zambranochalaximena@gmail.com
