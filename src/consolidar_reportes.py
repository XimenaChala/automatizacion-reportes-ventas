"""
Automatización de Consolidación de Reportes de Ventas
-------------------------------------------------------
Este script simula una tarea típica de automatización de procesos (RPA):
varias tiendas envían sus reportes de ventas diarios en archivos CSV
con formatos ligeramente distintos (fechas en formatos diferentes,
datos faltantes, etc.). En lugar de que una persona los revise y
consolide manualmente cada día, este script:

1. Busca automáticamente todos los archivos CSV en una carpeta de entrada.
2. Limpia y normaliza los datos (fechas, valores faltantes).
3. Consolida todo en un único reporte.
4. Calcula totales por producto, por vendedor y por día.
5. Genera un archivo Excel con el resumen listo para enviar o revisar,
   incluyendo un gráfico automático.
6. Registra cada ejecución en un archivo de log (auditoría del proceso).

Autor: Ximena del Pilar Zambrano Chala
"""

import argparse
import glob
import logging
import os
from datetime import datetime

import pandas as pd
from openpyxl.chart import BarChart, Reference


def configurar_logging(log_path: str) -> None:
    logging.basicConfig(
        filename=log_path,
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(message)s",
        encoding="utf-8",
    )
    logging.getLogger().addHandler(logging.StreamHandler())


FORMATOS_FECHA = ("%Y-%m-%d", "%d/%m/%Y")


def _parsear_fecha(valor):
    """Intenta convertir un valor de fecha probando varios formatos conocidos."""
    if pd.isna(valor):
        return pd.NaT
    texto = str(valor).strip()
    for formato in FORMATOS_FECHA:
        try:
            return pd.to_datetime(texto, format=formato)
        except ValueError:
            continue
    return pd.NaT


def cargar_archivos(carpeta_entrada: str) -> pd.DataFrame:
    """Busca y carga todos los CSV de la carpeta, agregando el origen del archivo."""
    archivos = glob.glob(os.path.join(carpeta_entrada, "*.csv"))
    if not archivos:
        raise FileNotFoundError(f"No se encontraron archivos .csv en '{carpeta_entrada}'")

    logging.info("Se encontraron %d archivo(s) para procesar.", len(archivos))
    dataframes = []
    for archivo in archivos:
        logging.info("Leyendo archivo: %s", os.path.basename(archivo))
        df = pd.read_csv(archivo)
        df["origen"] = os.path.basename(archivo)
        dataframes.append(df)

    return pd.concat(dataframes, ignore_index=True)


def limpiar_datos(df: pd.DataFrame) -> pd.DataFrame:
    """Normaliza fechas, corrige tipos y trata valores faltantes."""
    filas_antes = len(df)

    # Las fechas llegan en distintos formatos según la tienda de origen
    # (YYYY-MM-DD y DD/MM/YYYY). pandas no detecta bien formatos mixtos
    # en una sola columna, así que se intenta fila por fila con los
    # formatos conocidos antes de descartar un valor como inválido.
    df["fecha"] = df["fecha"].apply(_parsear_fecha)

    # Cantidades faltantes: se asume 1 unidad (venta mínima) y se deja
    # registro de cuántas se corrigieron, para trazabilidad del proceso.
    faltantes = df["cantidad"].isna().sum()
    if faltantes:
        logging.warning("%d registro(s) sin cantidad; se asumió 1 unidad.", faltantes)
        df["cantidad"] = df["cantidad"].fillna(1)

    df["cantidad"] = df["cantidad"].astype(int)
    df["precio_unitario"] = df["precio_unitario"].astype(float)
    df["total"] = df["cantidad"] * df["precio_unitario"]

    filas_invalidas = df["fecha"].isna().sum()
    if filas_invalidas:
        logging.warning("%d registro(s) con fecha inválida fueron descartados.", filas_invalidas)
        df = df.dropna(subset=["fecha"])

    logging.info("Limpieza completada: %d de %d filas conservadas.", len(df), filas_antes)
    return df


def generar_resumenes(df: pd.DataFrame):
    por_producto = (
        df.groupby("producto")["total"].sum().sort_values(ascending=False).reset_index()
    )
    por_vendedor = (
        df.groupby("vendedor")["total"].sum().sort_values(ascending=False).reset_index()
    )
    por_dia = df.groupby(df["fecha"].dt.date)["total"].sum().reset_index()
    return por_producto, por_vendedor, por_dia


def exportar_excel(df, por_producto, por_vendedor, por_dia, ruta_salida: str) -> None:
    with pd.ExcelWriter(ruta_salida, engine="openpyxl") as writer:
        df.sort_values("fecha").to_excel(writer, sheet_name="Detalle", index=False)
        por_producto.to_excel(writer, sheet_name="Por Producto", index=False)
        por_vendedor.to_excel(writer, sheet_name="Por Vendedor", index=False)
        por_dia.to_excel(writer, sheet_name="Por Día", index=False)

        # Gráfico de barras automático en la hoja "Por Producto"
        hoja = writer.sheets["Por Producto"]
        chart = BarChart()
        chart.title = "Ventas totales por producto"
        chart.y_axis.title = "Total ($)"
        chart.x_axis.title = "Producto"

        datos = Reference(hoja, min_col=2, min_row=1, max_row=len(por_producto) + 1)
        categorias = Reference(hoja, min_col=1, min_row=2, max_row=len(por_producto) + 1)
        chart.add_data(datos, titles_from_data=True)
        chart.set_categories(categorias)
        hoja.add_chart(chart, "E2")

    logging.info("Reporte generado exitosamente en: %s", ruta_salida)


def main():
    parser = argparse.ArgumentParser(
        description="Consolida reportes de ventas dispersos en un único informe Excel."
    )
    parser.add_argument("--input-dir", default="data", help="Carpeta con los CSV de entrada")
    parser.add_argument("--output-dir", default="output", help="Carpeta donde guardar el reporte")
    args = parser.parse_args()

    os.makedirs(args.output_dir, exist_ok=True)
    log_path = os.path.join(args.output_dir, "proceso.log")
    configurar_logging(log_path)

    logging.info("=== Inicio del proceso de consolidación de ventas ===")
    try:
        df = cargar_archivos(args.input_dir)
        df = limpiar_datos(df)
        por_producto, por_vendedor, por_dia = generar_resumenes(df)

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        ruta_salida = os.path.join(args.output_dir, f"reporte_consolidado_{timestamp}.xlsx")
        exportar_excel(df, por_producto, por_vendedor, por_dia, ruta_salida)

        print(f"\nProceso completado. Reporte disponible en: {ruta_salida}")
    except Exception as e:
        logging.error("El proceso falló: %s", e)
        raise
    finally:
        logging.info("=== Fin del proceso ===\n")


if __name__ == "__main__":
    main()
