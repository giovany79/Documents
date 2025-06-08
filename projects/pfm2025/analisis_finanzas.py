import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from datetime import datetime
import os
import locale
import matplotlib.ticker as mticker

# Configurar el estilo de las gráficas
plt.style.use('seaborn-v0_8')
sns.set_theme()

locale.setlocale(locale.LC_ALL, '')

def formato_pesos(valor):
    try:
        return "$ {:,.2f}".format(valor)
    except:
        return valor

def cargar_datos(archivo):
    """Carga los datos de la hoja 'Movements' del archivo Excel."""
    try:
        df = pd.read_excel(archivo, sheet_name='Movements')
        print("Datos cargados exitosamente de la hoja 'Movements'.")
        print("\nColumnas disponibles:", df.columns.tolist())
        print("\nPrimeras 5 filas:")
        print(df.head())
        return df
    except Exception as e:
        print(f"Error al cargar el archivo: {e}")
        return None

def analisis_basico(df):
    """Realiza análisis básico de los datos."""
    print("\n=== Análisis Básico ===")
    
    # Información general del dataset
    print("\nInformación del dataset:")
    print(df.info())
    
    # Estadísticas descriptivas
    print("\nEstadísticas descriptivas:")
    print(df.describe())
    
    # Valores nulos
    print("\nValores nulos por columna:")
    print(df.isnull().sum())

def crear_visualizaciones(df, carpeta_salida='graficas'):
    """Crea visualizaciones de los datos."""
    # Crear carpeta para las gráficas si no existe
    if not os.path.exists(carpeta_salida):
        os.makedirs(carpeta_salida)
    
    # Intentar crear diferentes tipos de gráficas según las columnas disponibles
    try:
        # Gráfica de líneas para series temporales
        if 'Fecha' in df.columns:
            plt.figure(figsize=(12, 6))
            df.plot(x='Fecha', y=df.select_dtypes(include=['float64', 'int64']).columns[0])
            plt.title('Evolución Temporal')
            plt.xticks(rotation=45)
            plt.tight_layout()
            plt.savefig(f'{carpeta_salida}/evolucion_temporal.png')
            plt.close()
        
        # Histograma para valores numéricos
        for col in df.select_dtypes(include=['float64', 'int64']).columns:
            plt.figure(figsize=(10, 6))
            sns.histplot(data=df, x=col)
            plt.title(f'Distribución de {col}')
            plt.tight_layout()
            plt.savefig(f'{carpeta_salida}/histograma_{col}.png')
            plt.close()
        
        # Gráfica de barras para categorías
        for col in df.select_dtypes(include=['object']).columns:
            plt.figure(figsize=(12, 6))
            df[col].value_counts().plot(kind='bar')
            plt.title(f'Distribución de {col}')
            plt.xticks(rotation=45)
            plt.tight_layout()
            plt.savefig(f'{carpeta_salida}/barras_{col}.png')
            plt.close()
        
        print(f"\nGráficas guardadas en la carpeta '{carpeta_salida}'")
    except Exception as e:
        print(f"Error al crear visualizaciones: {e}")

def ingresos_vs_egresos_por_mes(df, carpeta_salida='graficas'):
    """Genera un diagrama de barras de ingresos vs egresos por mes, ordenando los meses cronológicamente y calculando la diferencia."""
    if not os.path.exists(carpeta_salida):
        os.makedirs(carpeta_salida)
    # Normalizar nombres de columnas
    df = df.rename(columns=lambda x: x.strip().lower().replace(' ', '_'))
    # Convertir columna de fecha
    if 'date' in df.columns:
        df['date'] = pd.to_datetime(df['date'])
        df['mes'] = df['date'].dt.to_period('M')
    else:
        print('No se encontró la columna de fecha.')
        return None
    # Convertir Amount a numérico
    if 'amount' in df.columns:
        df['amount'] = pd.to_numeric(df['amount'], errors='coerce')
    else:
        print('No se encontró la columna Amount.')
        return None
    # Agrupar por mes y tipo
    if 'income/expensive' in df.columns:
        resumen = df.groupby(['mes', 'income/expensive'])['amount'].sum().unstack(fill_value=0)
        resumen = resumen.rename(columns={
            'income': 'Ingresos',
            'expensive': 'Egresos',
            'expense': 'Egresos'
        })
        resumen = resumen.sort_index()
        resumen.index = resumen.index.astype(str)
        if set(['Ingresos', 'Egresos']).issubset(resumen.columns):
            resumen['Diferencia'] = resumen['Ingresos'] - resumen['Egresos']
            resumen = resumen[['Ingresos', 'Egresos', 'Diferencia']]
        resumen[['Ingresos', 'Egresos']].plot(kind='bar', figsize=(12, 6))
        plt.title('Ingresos vs Egresos por Mes')
        plt.ylabel('Monto')
        plt.xlabel('Mes')
        plt.xticks(rotation=45)
        plt.tight_layout()
        ruta_grafica = f'{carpeta_salida}/ingresos_vs_egresos_por_mes.png'
        plt.savefig(ruta_grafica)
        plt.close()
        print(f"Gráfica guardada en {ruta_grafica}")
        return ruta_grafica, resumen
    else:
        print('No se encontró la columna income/expensive.')
        return None, None

def gastos_por_categoria(df, carpeta_salida='reporte'):
    """Genera un diagrama de pastel de gastos totales por categoría y devuelve también la tabla resumen."""
    # Normalizar nombres de columnas
    df = df.rename(columns=lambda x: x.strip().lower().replace(' ', '_'))
    if 'category' in df.columns and 'amount' in df.columns and 'income/expensive' in df.columns:
        # Filtrar solo egresos
        egresos = df[df['income/expensive'].str.lower().str.startswith('expens')]
        egresos['amount'] = pd.to_numeric(egresos['amount'], errors='coerce')
        resumen_cat = egresos.groupby('category')['amount'].sum().sort_values(ascending=False)
        plt.figure(figsize=(10, 8))
        plt.pie(resumen_cat, labels=resumen_cat.index, autopct='%1.1f%%', startangle=140, counterclock=False)
        plt.title('Gastos por Categoría')
        plt.tight_layout()
        ruta_grafica = os.path.join(carpeta_salida, 'gastos_por_categoria.png')
        plt.savefig(ruta_grafica)
        plt.close()
        print(f"Gráfica de gastos por categoría guardada en {ruta_grafica}")
        return 'gastos_por_categoria.png', resumen_cat
    else:
        print('No se encontraron las columnas necesarias para gastos por categoría.')
        return None, None

def historico_restaurant_food(df, carpeta_salida='reporte'):
    """Genera una gráfica de línea con el histórico mensual de gastos en las categorías 'restaurant' y 'food'."""
    df = df.rename(columns=lambda x: x.strip().lower().replace(' ', '_'))
    if 'category' in df.columns and 'amount' in df.columns and 'date' in df.columns and 'income/expensive' in df.columns:
        df['amount'] = pd.to_numeric(df['amount'], errors='coerce')
        df['date'] = pd.to_datetime(df['date'])
        df['mes'] = df['date'].dt.to_period('M')
        filtro = df['category'].str.lower().str.contains('restaurant|food')
        filtro &= df['income/expensive'].str.lower().str.startswith('expens')
        df_filtrado = df[filtro]
        if df_filtrado.empty:
            print('No hay datos para restaurant o food.')
            return None
        resumen = df_filtrado.groupby(['mes', 'category'])['amount'].sum().unstack(fill_value=0)
        resumen = resumen.sort_index()
        plt.figure(figsize=(12, 6))
        ax = plt.gca()
        resumen.plot(ax=ax, marker='o')
        plt.title('Histórico mensual de gastos: Restaurant y Food')
        plt.ylabel('Monto')
        plt.xlabel('Mes')
        plt.xticks(rotation=45)
        ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f'$ {int(x):,}'.replace(',', '.')))
        plt.tight_layout()
        ruta_grafica = os.path.join(carpeta_salida, 'historico_restaurant_food.png')
        plt.savefig(ruta_grafica)
        plt.close()
        print(f"Gráfica de histórico restaurant/food guardada en {ruta_grafica}")
        return 'historico_restaurant_food.png'
    else:
        print('No se encontraron las columnas necesarias para el histórico de restaurant y food.')
        return None

def generar_reporte(df, ruta_grafica, resumen, carpeta_salida='reporte'):
    """Genera un reporte en formato HTML con los análisis y la gráfica de ingresos vs egresos."""
    if not os.path.exists(carpeta_salida):
        os.makedirs(carpeta_salida)
    def formato_html(val):
        try:
            fval = float(val)
            color = 'red' if fval < 0 else 'green' if fval > 0 else 'black'
            return f'<span style="color:{color}">{formato_pesos(fval)}</span>'
        except:
            return val
    # Agregar fila de totales
    if resumen is not None:
        totales = resumen[['Ingresos', 'Egresos', 'Diferencia']].sum()
        totales_row = pd.DataFrame([totales], index=['Total'])
        resumen_con_total = pd.concat([resumen, totales_row])
    else:
        resumen_con_total = resumen
    resumen_html = resumen_con_total.style.format({
            'Ingresos': formato_pesos,
            'Egresos': formato_pesos,
            'Diferencia': formato_html
        }).to_html(escape=False) if resumen_con_total is not None else ''
    ruta_relativa = 'ingresos_vs_egresos_por_mes.png'
    # Nueva gráfica y tabla de gastos por categoría
    ruta_gastos_categoria, tabla_gastos_categoria = gastos_por_categoria(df, carpeta_salida)
    html_gastos_categoria = ''
    if ruta_gastos_categoria:
        tabla_html = tabla_gastos_categoria.apply(formato_pesos).to_frame('Monto').to_html()
        html_gastos_categoria = f"<div class='section'><h2>Gastos por Categoría</h2><img src='{ruta_gastos_categoria}' width='700'/>{tabla_html}</div>"
    # Gráfica histórico restaurant/food
    ruta_historico = historico_restaurant_food(df, carpeta_salida)
    html_historico = f"<div class='section'><h2>Histórico mensual de gastos: Restaurant y Food</h2><img src='{ruta_historico}' width='700'/></div>" if ruta_historico else ''
    reporte = f"""
    <html>
    <head>
        <title>Reporte de Análisis Financiero</title>
        <style>
            body {{ font-family: Arial, sans-serif; margin: 20px; }}
            h1, h2 {{ color: #2c3e50; }}
            .section {{ margin: 20px 0; padding: 20px; background-color: #f8f9fa; border-radius: 5px; }}
            table {{ border-collapse: collapse; width: 100%; }}
            th, td {{ border: 1px solid #ddd; padding: 8px; text-align: left; }}
            th {{ background-color: #f2f2f2; }}
        </style>
    </head>
    <body>
        <h1>Reporte de Análisis Financiero</h1>
        <div class='section'>
            <h2>Ingresos vs Egresos por Mes</h2>
            <img src='{ruta_relativa}' width='700'/>
            {resumen_html}
        </div>
        {html_gastos_categoria}
        {html_historico}
    </body>
    </html>
    """
    with open(f'{carpeta_salida}/reporte.html', 'w', encoding='utf-8') as f:
        f.write(reporte)
    print(f"\nReporte generado en '{carpeta_salida}/reporte.html'")

def main():
    archivo = 'FinanzasPersonalesIA2025.xlsx'
    df = cargar_datos(archivo)
    if df is None:
        return
    ruta_grafica, resumen = ingresos_vs_egresos_por_mes(df)
    generar_reporte(df, ruta_grafica, resumen)

if __name__ == "__main__":
    main() 