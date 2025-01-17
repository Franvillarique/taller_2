import pandas as pd  
import numpy as np  
import datetime  
from openpyxl.utils import get_column_letter  

# Leer el archivo Excel  
df = pd.read_excel('Base_control_2.xlsx')  

# Limpieza inicial de datos  
df_limpio = df.dropna(subset=['Nombre Ejecutivo'])  
df_limpio = df_limpio[  
    (df_limpio['Nombre Ejecutivo'] != 'n/d') &  
    (df_limpio['Nombre Ejecutivo'] != 'ejecutivo prueba') &  
    (~df_limpio['Nombre Ejecutivo'].str.contains('ejecutivo prueba', case=False, na=False))  
]  

# Convertir la columna B a datetime y filtrar solo los años 2023 y 2024  
df_limpio.iloc[:, 1] = pd.to_datetime(df_limpio.iloc[:, 1])  
df_limpio = df_limpio[df_limpio.iloc[:, 1].dt.year.isin([2023, 2024])]  

# Función para calcular el bono según el monto de la factura (Propuesta 1)  
def calcular_bono(monto):  
    if monto < 50000000:  
        return monto * 0.0025  
    elif 50000000 <= monto <= 200000000:  
        return monto * 0.0050  
    else:  
        return monto * 0.0060  

# Función para calcular el bono según el ingreso por operación (Propuesta 2)  
def calcular_bono_ingreso(ingreso):  
    if ingreso < 100000:  
        return ingreso * 0.0500  
    elif 100000 <= ingreso <= 1000000:  
        return ingreso * 0.0875  
    else:  
        return ingreso * 0.1000  

# Clasificar las facturas según su monto  
def clasificar_monto(monto):  
    if monto < 50000000:  
        return 'Bajo valor'  
    elif 50000000 <= monto <= 200000000:  
        return 'Medio Valor'  
    else:  
        return 'Alto Valor'  

# Aplicar el cálculo de bonos y clasificaciones  
df_limpio['Bono'] = df_limpio['Monto factura'].apply(calcular_bono)  
df_limpio['Bono_2'] = df_limpio['Ingreso por operación'].apply(calcular_bono_ingreso)  
df_limpio['Clasificación'] = df_limpio['Monto factura'].apply(clasificar_monto)  



def crear_resumen_completo(df):  
    # Agregar columnas de año y mes si no existen  
    df['Año'] = df.iloc[:, 1].dt.year  
    df['Mes'] = df.iloc[:, 1].dt.month  
    
    # Resumen mensual  
    resumen_mensual = df.groupby(['Nombre Ejecutivo', 'Año', 'Mes']).agg({  
        'Monto factura': 'sum',  
        'Ingreso por operación': 'sum',  
        'Bono': 'sum',  
        'Bono_2': 'sum',  
        'Nombre empresa solicitante': 'nunique'  
    }).reset_index()  
    
    # Calcular bono fijo mensual basado en cantidad de clientes  
    def calcular_bono_fijo(num_clientes):  
        if num_clientes < 8:  
            return 50000  
        elif num_clientes <= 12:  
            return 120000  
        else:  
            return 250000  
    
    def clasificar_trafico(num_clientes):  
        if num_clientes < 8:  
            return 'Bajo tráfico'  
        elif num_clientes <= 12:  
            return 'Medio tráfico'  
        else:  
            return 'Alto tráfico'  
    
    # Agregar bonos fijos y clasificación  
    resumen_mensual['Bono_Fijo'] = resumen_mensual['Nombre empresa solicitante'].apply(calcular_bono_fijo)  
    resumen_mensual['Clasificación_Tráfico'] = resumen_mensual['Nombre empresa solicitante'].apply(clasificar_trafico)  
    
    # Calcular totales de bonos (modificado según requerimiento)  
    resumen_mensual['Bono_Total_Mensual_Propuesta_1'] = (  
        resumen_mensual['Bono'] +  # Bono por monto factura  
        resumen_mensual['Bono_Fijo']  # Bono fijo por clientes  
    )  
    
    resumen_mensual['Bono_Total_Mensual_Propuesta_2'] = (  
        resumen_mensual['Bono_2'] +  # Bono por ingresos  
        resumen_mensual['Bono_Fijo']  # Bono fijo por clientes  
    )  
    
    # Renombrar columnas  
    resumen_mensual = resumen_mensual.rename(columns={  
        'Nombre Ejecutivo': 'Ejecutivo',  
        'Monto factura': 'Monto Total Facturas',  
        'Ingreso por operación': 'Ingreso Total',  
        'Bono': 'Bono Propuesta 1',  
        'Bono_2': 'Bono Propuesta 2',  
        'Nombre empresa solicitante': 'Cantidad Clientes'  
    })  
    
    # Crear resumen anual (modificado según requerimiento)  
    resumen_anual = resumen_mensual.groupby(['Ejecutivo', 'Año']).agg({  
        'Monto Total Facturas': 'sum',  
        'Ingreso Total': 'sum',  
        'Bono Propuesta 1': 'sum',  
        'Bono Propuesta 2': 'sum',  
        'Cantidad Clientes': 'mean',  # Promedio de clientes por mes  
        'Bono_Fijo': 'sum',  
        'Bono_Total_Mensual_Propuesta_1': 'sum',  
        'Bono_Total_Mensual_Propuesta_2': 'sum'  
    }).reset_index()  
    
    # Agregar clasificación promedio anual  
    resumen_anual['Clasificación_Tráfico_Promedio'] = resumen_anual['Cantidad Clientes'].apply(clasificar_trafico)  
    
    # Ordenar los resultados  
    resumen_mensual = resumen_mensual.sort_values(['Ejecutivo', 'Año', 'Mes'])  
    resumen_anual = resumen_anual.sort_values(['Ejecutivo', 'Año'])  
    
    # Reordenar las columnas del resumen mensual para mejor visualización  
    resumen_mensual = resumen_mensual[[  
        'Ejecutivo', 'Año', 'Mes',   
        'Monto Total Facturas', 'Ingreso Total',  
        'Bono Propuesta 1', 'Bono Propuesta 2',  
        'Cantidad Clientes', 'Clasificación_Tráfico',  
        'Bono_Fijo',  
        'Bono_Total_Mensual_Propuesta_1',  
        'Bono_Total_Mensual_Propuesta_2'  
    ]]  
    
    # Reordenar las columnas del resumen anual  
    resumen_anual = resumen_anual[[  
        'Ejecutivo', 'Año',  
        'Monto Total Facturas', 'Ingreso Total',  
        'Bono Propuesta 1', 'Bono Propuesta 2',  
        'Cantidad Clientes', 'Bono_Fijo',  
        'Bono_Total_Mensual_Propuesta_1',  
        'Bono_Total_Mensual_Propuesta_2',  
        'Clasificación_Tráfico_Promedio'  
    ]]  
    
    return resumen_mensual, resumen_anual  
    #
    # Crear los resúmenes  
resumen_mensual, resumen_anual = crear_resumen_completo(df_limpio)  

# Crear resumen por ejecutivo   
resumen_ejecutivo = df_limpio.groupby('Nombre Ejecutivo', as_index=False).agg({  
    'Monto factura': 'sum',  
    'Ingreso por operación': 'sum',  
    'Bono': 'sum',  
    'Bono_2': 'sum',  
    'Nombre empresa solicitante': 'nunique'  
})  

# Renombrar columnas del resumen ejecutivo  
resumen_ejecutivo.columns = ['Ejecutivo', 'Monto Total', 'Ingreso Total',
                              'Bono Total P1', 'Bono Total P2', 'Total Clientes']  

# Calcular totales por mes y año (corregido)  
df_limpio['Año'] = df_limpio.iloc[:, 1].dt.year  
df_limpio['Mes'] = df_limpio.iloc[:, 1].dt.month  

totales_mes_año = df_limpio.groupby(['Año', 'Mes']).agg({  
    'Monto factura': 'sum',  
    'Ingreso por operación': 'sum',  
    'Bono': 'sum',  
    'Bono_2': 'sum',  
    'Nombre empresa solicitante': 'nunique'  
}).reset_index()  

# Renombrar columnas de totales_mes_año  
totales_mes_año.columns = ['Año', 'Mes', 'Monto Total', 'Ingreso Total', 
                            'Bono Total P1', 'Bono Total P2', 'Total Clientes']  

# Exportar todos los resúmenes a Excel con formato mejorado  
with pd.ExcelWriter('Resultado_Bonos_2023_2024_Completo.xlsx', engine='openpyxl') as writer:  
    # Datos procesados  
    df_limpio.to_excel(writer, sheet_name='Datos_Procesados', index=False)  
    
    # Resumen mensual por ejecutivo  
    resumen_mensual.to_excel(writer, sheet_name='Resumen_Mensual', index=False)  
    
    # Resumen anual por ejecutivo  
    resumen_anual.to_excel(writer, sheet_name='Resumen_Anual', index=False)  
    
    # Resumen general por ejecutivo  
    resumen_ejecutivo.to_excel(writer, sheet_name='Resumen_Por_Ejecutivo', index=False)  
    
    # Totales por mes y año  
    totales_mes_año.to_excel(writer, sheet_name='Totales_Mes_Año', index=False)  
    
    # Crear pivot tables para análisis adicional  
    pivot_mensual = pd.pivot_table(  
        resumen_mensual,  
        values=['Monto Total Facturas', 'Ingreso Total',   
                'Bono_Total_Mensual_Propuesta_1', 'Bono_Total_Mensual_Propuesta_2'],  
        index=['Ejecutivo'],  
        columns=['Año', 'Mes'],  
        aggfunc='sum'  
    ).round(2)  
    
    pivot_mensual.to_excel(writer, sheet_name='Análisis_Mensual_Pivot')  
    
    # Aplicar formato a todas las hojas  
    for sheet_name in writer.sheets:  
        worksheet = writer.sheets[sheet_name]  
        
        # Ajustar el ancho de las columnas  
        for idx, col in enumerate(worksheet.columns, 1):  
            max_length = 0  
            column = get_column_letter(idx)  
            
            for cell in col:  
                try:  
                    if len(str(cell.value)) > max_length:  
                        max_length = len(str(cell.value))  
                except:  
                    pass  
            
            adjusted_width = (max_length + 2)  
            worksheet.column_dimensions[column].width = adjusted_width  

# Imprimir resúmenes para verificación  
print("\nResumen de datos exportados:")  
print(f"1. Datos Procesados: {len(df_limpio)} registros")  
print(f"2. Resumen Mensual: {len(resumen_mensual)} registros")  
print(f"3. Resumen Anual: {len(resumen_anual)} registros")  
print(f"4. Resumen por Ejecutivo: {len(resumen_ejecutivo)} ejecutivos")  
print(f"5. Totales por Mes y Año: {len(totales_mes_año)} períodos")  

# Mostrar ejemplos de cada resumen  
print("\nEjemplo de Resumen Mensual (primeras 3 filas):")  
print(resumen_mensual.head(3))  

print("\nEjemplo de Resumen Anual (primeras 3 filas):")  
print(resumen_anual.head(3))  

print("\nEjemplo de Totales por Mes y Año (primeras 3 filas):")  
print(totales_mes_año.head(3))
