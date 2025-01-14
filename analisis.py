import pandas as pd  
import os  

# Definir la ruta del directorio  
ruta_directorio = r'C:\Users\Francisco Villanueva\Desktop\TALLER_2'  

try:  
    # Leer el archivo Excel original  
    df = pd.read_excel(os.path.join(ruta_directorio, 'Base_control_2.xlsx'))  
    
    # Limpieza básica de datos  
    df_limpio = df.dropna(subset=['Nombre Ejecutivo'])  
    df_limpio = df_limpio[  
        (df_limpio['Nombre Ejecutivo'] != 'n/d') &  
        (df_limpio['Nombre Ejecutivo'] != 'ejecutivo prueba') &  
        (~df_limpio['Nombre Ejecutivo'].str.contains('ejecutivo prueba', case=False, na=False))  
    ]  
    
    # Convertir fecha y filtrar años  
    df_limpio.iloc[:, 1] = pd.to_datetime(df_limpio.iloc[:, 1])  
    df_limpio = df_limpio[df_limpio.iloc[:, 1].dt.year.isin([2023, 2024])]  
    
    # Funciones de cálculo de bonos  
    def calcular_bono_propuesta1(monto):  
        if monto < 50000000:  
            return monto * 0.0025  
        elif 50000000 <= monto <= 200000000:  
            return monto * 0.0050  
        else:  
            return monto * 0.0060  

    def calcular_bono_propuesta2(ingreso):  
        if ingreso < 100000:  
            return ingreso * 0.0500  
        elif 100000 <= ingreso <= 1000000:  
            return ingreso * 0.0875  
        else:  
            return ingreso * 0.1000  

    def calcular_bono_propuesta3(monto):  
        if monto < 50000000:  
            return monto * 0.0015  # 0.15%  
        elif 50000000 <= monto <= 200000000:  
            return monto * 0.0040  # 0.40%  
        else:  
            return monto * 0.0050  # 0.50%  

    def calcular_bono_fijo_propuesta3(num_clientes):  
        if num_clientes < 8:  
            return 50000  
        elif 8 <= num_clientes <= 14:  
            return 120000  
        else:  
            return 350000  
    
    # Aplicar cálculos de bonos  
    df_limpio['Bono_Propuesta_1'] = df_limpio['Monto factura'].apply(calcular_bono_propuesta1)  
    df_limpio['Bono_Propuesta_2'] = df_limpio['Ingreso por operación'].apply(calcular_bono_propuesta2)  
    df_limpio['Bono_Propuesta_3'] = df_limpio['Monto factura'].apply(calcular_bono_propuesta3)  
    
    # Agregar año y mes  
    df_limpio['Año'] = df_limpio.iloc[:, 1].dt.year  
    df_limpio['Mes'] = df_limpio.iloc[:, 1].dt.month  
    
    # Resumen mensual  
    resumen_mensual = df_limpio.groupby(['Nombre Ejecutivo', 'Año', 'Mes']).agg({  
        'Monto factura': ['sum', 'count', 'mean'],  
        'Ingreso por operación': ['sum', 'mean'],  
        'Bono_Propuesta_1': 'sum',  
        'Bono_Propuesta_2': 'sum',  
        'Bono_Propuesta_3': 'sum',  
        'Nombre empresa solicitante': 'nunique'  
    }).reset_index()  

    # Aplanar columnas multinivel  
    resumen_mensual.columns = ['Ejecutivo', 'Año', 'Mes', 'Monto_Total',   
                            'Num_Operaciones', 'Monto_Promedio',  
                            'Ingreso_Total', 'Ingreso_Promedio',  
                            'Bono_1', 'Bono_2', 'Bono_3',  
                            'Num_Clientes_Unicos']  

    # Agregar bono fijo propuesta 3  
    resumen_mensual['Bono_Fijo_Prop3'] = resumen_mensual['Num_Clientes_Unicos'].apply(  
        calcular_bono_fijo_propuesta3  
    )  

    # Calcular bonos totales  
    resumen_mensual['Bono_Total_Prop3'] = resumen_mensual['Bono_3'] + resumen_mensual['Bono_Fijo_Prop3']  

    # Análisis comparativo de las tres propuestas  
    comparativo = resumen_mensual.agg({  
        'Bono_1': ['mean', 'min', 'max', 'std'],  
        'Bono_2': ['mean', 'min', 'max', 'std'],  
        'Bono_Total_Prop3': ['mean', 'min', 'max', 'std']  
    }).round(2)  

    # Análisis por ejecutivo  
    analisis_ejecutivo = resumen_mensual.groupby('Ejecutivo').agg({  
        'Bono_1': 'mean',  
        'Bono_2': 'mean',  
        'Bono_Total_Prop3': 'mean',  
        'Num_Clientes_Unicos': 'mean'  
    }).round(2)  

    # Guardar resultados  
    ruta_archivo = os.path.join(ruta_directorio, 'Analisis_Comparativo_Propuestas.xlsx')  
    
    with pd.ExcelWriter(ruta_archivo, engine='openpyxl') as writer:  
        # Resumen mensual detallado  
        resumen_mensual.to_excel(  
            writer,   
            sheet_name='Resumen_Mensual',  
            index=False  
        )  
        
        # Comparativo general  
        comparativo.to_excel(  
            writer,   
            sheet_name='Comparativo_General'  
        )  
        
        # Análisis por ejecutivo  
        analisis_ejecutivo.to_excel(  
            writer,   
            sheet_name='Analisis_por_Ejecutivo'  
        )  
        
        # Datos procesados  
        df_limpio.to_excel(  
            writer,   
            sheet_name='Datos_Procesados',  
            index=False  
        )  
    
    print(f"\nAnálisis completado. Archivo guardado en: {ruta_archivo}")  
    
    # Mostrar resumen comparativo  
    print("\nResumen Comparativo de Propuestas:")  
    print("\nPromedios mensuales:")  
    print(f"Bono Propuesta 1: ${comparativo.loc['mean', 'Bono_1']:,.2f}")  
    print(f"Bono Propuesta 2: ${comparativo.loc['mean', 'Bono_2']:,.2f}")  
    print(f"Bono Propuesta 3: ${comparativo.loc['mean', 'Bono_Total_Prop3']:,.2f}")  
    
    print("\nDesviación estándar:")  
    print(f"Propuesta 1: ${comparativo.loc['std', 'Bono_1']:,.2f}")  
    print(f"Propuesta 2: ${comparativo.loc['std', 'Bono_2']:,.2f}")  
    print(f"Propuesta 3: ${comparativo.loc['std', 'Bono_Total_Prop3']:,.2f}")  

    # Identificar la propuesta más beneficiosa en promedio  
    promedios = {  
        'Propuesta 1': comparativo.loc['mean', 'Bono_1'],  
        'Propuesta 2': comparativo.loc['mean', 'Bono_2'],  
        'Propuesta 3': comparativo.loc['mean', 'Bono_Total_Prop3']  
    }  
    mejor_propuesta = max(promedios, key=promedios.get)  
    
    print(f"\nLa {mejor_propuesta} ofrece el mejor bono promedio mensual")  

except Exception as e:  
    print(f"Error durante el análisis: {str(e)}")