import pandas as pd
import numpy as np
from datetime import datetime

def crear_estado_resultados(
    ventas_brutas=500000,
    devoluciones_descuentos=100000,
    costo_ventas=250000,
    gastos_administrativos=50000,
    gastos_ventas=30000,
    ingresos_financieros=10000,
    gastos_financieros=20000,
    tasa_impuestos=0.30
):
    ventas_netas = ventas_brutas - devoluciones_descuentos
    utilidad_bruta = ventas_netas - costo_ventas
    gastos_operativos_totales = gastos_administrativos + gastos_ventas
    utilidad_operativa = utilidad_bruta - gastos_operativos_totales
    resultado_financiero_neto = ingresos_financieros - gastos_financieros
    utilidad_antes_impuestos = utilidad_operativa + resultado_financiero_neto
    impuestos = utilidad_antes_impuestos * tasa_impuestos
    utilidad_neta = utilidad_antes_impuestos - impuestos

    datos = {
        'Ventas brutas': ventas_brutas,
        'Devoluciones y descuentos': -devoluciones_descuentos,
        'Ventas netas': ventas_netas,
        'Costo de ventas': -costo_ventas,
        'Utilidad bruta': utilidad_bruta,
        'Gastos administrativos': -gastos_administrativos,
        'Gastos de ventas': -gastos_ventas,
        'Gastos operativos totales': -gastos_operativos_totales,
        'Utilidad operativa': utilidad_operativa,
        'Ingresos financieros': ingresos_financieros,
        'Gastos financieros': -gastos_financieros,
        'Resultado financiero neto': resultado_financiero_neto,
        'Utilidad antes de impuestos': utilidad_antes_impuestos,
        'Impuestos': -impuestos,
        '': np.nan,
        'UTILIDAD NETA': utilidad_neta
    }
 
    df = pd.DataFrame({
        'Concepto': list(datos.keys()),
        'Monto': list(datos.values())
    })
 
    return df


def exportar_a_excel(df, nombre_archivo="estado_resultados.xlsx"):
    with pd.ExcelWriter(nombre_archivo, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='Estado de Resultados', index=False, startrow=2)

        workbook = writer.book
        worksheet = writer.sheets['Estado de Resultados']

        formato_titulo = workbook.add_format({
            'bold': True, 'font_size': 16, 'align': 'center', 'valign': 'vcenter'
        })

        formato_encabezado = workbook.add_format({
            'bold': True, 'bg_color': '#366092', 'font_color': 'white', 
            'border': 1, 'align': 'center'
        })

        formato_moneda = workbook.add_format({
            'num_format': '$#,##0', 'border': 1
        })

        formato_moneda_negativo = workbook.add_format({
            'num_format': '$#,##0; [Red]($#,##0)', 'border': 1
        })

        formato_subtitulo = workbook.add_format({
            'bold': True, 'bg_color': '#D9E1F2', 'border': 1
        })
 
        formato_total = workbook.add_format({
            'bold': True, 'bg_color': '#FCE4D6', 'num_format': '$#,##0', 
            'border': 1, 'top': 2
        })
        worksheet.set_column('A:A', 30)
        worksheet.set_column('B:D', 15)

        worksheet.merge_range('A1:D1', 'ESTADO DE RESULTADOS', formato_titulo)

        fecha_actual = datetime.now().strftime('%d/%m/%Y')
        worksheet.write('B2', f'Fecha: {fecha_actual}')
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(2, col_num, value, formato_encabezado)

        for row_num in range(len(df)):
            for col_num in range(len(df.columns)):
                if col_num == 0:  # Columna de concepto
                    concepto = df.iloc[row_num, 0]
                    if pd.isna(concepto):
                        worksheet.write(row_num + 3, col_num, '', formato_moneda)
                    elif concepto in ['Devoluciones y descuentos',
                                      'Costos de ventas',
                                      'Gastos de ventas',
                                      'Gastos operativos totales',
                                      'Gastos financieros',
                                      'Resultado financiero neto',
                                      'Impuestos',
                                      'UTILIDAD NETA']:
                        worksheet.write(row_num + 3, col_num, concepto, formato_subtitulo)
                    else:
                        worksheet.write(row_num + 3, col_num, concepto)
                else:  # Columnas de valores
                    valor = df.iloc[row_num, col_num]
                    if pd.notna(valor):
                        if valor < 0:
                            worksheet.write(row_num + 3, col_num, valor, formato_moneda)
                        else:
                            worksheet.write(row_num + 3, col_num, valor, formato_moneda)
                    else:
                        worksheet.write(row_num + 3, col_num, '', formato_moneda)
        filas_totales = [1, 6, 7, 10, 11, 13, 15]
        for fila in filas_totales:
            for col_num in range(1, len(df.columns)):
                if col_num < len(df.columns):
                    valor = df.iloc[fila, col_num]
                    if pd.notna(valor):
                        worksheet.write(fila + 3, col_num, valor, formato_subtitulo)
        print(f"Archivo '{nombre_archivo}' exportado exitosamente!")
