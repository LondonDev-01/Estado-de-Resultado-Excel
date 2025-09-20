from read_excel import read as r
from writer_test import writer as w


def menu():

    return


def main():
    print("Ingresa los datos: ")

    ventas_brutas = input("Ventas Brutas: ")
    ventas_brutas = 0 if ventas_brutas == '' else int(ventas_brutas)

    devoluciones_descuentos = input("Devoluciones descuentos: ")
    devoluciones_descuentos = 0 if devoluciones_descuentos == '' else int(devoluciones_descuentos)

    costo_ventas = input("Costos Ventas: ")
    costo_ventas = 0 if costo_ventas == '' else int(costo_ventas)

    gastos_administrativos = input("Gastos administrativos: ")
    gastos_administrativos = 0 if gastos_administrativos == '' else int(gastos_administrativos)

    gastos_ventas = input("Gastos ventas: ")
    gastos_ventas = 0 if gastos_ventas == '' else int(gastos_ventas)

    ingresos_financieros = input("Ingresos financieros: ")
    ingresos_financieros = 0 if ingresos_financieros == '' else int(ingresos_financieros)

    gastos_financieros = input("Gastos financieros: ")
    gastos_financieros = 0 if gastos_financieros == '' else int(gastos_financieros)

    tasa_impuestos = input("Tasa de impuestos: ")
    tasa_impuestos = 0.30 if tasa_impuestos == '' else float(tasa_impuestos)

    df = w.crear_estado_resultados(
        ventas_brutas,
        devoluciones_descuentos,
        costo_ventas,
        gastos_administrativos,
        gastos_ventas,
        ingresos_financieros,
        gastos_financieros,
        tasa_impuestos,
    )
    w.exportar_a_excel(df, "prueba.xlsx")
    return


if __name__ == "__main__":
    main()
