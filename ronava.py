import os
import sys
import xml.etree.ElementTree as ET

import easygui as e
from openpyxl import Workbook
from openpyxl.styles import Border, Alignment, Side
from openpyxl.styles import borders
from openpyxl.cell import get_column_letter
from openpyxl.chart import BarChart, Reference
from openpyxl.writer.dump_worksheet import WriteOnlyCell
from openpyxl.styles import Style, Font


def create_formula(inicio, fin, fila):
    return '=SUM('+get_column_letter(inicio)+str(fila)+':'+get_column_letter(fin)+str(fila)+')'


def fill_cell(working_sheet, value):
    cell = WriteOnlyCell(working_sheet, value=value)
    cell.style = Style(font=Font(name='Calibri', size=11),
                       border=Border(left=Side(border_style=borders.BORDER_THIN,color='FF000000'),
                       right=Side(border_style=borders.BORDER_THIN,color='FF000000'),
                       top=Side(border_style=borders.BORDER_THIN, color='FF000000'),
                       bottom=Side(border_style=borders.BORDER_THIN,color='FF000000')))
    cell.alignment = Alignment(horizontal='center', vertical='center')

    return cell


def transform(file):
    wb = Workbook()
    ws = wb.get_sheet_by_name('Sheet')
    wb.remove_sheet(ws)
    ws = wb.create_sheet(title="Prueba")
    # ws.column_dimensions.group('A', 'D', hidden=True)
    '''
    for ws_column in range(1,10):
        ws.column_dimensions.
        ws.column_dimensions.group('A', 'I', hidden=True)
        col_letter = get_column_letter(ws_column)
        ws.column_dimensions[col_letter].bestFit = True
    '''
    tree = ET.parse(file)
    root = tree.getroot()
    tipo = root[0][0][0][0][0].attrib.get('name')

    if tipo == "datos_frx2xml":
        all = ['f'+str(i) for i in range(1, 61)]
        # Por persona
        ws.append(all)
        for fila in root.iter('datos_frx2xml'):
            to_fila = []
            for f in all:
                to_fila.append(fila.find(f).text)
            ws.append(to_fila)
        wb.save('datos_frx.xlsx')
    else:
        # Por Grupo
        dias = 1
        not_dias = 0

        # arreglo de fechas
        fechas = ['f'+str(i) for i in list(range(4, 19))]+['f'+str(i) for i in list(range(20, 36))]

        # arreglo de inasistencias
        inasistencias = ['f'+str(i) for i in list(range(38, 53))]+['f'+str(i) for i in list(range(54, 70))]

        first = True
        control = [0, 0, 0, 2]

        for f in root.iter('repinas_frx2xml'):
            control[2] = 0
            control[3] += 1
            # Formatos para la primera vez
            if first:
                first = False
                # Primera fila
                col = []

                # Fecha emision reporte
                fecha = f.find('f1').text

                if fecha is not None:
                    col.append('Fecha Emision: ')
                    col.append(fill_cell(ws, fecha))


                # Tipo de personal
                personal = f.find('f2').text

                if personal is not None:
                    col.append(fill_cell(ws, personal))

                departamento = f.find('f19').text

                if departamento is not None:
                    col.append(fill_cell(ws, departamento))

                # Agregamos la primera fila al trabajo
                ws.append(col)

                # Formato : id | personal | fechas | Total

                col = []

                # Agregamos nombre y ID
                col.append(fill_cell(ws, 'ID'))
                col.append(fill_cell(ws, 'Nombre'))

                # Agregamos las fechas
                for fecha in fechas:
                    nueva = f.find(fecha).text
                    if nueva is not None:
                        col.append(fill_cell(ws, nueva))
                    else:
                        control[0] += 1
                    control[1] += 1

                col.append(fill_cell(ws, 'Total'))
                col.append(fill_cell(ws, 'Total Sistema'))


                ws.append(col)

            col = []

            # id del obrero
            id = f.find('f36').text
            col.append(fill_cell(ws, int(id)))

            # nombre del obrero
            nombre = f.find('f37').text
            col.append(fill_cell(ws, nombre))

            # Agregamos inasistencias y cambiamos formato
            for inasistencia in inasistencias:
                lista = f.find(inasistencia).text
                try:
                    if ord(lista) == 189:
                        lista = 0.5
                    elif ord(lista) == 120:
                        lista = 1
                except:
                    pass
                finally:
                    control[2] += 1
                    if control[1] - control[0] >= control[2]:
                        col.append(fill_cell(ws, lista))

            col.append(fill_cell(ws, create_formula(3, control[1]-control[0]+2, control[3])))

            # Agregamos el total
            total = f.find('f53').text
            col.append(fill_cell(ws, int(total)))


            ws.append(col)

        # Guardamos el excel

        data = Reference(ws, min_col=32, min_row=2, max_row=38, max_col=32)
        cats = Reference(ws, min_col=1, min_row=3, max_row=38)
        chart2 = BarChart()
        chart2.type = "col"
        chart2.style = 12
        chart2.grouping = "stacked"
        chart2.title = 'Inasistencias por operario en Agosto'
        chart2.y_axis.title = 'Inasistencias'
        chart2.x_axis.title = 'Operario ID'
        chart2.add_data(data, titles_from_data=True)
        chart2.set_categories(cats)
        ws.add_chart(chart2, anchor="AH2")
        wb.save("ronava.xlsx")


yes = True

while yes:
    e.msgbox("Transformacion de archivos xml"+"\n"+"\tVersion 1.0.0"+"\n", image='images\\LogoRonava.png')

    directorio = e.diropenbox(title = "\t\tEscoger directorio", msg = "\n"+"\tSeleccione el directorio con los archivos csv")

    opciones = next(os.walk(directorio))[2]

    # Seleccionamos solo las opciones .csv

    xml = list()

    for file in opciones:
        if file[-4:] == ".xml":
            xml.append(file)

    archivos = e.multchoicebox(msg='Seleccione los archivos a transformar', title='Seleccion de archivos', choices=xml)

    e.msgbox("Iniciar transformacion")

    for file in archivos:
        transform((directorio+"%s"+file) % '\\\\')

    msg = "Do you want to continue?"
    title = "Please Confirm"
    if e.ccbox(msg, title):     # show a Continue/Cancel dialog
        pass  # user chose Continue
    else:
        sys.exit(0)     # user chose Cancel



