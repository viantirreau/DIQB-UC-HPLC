import pandas as pd
import xlsxwriter
from pdf_to_dict import split_samples_std, read_pdf, all_std_intersection
import os.path

def linear_fit(x_vals, y_vals):
    x =


def dict_to_xlsx(arch):
    base = os.path.basename(arch)
    filename = os.path.splitext(base)[0]
    txt = read_pdf(arch)
    samples, stds = split_samples_std(txt)
    if txt:
        nombres = all_std_intersection(txt)
        workbook = xlsxwriter.Workbook(f'Resultados {filename}.xlsx')
        for nombre_muestra in nombres:
            worksheet = workbook.add_worksheet(nombre_muestra)

            # Estilos
            exp = workbook.add_format()
            exp.set_font_script(1)
            center = workbook.add_format()
            center.set_align("center")
            center.set_align("vcenter")
            right = workbook.add_format()
            right.set_align("right")
            decimales3 = workbook.add_format()
            decimales3.set_num_format("0.000")

            # Escribir calibración
            worksheet.merge_range('A1:B1', "Curva de calibrado",
                                  cell_format=center)
            worksheet.write('A2', "mg/L", center)
            worksheet.write('B2', "Área", center)
            f = 0
            for fila, conc in enumerate(stds, 2):
                worksheet.write(fila, 0, conc)  # Concentración
                worksheet.write(fila, 1, stds[conc][nombre_muestra])  # Área
                f = fila
            worksheet.write(f + 2, 0, "m", right)
            worksheet.write_formula(f + 2, 1, f"=SLOPE(B3:B{f+1}, A3:A{f+1})")
            pos_m = (f + 2, 1)
            worksheet.write(f + 3, 0, "n", right)
            worksheet.write_formula(f + 3, 1,
                                    f"=INTERCEPT(B3:B{f+1}, A3:A{f+1})")
            pos_n = (f + 3, 1)

            worksheet.write_rich_string(f + 4, 0, "R", exp, "2", right)

            worksheet.write_formula(f + 4, 1,
                                    f"=(PEARSON(B3:B{f+1}, A3:A{f+1}))^2")
            scatter = workbook.add_chart(
                {"type": "scatter"})
            scatter.set_title({"name": f"Curva de calibrado {nombre_muestra}"})
            scatter.add_series({
                'categories': f"'{nombre_muestra}'!$A$3:$A${f+1}",
                'values': f"'{nombre_muestra}'!$B$3:$B${f+1}",
                'trendline': {
                    'type': 'linear',
                    'display_equation': True,
                    'display_r_squared': True
                },
            })

            scatter.set_legend({'position': 'none'})

            scatter.set_x_axis({'name': 'Conc. mg/L', 'name_layout': {
                'x': 0.9,
                'y': 0.7
            }
                                })
            scatter.set_y_axis({'name': 'Área', 'name_layout': {
                'x': 0.03,
                'y': 0.4
            }
                                })
            worksheet.insert_chart("D1", scatter)

            # Escribir muestras
            worksheet.write(f + 10, 0, "Muestra", center)
            worksheet.write(f + 10, 1, "Área", center)
            worksheet.write(f + 10, 2, "mg/L", center)
            for fila, nombre in enumerate(samples, f + 11):
                worksheet.write(fila, 0, nombre)
                worksheet.write(fila, 1, samples[nombre][nombre_muestra])
                worksheet.write_formula(
                    fila, 2,
                    f"=(B{fila+1}-B{pos_n[0]+1})/B{pos_m[0]+1}",
                    cell_format=decimales3)

        workbook.close()


dict_to_xlsx('sources/Series carotenos_25.09.18.pdf')
