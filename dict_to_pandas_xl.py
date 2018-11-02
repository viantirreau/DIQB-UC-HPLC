import pandas as pd
import numpy as np
import xlsxwriter
from pdf_to_dict import split_samples_std, read_pdf, all_std_intersection
import os.path


def linear_fit(x_vals, y_vals):
    x = np.array(x_vals)
    y = np.array(y_vals)
    return np.polyfit(x, y, 1)


def any_negative_concentration(np_fit, areas):
    m, n = np_fit
    return any(map(lambda y: (y - n) / m < 0, areas))  # y = mx + n


def linear_fit_zero_n(x_vals, y_vals):
    x = np.array(x_vals)
    y = np.array(y_vals)
    x = x[:, np.newaxis]
    return np.linalg.lstsq(x, y)[0]


def dict_to_xlsx(arch):
    base = os.path.basename(arch)
    filename = os.path.splitext(base)[0]
    txt = read_pdf(arch)
    samples, stds = split_samples_std(txt)
    if txt:
        nombres = all_std_intersection(txt)
        workbook = xlsxwriter.Workbook(f'Resultados {filename}.xlsx')
        for nombre_muestra in nombres:
            conc_negativas = False
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

            # Ver si no hay concentraciones negativas
            x_vals, y_vals = list(stds.keys()), [std[nombre_muestra] for std in
                                                 stds.values()]
            sample_vals = [sampl[nombre_muestra] for sampl in samples.values()]
            with_intercept = linear_fit(x_vals, y_vals)

            worksheet.write(f + 2, 0, "m", center)
            worksheet.write(f + 2, 1, "n", center)
            pos_m = (f + 3, 0)
            pos_n = (f + 3, 1)
            scatter = workbook.add_chart(
                {"type": "scatter"})
            scatter.set_title(
                {"name": f"Curva de calibrado {nombre_muestra}"})
            if any_negative_concentration(with_intercept, sample_vals):
                worksheet.write_array_formula(
                    f + 3, 0, f + 3, 1,
                    f"=LINEST(B3:B{f+1}, A3:A{f+1}, false, false)"
                )

                scatter.add_series({
                    'categories': f"'{nombre_muestra}'!$A$3:$A${f+1}",
                    'values': f"'{nombre_muestra}'!$B$3:$B${f+1}",
                    'trendline': {
                        'type': 'linear',
                        'intercept': 0,
                        'display_equation': True,
                        'display_r_squared': True
                    },
                })
            else:
                worksheet.write_rich_string(f + 2, 2, "R", exp, "2", center)
                worksheet.write_array_formula(
                    f + 3, 0, f + 3, 1,
                    f"=LINEST(B3:B{f+1}, A3:A{f+1}, true, false)")

                worksheet.write_formula(f + 3, 2,
                                        f"=(PEARSON(B3:B{f+1}, A3:A{f+1}))^2")
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

            scatter.set_x_axis({'name': 'Conc. mg/L'})
            scatter.set_y_axis({'name': 'Área'})
            worksheet.insert_chart("E1", scatter)

            # Escribir muestras
            worksheet.write(f + 10, 0, "Muestra", center)
            worksheet.write(f + 10, 1, "Área", center)
            worksheet.write(f + 10, 2, "mg/L", center)
            for fila, nombre in enumerate(samples, f + 11):
                worksheet.write(fila, 0, nombre)
                worksheet.write(fila, 1, samples[nombre][nombre_muestra])
                worksheet.write_formula(
                    fila, 2,
                    f"=(B{fila+1}-B{pos_n[0]+1})/A{pos_m[0]+1}",
                    cell_format=decimales3)

        workbook.close()


dict_to_xlsx('sources/Series carotenos_25.09.18.pdf')
