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


def dict_to_xlsx(arch, save_path):
    """
    :param arch: Path to PDF to be read
    :param save_path: Path for .xlsx file ti be written to
    :return:    0: File processed and saved successfully
                1: File lacks standard areas for molecule concentration
                2: File is locked
    """
    base = os.path.basename(arch)
    filename = os.path.splitext(base)[0]
    try:
        txt = read_pdf(arch)
    except PermissionError:
        return 2
    try:
        samples, stds = split_samples_std(txt)
    except ValueError:
        return 1
    if txt:
        names = all_std_intersection(txt)
        try:
            workbook = xlsxwriter.Workbook(
                os.path.join(save_path, f'Resultados {filename}.xlsx'))
        except PermissionError:
            return 2
        for sample_name in names:
            worksheet = workbook.add_worksheet(sample_name)

            # Styles
            exp = workbook.add_format()
            exp.set_font_script(1)
            center = workbook.add_format()
            center.set_align("center")
            center.set_align("vcenter")
            right = workbook.add_format()
            right.set_align("right")
            decimals3 = workbook.add_format()
            decimals3.set_num_format("0.000")
            light_blue = workbook.add_format()
            light_blue.set_bg_color("#CCCCFF")

            # Write calibration
            worksheet.merge_range('A1:B1', "Curva de calibrado",
                                  cell_format=center)
            worksheet.write('A2', "mg/L", center)
            worksheet.write('B2', "Área", center)
            r = 0
            for row, conc in enumerate(stds, 2):
                worksheet.write(row, 0, conc)  # Concentration
                worksheet.write(row, 1, stds[conc][sample_name])  # Area
                r = row

            # Ver si no hay concentraciones negativas
            x_vals, y_vals = list(stds.keys()), [std[sample_name] for std in
                                                 stds.values()]
            sample_vals = [sampl[sample_name] for sampl in samples.values() if
                           sample_name in sampl]
            with_intercept = linear_fit(x_vals, y_vals)

            worksheet.write(r + 2, 0, "m", center)
            worksheet.write(r + 2, 1, "n", center)
            pos_m = (r + 3, 0)
            pos_n = (r + 3, 1)
            scatter = workbook.add_chart(
                {"type": "scatter"})
            scatter.set_title(
                {"name": f"Curva de calibrado {sample_name}"})
            if any_negative_concentration(with_intercept, sample_vals):
                worksheet.write_array_formula(
                    r + 3, 0, r + 3, 1,
                    f"=LINEST(B3:B{r+1}, A3:A{r+1}, false, false)"
                )

                scatter.add_series({
                    'categories': f"'{sample_name}'!$A$3:$A${r+1}",
                    'values': f"'{sample_name}'!$B$3:$B${r+1}",
                    'trendline': {
                        'type': 'linear',
                        'intercept': 0,
                        'display_equation': True,
                        'display_r_squared': True
                    },
                })
            else:
                worksheet.write_rich_string(r + 2, 2, "R", exp, "2", center)
                worksheet.write_array_formula(
                    r + 3, 0, r + 3, 1,
                    f"=LINEST(B3:B{r+1}, A3:A{r+1}, true, false)")

                worksheet.write_formula(r + 3, 2,
                                        f"=(PEARSON(B3:B{r+1}, A3:A{r+1}))^2")
                scatter.add_series({
                    'categories': f"'{sample_name}'!$A$3:$A${r+1}",
                    'values': f"'{sample_name}'!$B$3:$B${r+1}",
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
            worksheet.write(r + 10, 0, "Muestra", center)
            worksheet.write(r + 10, 1, "Área", center)
            worksheet.write(r + 10, 2, "mg/L", center)
            worksheet.write(r + 10, 3, "OD", center)
            worksheet.write(r + 10, 4, "mL vial", center)
            worksheet.write(r + 10, 5, "g Biomasa", center)
            worksheet.write(r + 10, 6, "mg/g", center)

            for row, nombre in enumerate(samples, r + 11):
                worksheet.write(row, 0, nombre)
                worksheet.write(row, 1, samples[nombre].get(sample_name, 0))
                worksheet.write_formula(
                    row, 2,
                    f"=(B{row+1}-B{pos_n[0]+1})/A{pos_m[0]+1}",
                    cell_format=decimals3)
                worksheet.write_number(row, 3, 20, light_blue)
                worksheet.write_number(row, 4, 1, light_blue)
                worksheet.write_formula(
                    row, 5, f"=0.4*D{row+1}*E{row+1}/1000"
                )
                worksheet.write_formula(
                    row, 6, f"=C{row+1}/(F{row+1}*1000)"
                )
        try:
            workbook.close()
        except PermissionError:
            return 2
        return 0


if __name__ == '__main__':
    dict_to_xlsx(
        'C:/Users/Victor/Documents/iPre/Script HPLC/sources/Series '
        'carotenos_16.10.18_IS.pdf', "C:/Users/Victor/Desktop/")