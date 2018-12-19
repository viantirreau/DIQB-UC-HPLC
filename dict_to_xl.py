import numpy as np
import xlsxwriter
from pdf_to_dict import read_pdf
import os.path
import time
import re


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


def dict_to_xlsx(arch, save_path, sgn_progress=None, report_od=False):
    """
    :param arch: Path to PDF to be read
    :param save_path: Path for .xlsx file ti be written to
    :param sgn_progress: pyqtSignal for reporting progress to GUI
    :param report_od: Whether to include OD and biomass-based yield calculations
    :return:    0: File processed and saved successfully
                1: File lacks standard areas for molecule concentration
                2: File is locked
                3: Error in processing
    """
    base = os.path.basename(arch)
    filename = os.path.splitext(base)[0]
    try:
        result, molecule_names = read_pdf(arch, sgn_progress)
    except PermissionError:
        return 2
    # if not result.get("standards"):
    #     return 1
    # try:
    #     samples, stds = split_samples_std(txt)
    # except ValueError:
    #     return 1
    if all(i in result for i in ("standards", "samples", "int_standards")):
        samples = result["samples"]
        standards = result["standards"]
        int_standards = result["int_standards"]
        try:
            workbook = xlsxwriter.Workbook(
                os.path.join(save_path, f'Resultados {filename}.xlsx'))
        except PermissionError:
            return 2

        # Styles
        exp = workbook.add_format()
        exp.set_font_script(1)
        center = workbook.add_format()
        center.set_align("center")
        center.set_align("vcenter")
        center_ = workbook.add_format()
        center_.set_align("center")
        center_.set_align("vcenter")
        center_.set_border(1)
        right = workbook.add_format()
        right.set_align("right")
        decimals3 = workbook.add_format()
        decimals3.set_num_format("0.000")
        decimals3.set_align("center")
        decimals3.set_align("vcenter")
        decimals3.set_border(1)
        light_blue = workbook.add_format()
        light_blue.set_bg_color("#CCCCFF")
        light_blue.set_align("center")
        light_blue.set_align("vcenter")
        light_blue.set_border(1)
        red = workbook.add_format()
        red.set_bg_color("#FFAAAA")
        red.set_align("center")
        red.set_align("vcenter")

        for molecule in molecule_names:
            sanitized_molecule = re.sub(r"[:\\\[\]*?/]+", "", molecule)
            worksheet = workbook.add_worksheet(sanitized_molecule)
            worksheet.set_column(0, 0, 3)
            worksheet.set_row(0, 9)

            # If there are no standards or less than two points
            # needed for the fit, do not report concentrations, only areas
            if molecule not in standards or len(
                    standards.get(molecule, {}).keys()) < 2:
                worksheet.merge_range('B2:F2',
                                      "DATOS DE CALIBRADO NO ENCONTRADOS",
                                      cell_format=red)
                worksheet.write('B4', "Muestra", center_)
                worksheet.write('C4', "Área", center_)
                for row, smpl in enumerate(samples.keys(), 4):
                    area = samples[smpl].get(molecule, 0)
                    worksheet.write(row, 1, smpl, center_)
                    worksheet.write_number(row, 2, area, center_)

            else:
                # Write calibration
                worksheet.merge_range('B2:C2', f"STD {molecule}",
                                      cell_format=center_)
                worksheet.merge_range('B3:C3', "Curva de calibrado",
                                      cell_format=center_)
                worksheet.write('B4', "Conc.", center_)
                worksheet.write('C4', "Área", center_)
                increasing = sorted(list(standards[molecule].keys()))
                r = 0
                for r_, conc in enumerate(increasing, 4):
                    r = r_
                    area = standards[molecule][conc]
                    worksheet.write_number(r, 1, conc, center_)
                    worksheet.write_number(r, 2, area, center_)
                linear_fit_row = r + 3

                worksheet.write(linear_fit_row - 1, 1, "m", center_)
                worksheet.write(linear_fit_row - 1, 2, "n", center_)

                # Filter only greater than 0 areas
                sample_areas = filter(lambda x: x > 0,
                                      [samples[spl_name][molecule] for spl_name
                                       in samples if molecule in
                                       samples[spl_name]])

                scatter = workbook.add_chart(
                    {"type": "scatter"})
                scatter.set_title(
                    {"name": f"Curva de calibrado {molecule}"})

                neg_conc = False
                # Check for non-negative concentrations
                x_vals, y_vals = increasing, [standards[molecule][i] for i
                                              in
                                              increasing]
                try:
                    with_intercept = linear_fit(x_vals, y_vals)
                    neg_conc = any_negative_concentration(with_intercept,
                                                          sample_areas)
                    if neg_conc:
                        print(f"Negative concentration detected for {molecule}",
                              "Parameters ",
                              f"m:{with_intercept[0]}, n:{with_intercept[1]}")

                except np.linalg.linalg.LinAlgError as err:
                    neg_conc = True
                    print(x_vals, y_vals)

                if neg_conc:

                    worksheet.write_array_formula(
                        linear_fit_row, 1, linear_fit_row, 2,
                        f"=LINEST(C5:C{r+1}, B5:B{r+1}, false, false)",
                        cell_format=center_
                    )

                    scatter.add_series({
                        'categories': f"'{molecule}'!$B$5:$B${r+1}",
                        'values': f"'{molecule}'!$C$5:$C${r+1}",
                        'trendline': {
                            'type': 'linear',
                            'intercept': 0,
                            'display_equation': True,
                            'display_r_squared': True
                        },
                    })
                else:
                    worksheet.write_rich_string(linear_fit_row - 1, 3, "R", exp,
                                                "2", center_)
                    worksheet.write_array_formula(
                        r + 3, 1, r + 3, 2,
                        f"=LINEST(C5:C{r+1}, B5:B{r+1}, true, false)",
                        cell_format=center_)

                    worksheet.write_formula(r + 3, 3,
                                            f"=(PEARSON("
                                            f"C5:C{r+1}, B5:B{r+1}))^2",
                                            cell_format=center_)
                    scatter.add_series({
                        'categories': f"'{molecule}'!$B$5:$B${r+1}",
                        'values': f"'{molecule}'!$C$5:$C${r+1}",
                        'trendline': {
                            'type': 'linear',
                            'display_equation': True,
                            'display_r_squared': True
                        },
                    })

                scatter.set_legend({'position': 'none'})

                scatter.set_x_axis({'name': 'Conc. [unid]'})
                scatter.set_y_axis({'name': 'Área'})
                worksheet.insert_chart("J2", scatter)

                sample_area_row = linear_fit_row + 3
                worksheet.write(sample_area_row - 1, 1, "Muestra", center_)
                worksheet.write(sample_area_row - 1, 2, "Área", center_)
                worksheet.write(sample_area_row - 1, 3, "Conc.", center_)
                if report_od:
                    worksheet.write(sample_area_row - 1, 4, "OD", center_)
                    worksheet.write(sample_area_row - 1, 5, "mL vial", center_)
                    worksheet.write(sample_area_row - 1, 6, "g Biomasa",
                                    center_)
                    worksheet.write(sample_area_row - 1, 7, "mg/g", center_)

                for row, spl_name in enumerate(samples, sample_area_row):
                    area = samples[spl_name].get(molecule, 0)
                    worksheet.write(row, 1, spl_name, center_)
                    worksheet.write_number(row, 2, area, center_)
                    worksheet.write_formula(
                        row, 3,
                        f"=IF(C{row+1}=0,0,"
                        f"(C{row+1}-C{linear_fit_row+1})/B{linear_fit_row+1})",
                        cell_format=decimals3)
                    if report_od:
                        worksheet.write_number(row, 4, 20, light_blue)
                        worksheet.write_number(row, 5, 1, light_blue)
                        worksheet.write_formula(
                            row, 6, f"=0.4*E{row+1}*F{row+1}/1000", center_)
                        worksheet.write_formula(
                            row, 7, f"=D{row+1}/(G{row+1}*1000)", center_)

                    """
                    for sample_name in samples:
                        # Ver si no hay concentraciones negativas
                        
                        sample_vals = [sampl[sample_name] for sampl in samples.values() if
                                       sample_name in sampl]
                        
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
                            """
        try:
            workbook.close()
        except PermissionError:
            return 2
        return 0
    return 3


if __name__ == '__main__':
    c = 0
    t0 = time.time()
    for file in os.listdir(os.getcwd()):
        if file.endswith(".pdf"):
            print(file)
            if "carot" in file.lower():
                dict_to_xlsx(
                    file, "C:/Users/Victor/Desktop/res/", None, True)
            else:
                dict_to_xlsx(
                    file, "C:/Users/Victor/Desktop/res/",
                    None)
            c += 1
    t_tot = time.time() - t0
    print(f"Procesados {c} archivos en {t_tot:.4f} s")
