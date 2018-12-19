from pdfminer.pdfparser import PDFParser, PDFDocument
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter, resolve1
from pdfminer.converter import PDFPageAggregator, TextConverter
from pdfminer.layout import LAParams, LTTextBox
import re
import io
from functools import reduce
import os
import time

# RegEx compiles
F = re.MULTILINE | re.IGNORECASE
EMPTY = re.compile(r"^( |\d|,)+$", F)
SAMPLE_NAME = re.compile(r"(?:Sample Name|Nombre de la muestra):\s*(.+)$", F)
BLANK = re.compile(r"(?:Sample Name|Nombre de la muestra):\s*BCO", F)
DETECTION = re.compile(r"\d+\s\w*\s( |\d|,)*", F)
INNER_DET = re.compile(r"^\d+\s(\w*(?:\s\w*)*)\s( |\d|,)+", F)
VERTICAL_TEXT = re.compile(r"\s((?!0)(\S\s+)+)0 (\d+ +)+", F)
STD_SAMPLE_NAME = re.compile(r"(?:Sample Name|Nombre de la muestra):\s*St", F)
STD_VIAL_TYPE = re.compile(r"(Vial Type|Tipo):\s*std", F)


def read_pdf(path, report_progress_sgn=None):
    """
    Reads the pdf file given in path and reports progress if a signal is given

    :return: Dict with format
    {"samples":
        {"sample_name_1":
            {"molecule_1": area_m1, "molecule_2": area_m2}
        }
     "standards":
        {"molecule_1":
            {concentration_1: area_1, concentration_2: area_2}
        }
     "int_standards":
        {"sample_name_1":
            {"molecule_1": area_m1}
         "sample_name_2":
            {"molecule_1": area_m1}
        }
    }
    """
    # https://stackoverflow.com/a/45103154
    with open(path, "rb") as file_content:
        parser = PDFParser(file_content)
        doc = PDFDocument(caching=False)
        parser.set_document(doc)
        doc.set_parser(parser)
        doc.initialize('')
        rsrcmgr = PDFResourceManager()
        laparams = LAParams()
        laparams.char_margin = 200  # 200  Seems important for \n separation
        laparams.word_margin = 2  # 2 Seems to break for >4
        # device = PDFPageAggregator(rsrcmgr, laparams=laparams)
        retstr = io.StringIO()
        device = TextConverter(rsrcmgr, retstr, laparams=laparams)
        interpreter = PDFPageInterpreter(rsrcmgr, device)
        sample_names_set = set()
        molecule_names_set = set()
        sample_types = {}
        last_sample_name = None
        processed = {"samples": {},
                     "standards": {},
                     "int_standards": {}}

        tot_pages = resolve1(doc.catalog["Pages"])["Count"]
        print("Total:", tot_pages)
        # Process each page contained in the document.
        for n, page in enumerate(doc.get_pages(), 1):
            if report_progress_sgn:
                report_progress_sgn.emit(n / tot_pages)
            else:
                pass
                # print("-" * 20 + "\n", n, '\n')
            interpreter.process_page(page)
            text = retstr.getvalue()
            is_standard = False
            sample_name = None

            if BLANK.search(text):
                retstr.truncate(0)
                retstr.seek(0)
                continue

            if STD_SAMPLE_NAME.search(text) or STD_VIAL_TYPE.search(text):
                is_standard = True

            found_sample_name = SAMPLE_NAME.findall(text)
            if len(found_sample_name) > 0:
                if found_sample_name[0] != "":
                    sample_name = found_sample_name[0]

            if sample_name is None:
                continue  # Probably an almost empty page
            if sample_name == "":
                sample_name = "Sin nombre"

            if sample_name not in sample_names_set:
                sample_names_set.add(sample_name)
            # If the same sample names are separated by more than a page
            # (with recognized sample names), they should be treated as
            # different samples, hence the name change.
            elif last_sample_name != sample_name:
                sample_name += "_1"
                sample_names_set.add(sample_name)

            last_sample_name = sample_name

            # Deal with multi-page standards
            if sample_name not in sample_types:
                if is_standard:
                    sample_types[sample_name] = "standard"
                else:
                    sample_types[sample_name] = "sample"
            elif sample_types[sample_name] == "standard":
                is_standard = True

            # vert = VERTICAL_TEXT.findall(text)
            # if vert:
            #     if vert[0]:
            #         print(
            #             vert[0][0][::-1].replace("\n\n", " ").replace("\n", ""))
            data_section = False
            started_saving_data = False
            col_names = None
            for line in text.split("\n"):
                if all(i in line for i in ("Area", "Name")):
                    col_names = [i for i in line.split(" ") if not i.isdigit()]
                    data_section = True
                    continue
                if data_section:
                    if line == "" and started_saving_data:
                        break

                    line = re.sub(' +', ' ', line).strip()
                    vals = [i for i in line.split(" ") if i != ""]
                    # If already started saving data
                    # and this line is blank, break to save cycles
                    if len(vals) < 3 and started_saving_data:
                        break
                    # Not a full row, not enough data available
                    if len(vals) < len(col_names):
                        continue
                    if len(vals) > len(col_names):
                        # Account for a longer-than-expected Name field
                        # with spaces (which are the delimiters)
                        started_saving_data = True
                        col_names_ = col_names[:]
                        # print(line)
                        positive_name_col_idx = col_names_.index("Name")
                        reverse_name_col_idx = positive_name_col_idx - len(
                            col_names_) + 1
                        line_dict = {}
                        # Add elements from the left of the Name column
                        for positive_idx in range(positive_name_col_idx):
                            col_name, value = col_names_.pop(0), vals.pop(0)
                            line_dict[col_name] = value
                        # Add elements by reverse indexing to the right of Name
                        for negative_idx in range(reverse_name_col_idx, 0):
                            col_name = col_names_.pop(negative_idx)
                            value = vals.pop(negative_idx)
                            line_dict[col_name] = value
                        # The only values left should be the Name with spaces
                        if len(vals) > 0 and len(col_names_) > 0:
                            line_dict["Name"] = " ".join(vals)
                    else:
                        started_saving_data = True
                        line_dict = {k: v for k, v in zip(col_names, vals)}
                    # By now, line_dict should be filled with the sought values
                    if "Name" in line_dict and "Area" in line_dict:
                        name, area = line_dict["Name"], line_dict["Area"]
                        molecule_names_set.add(name)
                        try:
                            area = float(area.replace(",", "."))
                        except ValueError:
                            continue

                        # Standard molecule
                        if is_standard:
                            conc = line_dict.get("Conc")
                            if conc is None:
                                # Fallback: use the last column blindly
                                conc = list(line_dict.values())[-1]
                            try:
                                conc = float(conc.replace(",", "."))
                            except ValueError:
                                continue
                            if name in processed["standards"]:
                                processed["standards"][name][conc] = area
                            else:
                                processed["standards"][name] = {conc: area}
                        # Assume Internal Standard
                        elif line_dict["Name"].lower().endswith("_is"):
                            if sample_name in processed["int_standards"]:
                                processed["int_standards"][sample_name][
                                    name] = area
                            else:
                                processed["int_standards"][sample_name] = {
                                    name: area}
                        # Otherwise it is a sample
                        else:
                            if sample_name in processed["samples"]:
                                processed["samples"][sample_name][name] = area
                            else:
                                processed["samples"][sample_name] = {name: area}

            retstr.truncate(0)
            retstr.seek(0)
            """
            layout = device.get_result()
            for lt_obj in layout:
                if isinstance(lt_obj, LTTextBox):
                    for subgroup in lt_obj.get_text().split("\n"):
                        if EMPTY.match(subgroup) or subgroup.count(
                                " ") == 0:
                            continue
                        detection = r"\d+\s\w*\s( |\d|,)*"
                        subgroup = re.sub(' +', ' ', subgroup)
                        if SAMPLE_NAME.match(subgroup):
                            last_sample = re.search(r"(?<=^Sample Name: ).*",
                                                    subgroup).group(0)
                            while last_sample in dict_samples:
                                last_sample += "B"

                        if DETECTION.match(subgroup):
                            inner_detection = INNER_DET.search(subgroup)
                            if inner_detection:
                                dict_samples[last_sample].append(
                                    [inner_detection.group(1),
                                     subgroup.split(" ")[-4]
                                     ])
                                     
                """
    # print(processed)
    # print(sample_types)
    # print(molecule_names_set)
    # return tot_pages
    return processed, sorted(list(molecule_names_set))


def all_std_intersection(txt):
    stds = {k: v for k, v in txt.items() if re.match("^Std\d+", k)}
    if not stds:
        raise ValueError
    return reduce(set.intersection, [{i[0] for i in j} for j in stds.values()])


def split_samples_std(txt):
    """
    Lee un archivo mediante la función read_pdf y transforma el dict de su
    output, en formato nombre_muestra: lista_detecciones a 2 dict, uno
    para las muestras, en el mismo formato, y otro para los standard para la
    curva de calibrado. El formato del primero es
    {sample_name: [[molecule1, area1], [molecule2, area2]]}.
    El formato para el segundo es
    {concentracion_std: [[molec1, area1], [molec2, area2]]}
    """
    stds = {}
    nombres_stds = set()
    for nom_muestra in txt:
        if re.match("^Std\d+", nom_muestra, flags=re.IGNORECASE):
            nombres_stds.add(nom_muestra)
            masa_por_litro = re.search(r"(?<=^Std)(\d+)", nom_muestra).group(1)
            if masa_por_litro.isdigit():
                stds[masa_por_litro] = txt[nom_muestra]
    all_std_intersect = all_std_intersection(txt)
    std_filtrado = {
        int(k): {j[0]: float(j[1]) for j in v if j[0] in all_std_intersect}
        for k, v in stds.items()}
    samples_filtrado = {
        k: {j[0]: float(j[1]) for j in v if j[0] in all_std_intersect} for
        k, v in txt.items() if k not in nombres_stds}
    # Sort nombres
    samples_filtrado = {k: samples_filtrado[k] for k in
                        sorted(samples_filtrado)}
    return samples_filtrado, std_filtrado


if __name__ == '__main__':
    # c = 0
    # t0 = time.time()
    # for file in os.listdir(os.getcwd()):
    #     if file.endswith(".pdf"):
    #         print(file)
    #         c += read_pdf(file, None)
    # t_tot = time.time() - t0
    # print(f"Procesadas {c} páginas en {t_tot:.4f} s -> {c/t_tot:.4f} p/s")
    read_pdf("Series Azucares_30.11.18.pdf", None)
