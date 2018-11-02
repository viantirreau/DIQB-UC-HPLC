from pdfminer.pdfparser import PDFParser, PDFDocument
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import PDFPageAggregator
from pdfminer.layout import LAParams, LTTextBox
import re
from collections import defaultdict
from functools import reduce


def read_pdf(path):
    """
    Lee el archivo pdf proporcionado en path.
    Retorna un diccionario con el formato
    {sample_name: [[molecule1, area1], [molecule2, area2]]}
    """
    # https://stackoverflow.com/a/45103154
    file_content = open(path, "rb")
    parser = PDFParser(file_content)
    doc = PDFDocument()
    parser.set_document(doc)
    doc.set_parser(parser)
    doc.initialize('')
    rsrcmgr = PDFResourceManager()
    laparams = LAParams()
    laparams.char_margin = 200  # 200  Clave para regular la separación en \n
    laparams.word_margin = 2  # 2 Parece afectar negativamente para >4
    device = PDFPageAggregator(rsrcmgr, laparams=laparams)
    interpreter = PDFPageInterpreter(rsrcmgr, device)
    dict_samples = defaultdict(list)
    last_sample = "undefined"

    # Process each page contained in the document.
    for page in doc.get_pages():
        interpreter.process_page(page)
        layout = device.get_result()
        for lt_obj in layout:
            if isinstance(lt_obj, LTTextBox):
                for subgroup in lt_obj.get_text().split("\n"):
                    if re.match(r"^( |\d|,)+$", subgroup) or subgroup.count(
                            " ") == 0:
                        continue
                    sample_name = r"^Sample Name:.+"
                    detection = r"\d+\s\w*\s( |\d|,)*"
                    subgroup = re.sub(' +', ' ', subgroup)
                    if re.match(sample_name, subgroup):
                        last_sample = re.search(r"(?<=^Sample Name: ).*",
                                                subgroup).group(0)
                        while last_sample in dict_samples:
                            last_sample += "B"

                    if re.match(detection, subgroup):
                        dict_samples[last_sample].append(
                            [re.search(
                                r"^\d+\s(\w*(?:\s\w*)*)\s( |\d|,)+",
                                subgroup).group(1),
                             subgroup.split(" ")[-4]
                             ])
    file_content.close()
    return dict_samples


def all_std_intersection(txt):
    stds = {k: v for k, v in txt.items() if re.match("^Std\d+", k)}
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
        if re.match("^Std\d+", nom_muestra):
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
    print(split_samples_std(read_pdf('sources/Series carotenos_25.09.18.pdf')))
