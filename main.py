import pandas
import csv
import re
import os
import shutil
import in_place
import subprocess
from pylatex import Document, Section, Subsubsection, MiniPage, NoEscape, Figure, NewPage, Package, Command
from pylatex.utils import bold, italic
from wand.image import Image

sheet_file = '<CATALOG.ods>'
csv_file = '<CATALOG.csv>'
cat_folder = 'cat_descriptions'
photos_folder = 'cat_photos'
resized_photos = 'resized_photos'
categories = {}

pandas.read_excel(sheet_file,
                  sheet_name='Object recording sheet').to_csv(csv_file,
                                                              index=None, header=True, sep=',')

with open(csv_file, newline='') as file:
    reader = csv.DictReader(file)
    for row in reader:
        if row['Category'] not in categories and row['Category'] != "":
            cat = re.sub('\s', '_', row['Category'])
            try:
                text = os.popen(f" odt2txt {cat_folder}/{cat}.odt").read()
                categories[row['Category']] = text
            except FileNotFoundError as err:
                print(f"{err}")


def display_cm(cm_row):
    """
    Display "cm" string if cm_row parameters is not empty
    :param cm_row: element of <el> list
    :return: return "cm" string
    """
    if cm_row != "":
        return "cm"
    else:
        return ""


def fill_document(doc, categorie):
    """
    Function used for fill the LateX Document with information from .csv and .txt files. Also puts Photos from photos folder
    to each object of the collection.
    :param doc: Object from class pylatex.Document
    :param categorie: List of Catégories found in csv file
    :return: Document Object in LateX
    """
    id_count = 1
    doc.packages.append(Package('placeins'))
    for cat in categorie:
        doc.append(NewPage())
        with doc.create(Section(f'{cat}', numbering=1)):
            doc.append(categories[cat])
            with open(csv_file, newline='') as file_2:
                csv_reader = csv.DictReader(file_2)
                for row in csv_reader:
                    if row['Category'] == cat and row['Zone'] != "":
                        el = [row['ID_Object'].zfill(6), row['Zone'], row['Building'],
                              row['Street_CityGate'], row['Tayara'], row['Material'],
                              row['Ref_Photo'], row['Location'], row['Conservation required'],
                              row['Nbr_frag'], row['Length (cm)'], row['Width (cm)'], row['Height (cm)'],
                              row['Thickness (cm)'], row['Diameter (cm)'],
                              row['Identification'], row['Description'], row['Comments'], row['SU'], row['Room']]
                        with doc.create(Subsubsection(bold(f"{id_count} - {el[15]}"), numbering=0)):
                            doc.append(italic(f'{el[16]} {el[17]}\n\n'))
                            with doc.create(MiniPage(width=NoEscape('10cm'))):
                                doc.append(bold(f'N° : '))
                                doc.append(f"OBJ{re.sub('.[0-9]$', '', el[0])}\n")
                                doc.append(bold('Place of discovery : '))
                                if el[1] != "":
                                    zone = f"{el[1]}"
                                    doc.append(f'{zone}')
                                if el[2] != "":
                                    building = f"{el[2]}"
                                    doc.append(f'_B{building}')
                                if el[3] != "":
                                    reg = "CG.+"
                                    street = f"{el[3]}"
                                    if re.match(reg, el[3]):
                                        doc.append(f'_{street}')
                                    else:
                                        doc.append(f'_S{street}')
                                if el[4] != "":
                                    tayara = f"{re.sub('.[0-9]$', '', el[4])}"
                                    doc.append(f'_T{tayara}')
                                if el[19] != "":
                                    room = f"{el[19]}"
                                    doc.append(f'_{room}')
                                if el[18] != "":
                                    su = f"{re.sub('.[0-9]$', '', el[18])}"
                                    doc.append(f'_SU{su}')
                                doc.append(f'\n')
                                doc.append(f'')
                                doc.append(bold('Material : '))
                                doc.append(f'{el[5]}\n')
                                doc.append(bold(f'Inventory number : '))
                                doc.append(f"{re.sub('.[0-9]$', '', el[6])}\n")
                                doc.append(bold(f'Location : '))
                                doc.append(f'{el[7]}\n')
                                doc.append(bold(f'Conservation required : '))
                                doc.append(f'{el[8]}\n')
                            with doc.create(MiniPage(width=NoEscape('10cm'))):
                                doc.append(bold(f'Number of objects/fragments : '))
                                doc.append(f'{el[9]}\n')
                                doc.append(bold(f'Length : '))
                                doc.append(f'{el[10]} {display_cm(el[10])}\n')
                                doc.append(bold(f'Width : '))
                                doc.append(f'{el[11]} {display_cm(el[11])}\n')
                                doc.append(bold(f'Heigth : '))
                                doc.append(f'{el[12]} {display_cm(el[12])}\n')
                                doc.append(bold(f'Thickness : '))
                                doc.append(f'{el[13]} {display_cm(el[13])}\n')
                                doc.append(bold(f'Diameter : '))
                                doc.append(f'{el[14]} {display_cm(el[14])}\n')
                            with doc.create(Figure(position='h!')) as fig:
                                for photo in os.listdir(photos_folder):
                                    if photo == f"{el[6]}.JPG":
                                        img = Image(filename=f"{photos_folder}/{photo}")
                                        img.resize(int(img.width / 3), int(img.height / 3))
                                        img.save(filename=f"resized_photos/{photo}")
                                        fig.add_image(f"{resized_photos}/{photo}", width='150px')
                                        doc.append(italic(f"\n{id_count}-{el[15]}-OBJ{re.sub('.[0-9]$', '', el[0])}\n"))
                        doc.append(Command('FloatBarrier'))
                        id_count += 1


if __name__ == '__main__':
    wd = os.getcwd()
    os.mkdir(resized_photos)
    geometry_options = {"tmargin": "2cm", "lmargin": "2cm"}
    doc = Document('rapport', geometry_options=geometry_options)
    fill_document(doc, categories)
    doc.generate_tex()
    with in_place.InPlace('rapport.tex') as file:
        for line in file:
            line = line.replace('photos/{', 'photos/')
            line = line.replace('}.JPG', '.JPG')
            file.write(line)
    cmd = subprocess.Popen(['latexmk', '--pdf', '--interaction=nonstopmode',
                            f'{wd}/rapport.tex'],
                           stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
    out, err = cmd.communicate()
    if err is not None:
        print(err)
    shutil.rmtree(resized_photos)
    os.remove(csv_file)
