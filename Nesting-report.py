import os
from datetime import datetime
import configparser
import ewd
from pdfkit import pdfkit
from ewd import groups
from company import gdb
from company import dlg
from company import nest
from company.gdb import cad
from company.gdb import view
from sclcore import do_debug
from sclcore import execute_command_bool as exec_bool


class ReportSheet():
    def __init__(self, sheet, mat_leftover, mat_reusable, area, count, img_path, total_area, total_reusable, total_garbage):
        self.sheet = sheet
        self.mat_leftover = mat_leftover
        self.mat_reusable = mat_reusable
        self.area = area
        self.count = count
        self.img_path = img_path
        self.total_area = total_area
        self.total_reusable = total_reusable
        self.total_garbage = total_garbage



def nesting_report():
    """
    The program loads settings from an .ini file, creates a new folder, and generates an HTML report with applied CSS. 
    It includes detailed information about each sheet, calculates individual and total efficiency metrics, and then converts the final HTML report into a PDF format.
    """

    do_debug()
    rotate, general_folder, nice_design, reports_pdfs_together, only_measures_BW, divide_material = read_config_ini()

    project_name = get_or_create_project_name()

    folder = make_or_delete_folder(general_folder, project_name)

    report_file_path = create_report_file_path(folder)

    img_ext = ".jpg"
    logo = "C:\...\logo.png"

    set_view_and_shading(nice_design)


    materials_dict = sort_for_material()
    counter_for_full_pdf = 0

    for material_and_thickness, sheets_values in materials_dict.items():
        material, thickness = material_and_thickness  #extract material and thickness from the key

        #html with a new name
        if divide_material:
            report_file_path = f"{folder}\\{material}_{thickness}.html"
            try:
                os.makedirs(os.path.dirname(report_file_path), exist_ok=True)
            except OSError as e:
                dlg.output_box(f"Fehler beim Ordner erstellen in {os.path.dirname(report_file_path)}")

            project_name_mat_thick = f"{project_name}_{material}_{thickness}".replace('.ewd', '')
            create_report(report_file_path, project_name_mat_thick, folder, img_ext, logo, nice_design, rotate, reports_pdfs_together, divide_material, sheets_values, materials_dict, 0)

        else: #if not divide_material:
            sheets = nest.get_sheets()
            create_report(report_file_path, project_name, folder, img_ext, logo, nice_design, rotate, reports_pdfs_together, divide_material, sheets, materials_dict, counter_for_full_pdf)
            counter_for_full_pdf +=1



def run_config():
    config.run_config()


def read_config_ini():
    try:
        ini_path = ewd.explode_file_path(r'%MACHPATH%\script\config.ini')

        config = configparser.ConfigParser()
        config.read(ini_path)

        rotate = config.get('Einstellung', 'rotate') #if True, rotate
        rotate = False if rotate == "0" or rotate =="False" else True

        #path for the folder
        general_folder =  config.get('Pfad', 'report_pfad')
        #create new unique folder
        general_folder = os.path.join(general_folder, 'Report_new')
        #regardless of the user choice, there will be  created a new folde, that will only have report files, that are safe to delete

        nice_design = config.get('Druckeinstellungen', 'nice_design') #if True, nice_design
        nice_design = False if nice_design == "0" or nice_design =="False" else True

        reports_pdfs_together = config.get('Druckeinstellungen', 'reports_pdfs_together') #if True, reports_pdfs_together
        reports_pdfs_together = False if reports_pdfs_together == "0" or reports_pdfs_together =="False" else True

        only_measures_BW = config.get('Druckeinstellungen', 'only_measures_BW') #if True, only_measures_BW
        only_measures_BW = False if only_measures_BW == "0" or only_measures_BW =="False" else True

        divide_material = config.get('Druckeinstellungen', 'divide_material') #if True, divide the report depending on the material
        divide_material = False if divide_material == "0" or divide_material =="False" else True

        return rotate, general_folder, nice_design, reports_pdfs_together, only_measures_BW, divide_material

    except FileNotFoundError:
        dlg.output_box('Fehler: Die Konfigurationsdatei "config.ini" wurde nicht gefunden. Bitte überprüfen Sie den Dateipfad.')
    except configparser.NoSectionError:
        dlg.output_box("Fehler: Die Sektion 'Pfad' fehlt in der Konfigurationsdatei.")
    except configparser.NoOptionError:
        dlg.output_box("Fehler: Die Option 'report_pfad' fehlt in der Sektion 'Pfad'.")
    except KeyError as e:
        dlg.output_box(f"Konfigurationsparameter nicht gefunden: {e}")
    except ValueError as e:
        dlg.output_box(f"Ungültiger Wert für den Konfigurationsparameter: {e}")
    except Exception as e:
        dlg.output_box(f"Ein unerwarteter Fehler ist aufgetreten: {e}")


def get_or_create_project_name():
    project_name = ewd.get_project_name() #get the name of the opened ewd project
    if not project_name.endswith (".ewd"):
        project_name = datetime.datetime.now().strftime("%Y%m%d_%H-%M")
        ewd.save_project(ewd.explode_file_path(f"%TEMPPATH%//{project_name}.ewd"))
        project_name = f"{project_name}.ewd"
    return project_name



def remove_existing_folder_with_same_name(folder):
    try:
        if os.path.exists(folder):
            dlg.output_box(f"Der Ordner '{folder}' und sein Inhalt werden gelöscht")
            #add ok / cancel - config
            for root, dirs, files in os.walk(folder, topdown=False):
                for name in files:
                    os.remove(os.path.join(root, name))
                for name in dirs:
                    os.rmdir(os.path.join(root, name))
            os.rmdir(folder)
        else:
            dlg.output_box(f"Der Ordner '{folder}' existiert nicht.")
    except Exception as e:
        dlg.output_box(f"Ein Fehler ist aufgetreten: {e}")

def create_report_file_path(folder):
    report_file_path = f'{folder}\\report.html'
    try:
        os.makedirs(os.path.dirname(report_file_path), exist_ok=True)
    except OSError as e:
        dlg.output_box(f"Fehler beim Ordner erstellen in {os.path.dirname(report_file_path)}")
    return report_file_path



def set_view_and_shading(nice_design):
    #switch to top view; wireframe
    view.set_std_view_eye()
    if not nice_design:
        exec_bool("SetShading")



def sort_for_material():
    materials_dict = {}
    sheets = nest.get_sheets()
    for sheet in sheets:
        material = nest.get_sheet_property(sheet, nest.SheetProperties.MATERIAL)
        thickness = nest.get_sheet_property(sheet, nest.SheetProperties.THICKNESS)
        key = (material, thickness) #tuple
        if key in materials_dict:
            materials_dict[key].append(sheet)
        else:
            #materials_dict[(material, thickness)] = []   #tuple = list
            materials_dict[key] = [sheet]
    return materials_dict


def create_report(report_file_path, project_name, folder, img_ext, logo, nice_design, rotate, reports_pdfs_together, divide_material, sheets_to_report, materials_dict, counter_efficiency_total, counter_for_full_pdf):
    try:
        with open(report_file_path, 'w', encoding='utf-8') as html_file: #html_file is an object

            if counter_for_full_pdf == 0: 
                #white header and css, IF divide_material=True (bc then counter_for_full_pdf==0) OR if that's the first key (pdf) for reports_pdfs_together=True

                #HTML header and file name
                html_header_write(html_file, project_name)

                #here add write_fancy_css or write_css_for_printing
                if nice_design:
                    write_nice_css(html_file)
                else:
                    write_css_printing(html_file)

            counter_sheet_in_sheets = 0

            if rotate: #rotate sheets by 90 degrees
                for sheet in sheets_to_report:
                    cad.rotate(sheet, 0, 0, -90, False)

            #if reports_pdfs_together and divide_material:
            #I ALWAYS count stuff separately for material
            if counter_for_full_pdf == 0:
                #i need to have these variables as 0 for each key only if reports are split for material, or if that's the first key (pdf) for reports_pdfs_together=True
                #it could be if divide_material or counter_for_full_pdf == 0: but it's the same thing since in nesting_report if divide_material  I'm always passing 0 for counter_for_full_pdf
                total_area = 0        # m2
                total_garbage = 0     # %
                total_reusable = 0    # %

            for sheet in sheets_to_report:
                sheet_obj = get_sheet_obj(folder, sheet, counter_sheet_in_sheets, img_ext, total_area, total_reusable, total_garbage)

                total_area, total_garbage, total_reusable, counter_sheet_in_sheets = write_html(folder, html_file, logo, project_name, sheets_to_report, reports_pdfs_together, nice_design, divide_material, sheet_obj, counter_efficiency_total)

                if rotate:
                    cad.rotate(sheet, 0, 0, 90, False)

            close_html(html_file)

        try:
            if divide_material:
                output_pdf = os.path.join(folder, f'{project_name}.pdf')
            else:
                output_pdf = os.path.join(folder, 'report.pdf')
                test = f'{project_name}.pdf'
                #!!!!!!!!!!!!!!!!!!






            to_pdf(report_file_path, output_pdf)
        except Exception as e:
            dlg.output_box(f" :C {e}")

    except IOError as e:
        dlg.output_box(f"Ein Fehler ist beim Schreiben der Datei '{report_file_path}' aufgetreten: {e}")



def get_sheet_obj(folder, sheet, counter_sheet_in_sheets, img_ext, total_area, total_reusable, total_garbage):

    img_path = f"{folder}\{sheet}{img_ext}"
    area = nest.get_sheet_property(sheet, nest.SheetProperties.AREA)
    mat_leftover = nest.get_sheet_property(sheet, nest.SheetProperties.RATE_LEFT_OVER)      # % of sheet   garbage not reusable material
    mat_reusable = nest.get_sheet_property(sheet, nest.SheetProperties.RATE_REUSABLE)       # % of sheet   reusable material
    area = (area / 1000000)   # m2
    total_area += area        # m2
    total_reusable += mat_reusable   # %
    total_garbage += mat_leftover    # %

    if not os.path.isfile(img_path):
        os.makedirs(os.path.dirname(img_path), exist_ok=True)

    view.zoom_on_object(sheet, ratio=1)
    nest.get_sheet_preview(sheet, img_path, 0.35) # 0.35, so the lines will be thicker

    #all the html, incl. efficiency
    return ReportSheet(
        sheet=sheet,
        mat_leftover=mat_leftover,
        mat_reusable=mat_reusable,
        area=area,
        counter_sheet_in_sheets=counter_sheet_in_sheets,
        img_path=img_path,
        total_area=total_area,
        total_reusable=total_reusable,
        total_garbage=total_garbage
    )



def html_header_write(html_file_object, project_name):
    """
    Write HTML header

    :param html_file_object: the file object to which the HTML content will be written
    :type html_file_object: file-like object (e.g., obtained via open() in write mode)
    :param project_name: the name of the project
    :type project_name: str
    """

    line = f"""<!DOCTYPE HTML>
<HTML lang="de">
<HEAD>
    <META charset="UTF-8">
    <META name="viewport" content="width=device-width, initial-scale=1.0">
    <TITLE>Nesting report for {project_name}</TITLE>
    """
    html_file_object.write(line)



def write_nice_css(html_file_object):   #here I could do a fancier design, if that would be a good idea
    """
    Write vissually appealing CSS if nice_design is True.

    :param html_file_object: the file object to which the HTML content will be written
    :type html_file_object: file-like object (e.g., obtained via open() in write mode)
    """

    line = """
    <STYLE>
        body {
            font-family: sans-serif;
            margin-left: 15px
        }

        table {
            border: 1px solid rgb(186, 186, 186);
            margin-bottom: 10px;
            /* border-collapse: collapse;   uncomment for a singular line border*/
        }

        .table-container {
            display: table;
        }

        .mainTable {
            display: inline-block;
            width: fit-content;
        }

        .adjustable-table {
            width: 100%;
        }

        .page-break-after {
            page-break-after: always;
        }

        .page-break-before {
            page-break-before: always;
        }

        th, td {
            border: 1px solid #bababa;
            text-align: left;
            font-weight: normal;
            padding: 3px;
            background-color: rgba(186, 186, 186, 0.17);
            border-radius: 2px;
        }

        .td-right {
            text-align: right;
        }

        #thick-border {
            border-width: 3px;
        }

        .center-text, #thick-border th.center-text {
            font-size: 22px;
            padding: 3px;
            text-align: center;
            font-weight: 500;
        }

        #thick-border td, 
        #thick-border th {
            padding: 5px;
            font-size: 18px;
        }

        .green {
            background-color: rgba(105, 191, 74, 0.556);
        }

        .grey {
            background-color: rgba(186, 186, 186, 0.632);
        }
        
        .right-align {
            text-align: right;
            font-size: 20px;
        }

    </STYLE>
    </HEAD>
    <BODY>
"""
    html_file_object.write(line)



def write_css_printing(html_file_object):
    """
    Write minimalisting black and white style CSS if nice_design is not True.

    :param html_file_object: the file object to which the HTML content will be written
    :type html_file_object: file-like object (e.g., obtained via open() in write mode)
    """

    line = """
    <STYLE>
        body {
            font-family: sans-serif;
            margin-left: 15px
        }

        table {
            border: 1px solid rgb(186, 186, 186);
            margin-bottom: 10px;
            border-collapse: collapse;
        }

        .table-container {
            display: table;
        }

        .mainTable {
            display: table;
            width: 100%; /*or width: fit-content;    ?*/
        }

        .adjustable-table {
            width: 100%;
        }

        .page-break-after {
            page-break-after: always;
        }

        .page-break-before {
            page-break-before: always;
        }

        th, td {
            border: 1px solid #bababa;
            text-align: left;
            font-weight: normal;
            padding: 3px;
            border-radius: 2px;
        }

        .td-right {
            text-align: right;
        }

        .center-text, #thick-border th.center-text {
            font-size: 22px;
            padding: 3px;
            text-align: center;
            font-weight: 500;
        }

        #thick-border td, 
        #thick-border th {
            padding: 5px;
            font-size: 18px;
        }

        .right-align {
            text-align: right;
            font-size: 20px;
        }

    </STYLE>
    </HEAD>
    <BODY>
"""
    html_file_object.write(line)



























def efficiency_for_sheet(html_file, pieces, area, curr1, curr2):
    """
    Efficiency for each sheet
    
    :param pieces: the number of pieces
    :type pieces: int
    :param area: the area of the sheet
    :type area: float
    :param curr1: the percentage of not reusable material
    :type curr1: float
    :param curr2: the percentage of reusable material
    :type curr2: float
    :return: string that will be written into HTML
    :rtype: str
    """

    html_file.write(f"""<BR/>
        <BR/>
        <TABLE class="adjustable-table">
        <TH colspan="3" class="center-text">Effizienzbericht</TH>
        <TR>
            <TD>Gutteile</TD>
            <TH colspan="2" align="left">{int(pieces)}</TH>
        </TR>
        <TR>
            <TD>Fläche der Platte</TD>
            <TH colspan="2" align="left">{round(area, 2)} m²</TH>
        </TR>
        <TR>
            <TD class="green">Wiederverwendbares Material</TD>
            <TD class="green">{round(curr2, 2)}% der Platte</TD>
            <TD class="green">{round(curr2 * area /100, 2)} m²</TD>
        </TR>
        <TR>
            <TD class="grey">Nicht wiederverwendbares Material</TD>
            <TD class="grey">{round(curr1, 2)}% der Platte</TD>
            <TD class="grey">{round(curr1 * area /100, 2) } m²</TD>
        </TR>
    </TABLE>
        """)


def efficiency_sheets_total(html_file_object, number_of_sheets, total_area, total_reusable, total_garbage):
    """
    Total report for all sheets, if the sheet count exceeds 1
    
    :param html_file_object: the file object to which the HTML content will be written
    :type html_file_object: file-like object (e.g., obtained via open() in write mode)
    :param number_of_sheets: the total number of sheets
    :type number_of_sheets: int
    :param total_area: the total area of all sheets
    :type total_area: float
    :param total_reusable: the total percentage of reusable material
    :type total_reusable: float
    :param total_garbage: the total percentage of not reusable material
    :type total_garbage: float
    """

    average_reusable = total_reusable / number_of_sheets #  %  -  I need to not only add % from every sheet, but also get the Durchschnitt per sheet
    average_garbage = total_garbage / number_of_sheets   #  %

    html_file_object .write(f"""
        <TABLE id="thick-border" class="adjustable-table">
            <TH colspan="3" class="center-text">Gesamtwirkungsgradbericht</TH>

            <TR>
                <TH align="left">Anzahl Sheets</TH>
                <TH colspan="2" align="left">{number_of_sheets}</TH>
            </TR>

            <TR>
                <TD>Gesamtfläche</TD>
                <TH colspan="2" align="left">{round(total_area, 2)} m²</TH>
            </TR>


            <TR>
                <TD class="green">Gesamt wiederverwendbares Material</TD>
                <TH colspan="2" align="left" class="green">{round(total_area * total_reusable / 100 / number_of_sheets, 2)} m²</TH>
            </TR>            
            <TR>
                <TD class="grey">Gesamt nicht wiederverwendbares Material</TD>
                <TH colspan="2" align="left" class="grey">{round(total_area * total_garbage / 100 / number_of_sheets, 2)} m²</TH>
            </TR>


            <TR>
                <TD class="green">Wiederverwendbares Material pro Platte im Durchschnitt:</TD>
                <TD class="green">{round(total_reusable / number_of_sheets, 2)}%</TD>
                <TD class="green">{round(total_area * average_reusable / 100 / number_of_sheets, 2)} m²</TD>
            </TR>

            <TR>
                <TD class="grey">Nicht wiederverwendbares Material pro Platte im Durchschnitt</TD>
                <TD class="grey">{round(total_garbage / number_of_sheets, 2)}%</TD>
                <TD class="grey">{round(total_area * average_garbage / 100 / number_of_sheets, 2)} m²</TD>
            </TR>
        </TABLE>
        """)



def close_html(html_file_object):
    """
    Closing HTML document

    :param html_file_object: the file object to which the HTML content will be written
    :type html_file_object: file-like object (e.g., obtained via open() in write mode)
    """
    line = '</BODY></HTML>'
    html_file_object.write(line)



def to_pdf(report_html_path, output_pdf_path):
    """
    Convert HTML to PDF

    :param report_html_path: the file path where the generated HTML report will be saved
    :type report_html_path: str
    :param output_pdf_path: the file path where the generated PDF report will be saved
    :type output_pdf_path: str
    """
    do_debug()
    path_to_wkhtmltopdf = r'C:\Program Files\...\wkhtmltopdf.exe'
    subprocess.call([path_to_wkhtmltopdf, report_html_path, output_pdf_path])




if __name__ == '__main__':
    nesting_report()
