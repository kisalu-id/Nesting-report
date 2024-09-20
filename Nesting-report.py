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
    config_nr.run_config()
    try:
        ini_path = ewd.explode_file_path(r'%MACHPATH%\script\config_nr.ini')

        config = configparser.ConfigParser()
        config.read(ini_path)

        rotate = config.get('Einstellung', 'rotate') #if True, rotate
        rotate = False if rotate == "0" or rotate =="False" else True

        #if checked, delete the previous folder with the same name
        delete_folder =  config.get('Einstellung', 'delete_folder')
        delete_folder = False if delete_folder == "0" or delete_folder =="False" else True

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


    except FileNotFoundError:
        dlg.output_box('Fehler: Die Konfigurationsdatei "config_nr.ini" wurde nicht gefunden. Bitte überprüfen Sie den Dateipfad.')
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

    project_name = ewd.get_project_name() #get the name of the opened ewd project
    if not project_name.endswith (".ewd"):
        project_name = datetime.datetime.now().strftime("%Y_%m_%d_%H-%M")
        ewd.save_project(ewd.explode_file_path(f"%TEMPPATH%//{project_name}.ewd"))
        project_name = f"{project_name}.ewd"

    #subfolder with the project name: it will be deleted, if it already exists, then the new folder will be created
    #(so the date of creation of this folder on user's pc will be fresh -> easy to sort)
    folder = os.path.join(general_folder, f'{os.path.splitext(project_name)[0]}')

    #if folder exists and user set this setting, delete
    if delete_folder and os.path.exists(folder):
        #maybe later add "ok" and "cancel"
        remove_existing_folder_with_same_name(folder)

    else: #if folder doesn't exist, create
        try:
            os.makedirs(folder, exist_ok=False)
        except OSError as e:
            dlg.output_box(f"Fehler beim Ordner erstellen in {folder}")


    report_file_path = f'{folder}\\report.html'
    try:
        os.makedirs(os.path.dirname(report_file_path), exist_ok=True)
    except OSError as e:
        dlg.output_box(f"Fehler beim Ordner erstellen in {os.path.dirname(report_file_path)}")

    img_ext = ".jpg"

    logo = "C:\Program Files\...\logo.png"

    #switch to top view; wireframe
    view.set_std_view_eye()
    # exec_bool("SetWireFrame")
    exec_bool("SetShading")
    ### to execute this with just a wrapper around the scl, we can do sclcore.execute_command_boolean("SetWireFrame")

    materials_dict = {}


    #try:

    #what i do:
    #sort sheets depending on material and thickness
    #for each materials_dict[key] generate html file, with one or mpre shets; if len(sheets) > 1, write (optionally, separate) general report
    if divide_material:
        sheets = nest.get_sheets()
        for sheet in sheets:
            material = nest.get_sheet_property(sheet, nest.SheetProperties.MATERIAL)
            thickness = nest.get_sheet_property(sheet, nest.SheetProperties.THICKNESS)
            key = (material, thickness) #tuple
            sort_for_material(sheet, materials_dict, key) #basically sorting and saving as a dictionary
        for key in materials_dict:
            materials_dict[key].sort()
        #ok. i sorted


        #now do the fork of "if divide_material or not"

        #ig delete the report_file_path = f'{folder}\\report.html'

        for material_and_thickness, sheets_values in materials_dict.items():
            material, thickness = material_and_thickness  #extract material and thickness from the key

            #html with a new name
            report_file_path = f"{folder}\\{material}_{thickness}.html"
            try:
                os.makedirs(os.path.dirname(report_file_path), exist_ok=True)
            except OSError as e:
                dlg.output_box(f"Fehler beim Ordner erstellen in {os.path.dirname(report_file_path)}")
            #do the thing  !
            project_name = f"{material}_{thickness}_{project_name}"
            the_thing(report_file_path, project_name, nice_design, rotate, folder, img_ext, logo, reports_pdfs_together, divide_material, sheets_values)

    else: #if not divide_material:
        #do the thing  !
        sheets = nest.get_sheets()
        the_thing(report_file_path, project_name, nice_design, rotate, folder, img_ext, logo, reports_pdfs_together, divide_material, sheets)


def the_thing(report_file_path, project_name, nice_design, rotate, folder, img_ext, logo, reports_pdfs_together, divide_material, sheets):
    try:
        with open(report_file_path, 'w', encoding='utf-8') as html_file: #html_file is an object

            #HTML header and file name
            html_header_write(html_file, project_name)

            #here add write_fancy_css or write_css_for_printing
            if nice_design:
                write_nice_css(html_file)
            else:
                write_css_printing(html_file)

            count = 0

            if rotate: #rotate sheets by 90 degrees
                for sheet in sheets:
                    cad.rotate(sheet, 0, 0, -90, False)

            if not divide_material:
                sheets = nest.get_sheets()
            total_area = 0        # m2
            total_garbage = 0     # %
            total_reusable = 0    # %
            for sheet in sheets:
                #name of the jpg
                img_path = f"{folder}\{sheet}{img_ext}"
                area = nest.get_sheet_property(sheet, nest.SheetProperties.AREA)                 #_NSheetArea
                curr1 = nest.get_sheet_property(sheet, nest.SheetProperties.RATE_LEFT_OVER)      #_NSheetRateLeftOver  % of sheet   garbage not reusable material
                curr2 = nest.get_sheet_property(sheet, nest.SheetProperties.RATE_REUSABLE)       #_NSheetRateReusable  % of sheet   reusable material
                area = (area / 1000000)   # m2
                total_area += area        # m2
                total_garbage += curr1    # %
                total_reusable += curr2   # %

                object_path = sheet

                if not os.path.isfile(img_path):
                    os.makedirs(os.path.dirname(img_path), exist_ok=True)

                view.zoom_on_object(object_path, ratio=1)
                nest.get_sheet_preview(sheet, img_path, 0.35) # 0.35, so the lines will be thicker

                #all the html, incl. efficiency
                write_html(html_file, logo, project_name, sheet, sheets, count, total_area, curr1, curr2, total_reusable, total_garbage, img_path, area, reports_pdfs_together, folder, nice_design, divide_material)
                count += 1
                
                if rotate:
                    cad.rotate(sheet, 0, 0, 90, False)

            close_html(html_file)

        try:
            if divide_material:
                output_pdf = os.path.join(folder, f'{project_name}.pdf')
            else:
                output_pdf = os.path.join(folder, 'report.pdf')
            to_pdf(report_file_path, output_pdf)
        except Exception as e:
            dlg.output_box(f" :C {e}")

    except IOError as e:
        dlg.output_box(f"Ein Fehler ist beim Schreiben der Datei '{report_file_path}' aufgetreten: {e}")



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


def html_header_write(html_file, project_name):
    """
    Write HTML header
    :param project_name: the name of the project
    :type project_name: str
    :return: string that will be written into HTML
    :rtype: str
    """

    html_file.write(f"""<!DOCTYPE HTML>
<HTML lang="de">
<HEAD>
    <META charset="UTF-8">
    <META name="viewport" content="width=device-width, initial-scale=1.0">
    <TITLE>Nesting report for {project_name}</TITLE>
    """)


def write_css(file):
    """
    :param file: the name of the HTML file to write the CSS into
    :type file: str
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

        #mainTable {
            display: inline-block;
            width: fit-content;
        }

        .adjustable-table {
            width: 100%;
        }

        th, td {
            border: 1px solid black;
            text-align: left;
            font-weight: normal;
            padding: 3px;
            background-color: rgba(186, 186, 186, 0.17);
            border-radius: 2px;
            border-color: #bababa;
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
    file.write(line)


def write_html(file, logo, project_name, sheet, sheets, count, total_area, curr1, curr2, total_reusable, total_garbage, img_path, area, reports_pdfs_together, folder):
    """
    Get the sheet properties, write html file and piece properties.

    :param file: the name of the HTML file to write into
    :type file: str
    :param logo: the URL of the logo image file
    :type logo: str
    :param project_name: the name of the project
    :type project_name: str
    :param sheet: the sheet name
    :type sheet: str
    :param sheets: a list of sheets
    :type sheets: list
    :param count: the number of sheets... I think.... it's quite misterious actually, how that was used in the SCL code
    :type count: int
    :param total_area: the total area of all sheets, to use for total_efficiency for the last sheet
    :type total_area: float
    :param curr1: the percentage of not reusable material
    :type curr1: float
    :param curr2: the percentage of reusable material
    :type curr2: float
    :param total_reusable: the total percentage of reusable material
    :type total_reusable
    :param area: the area of the sheet
    :type area: float
    """

    pieces = nest.get_sheet_property(sheet, nest.SheetProperties.PIECES_NUMBER)      #_NSheetNumPieces
    thickness = nest.get_sheet_property(sheet, nest.Properties.SHEET_THICKNESS)      #_NSheetRateReusable
    material = nest.get_sheet_property(sheet, nest.SheetProperties.MATERIAL)         #_NSheetMaterial
    width = nest.get_sheet_property(sheet, nest.SheetProperties.WIDTH)               #_NSheetWidth
    height = nest.get_sheet_property(sheet, nest.SheetProperties.HEIGHT)             #_NSheetHeight
    current_date = datetime.datetime.now().strftime("%d.%m.%Y")

    if count == 0:
        line = '<DIV style="display: inline-block; width: 100%; text-align: left;"> '
        line += f'<IMG src="file:///{logo}" alt="Logo" style="vertical-align: middle; width: 70px; height: 70px;"> '
        line += f'<SPAN style="font-size: 35px; margin-left: 20px; vertical-align: middle;">Projekt: {os.path.splitext(project_name)[0]} </SPAN>'
        line += '</DIV> '
        file.write(line)

    #table for sheet information
    line = '<DIV class="table-container"> <TABLE '
    if count != 0:
        line += 'style="page-break-before:always" '
    line += 'id="mainTable"> '
    file.write(line)

    #row with sheet name
    line = f'<TR> <TD style="font-size:30px" colspan="6">{sheet}</TD>'

    s_count = str(count).zfill(3)

    #adding the date
    line += f'<TD colspan="4" class="right-align">{current_date}</TD> </TR> '
    file.write(line)

    #sheet information - width, height, thickness, name, material
    line = f"""
    <TR>
        <TD align="middle">Breite</TD>
        <TD align="middle">{round(width, 2)}</TD>
        <TD align="middle">Höhe</TD>
        <TD align="middle">{round(height, 2)}</TD>
        <TD align="middle">Stärke</TD>
        <TD align="middle">{round(thickness, 2)}</TD>
        <TD align="middle">Material</TD>
        <TD align="middle">{material}</TD>
    </TR>
    """
    file.write(line)

    #picture from the sheet
    size_img = "width=\"1200pt\"" if width > 3 * height else "height=\"400pt\""
    line = f'<TR> <TD colspan="10"> <IMG src="file:///{img_path}"{size_img}> </TD></TR>'
    file.write(line)

    #write down the individual information about the components.
    pieces_on_sheet = nest.get_pieces(sheet)
    n_piece_count = 1
    for piece in pieces_on_sheet:
        piece_width = nest.get_piece_property(piece, nest.PieceProperties.WIDTH)
        piece_height = nest.get_piece_property(piece, nest.PieceProperties.HEIGHT)
        piece_label = nest.get_piece_property(piece, nest.PieceProperties.LABEL)

        line = f"""
        <TR class="adjustable-table" style="width: 100%;">
            <TD align="middle">Nr.</TD>
            <TD align="middle">{n_piece_count}</TD>
            <TD align="middle">Bezeichnung</TD>
            <TD align="middle">{piece_label}</TD>
            <TD align="middle">Breite</TD>
            <TD align="middle">{round(piece_width, 2)}</TD>
            <TD align="middle">Höhe</TD>
            <TD align="middle">{round(piece_height, 2)}</TD>
        </TR>

        """
        file.write(line)
        n_piece_count += 1

    line = '</TABLE>' #closing mainTable
    file.write(line)

    count_efficiency(file, sheet, sheets, pieces, total_area, curr1, curr2, total_reusable, total_garbage, area, reports_pdfs_together, folder, logo, project_name)

    line = '</DIV>' #closing <DIV class="table-container">
    file.write(line)




def count_efficiency(file, sheet, sheets, pieces, total_area, curr1, curr2, total_reusable, total_garbage, area, reports_pdfs_together, folder, logo, project_name):
    """
    Count efficiency for each sheet, if there's more than 1 sheet - do the total efficiency

    :param file: the name of the HTML file to write into
    :type file: str
    :param sheet: the sheet name
    :type sheet: str
    :param sheets: a list of sheets
    :type sheets: list
    :param pieces: the number of pieces
    :type pieces: int
    :param total_area: the total area of all sheets m²
    :type total_area: float
    :param curr1: the percentage of not reusable material %
    :type curr1: float
    :param curr2: the percentage of reusable material %
    :type curr2: float
    :param total_reusable: the total percentage of reusable material %
    :type total_reusable: float
    :param total_garbage: the total percentage of not reusable material %
    :type total_garbage: float
    :param area: the area of the sheet
    :type area: float
    """

    #    autocam_aktivieren = config.get('SETTINGS', 'autocam_aktivieren')
    #    if autocam_aktivieren:
    sheets = nest.get_sheets() #do i need that?.. or then delete an argument to this function

    if sheet:
        efficiency_for_sheet(file, pieces, area, curr1, curr2)

        #here we do total_efficiency after every sheet was dealt with
        #if current sheet is the last sheet in list of sheets AND it's not a single sheet in a list AND if "reports_pdfs_together" is True
        if sheet == sheets[-1] and sheets[0] != sheet and reports_pdfs_together == True:
            number_of_sheets = len(sheets)
            efficiency_sheets_total(file, number_of_sheets, total_area, total_reusable, total_garbage)

        if sheet == sheets[-1] and sheets[0] != sheet and reports_pdfs_together != True:
            
            total_report_path = os.path.join(folder, 'total_report.html')
            try:
                with open(total_report_path, 'w', encoding='utf-8') as total_report_html:
                    #HTML header and file name
                    html_header_write(total_report_html, project_name)
                    #here add write_fancy_css or write_css_for_printing
                    write_css(total_report_html)

                    number_of_sheets = len(sheets)
                    efficiency_sheets_total(total_report_html, number_of_sheets, total_area, total_reusable, total_garbage)
                    close_html(total_report_html)
                    output_pdf = os.path.join(folder, 'total_report.pdf')
                to_pdf(total_report_path, output_pdf)

            except IOError as e:
                dlg.output_box(f"Ein Fehler ist beim Schreiben der Datei '{total_report_path}' aufgetreten: {e}")


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
