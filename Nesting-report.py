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


def nesting_report():
    """
    Get the "turn" setting from ini, get or create the folder, delete previous files there, create and open HTML, delete empty sheets, write info about each sheet,
    count efficiency, write css and html
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
        # - regardless of the user choice, there will be  created a new folde, that will only have report files, that are safe to delete

        nice_design = config.get('Druckeinstellungen', 'nice_design') #if True, nice_design
        nice_design = False if nice_design == "0" or nice_design =="False" else True

        reports_pdfs_together = config.get('Druckeinstellungen', 'reports_pdfs_together') #if True, reports_pdfs_together
        reports_pdfs_together = False if reports_pdfs_together == "0" or reports_pdfs_together =="False" else True

        only_measures_BW = config.get('Druckeinstellungen', 'only_measures_BW') #if True, only_measures_BW
        only_measures_BW = False if only_measures_BW == "0" or only_measures_BW =="False" else True


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


    ###if project is not saved --> save it in temp folder in program or in a temp dircetory in the config file
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
        #add "ok" and "cancel"
        remove_existing_folder_with_same_name(folder)

    else: #if folder doesn't exist, create
        try:
            os.makedirs(folder, exist_ok=False)
        except OSError as e:
            dlg.output_box(f"Fehler beim Ordner erstellen in {folder}")


    report_file = f'{folder}\\report.html'
    try:
        os.makedirs(os.path.dirname(report_file), exist_ok=True)
    except OSError as e:
        dlg.output_box(f"Fehler beim Ordner erstellen in {os.path.dirname(report_file)}")

        ### better code would proably to use a .join function later instead of adding the //
    if not folder.endswith("\\"):
        os.path.join(folder, '')
    img_ext = ".jpg"

    logo = "C:/ProgramData/.../logo_black.png"

    #switch to top view; wireframe
    view.set_std_view_eye()
    exec_bool("SetWireFrame")
    # exec_bool("SetShading")

    try:
        with open(report_file, 'w', encoding='utf-8') as html_file:
            #HTML header and file name
            html_header_write(html_file, project_name)

            write_css(html_file)
            sheets = nest.get_sheets()
            count = 0
            for sheet in sheets:
                if rotate: #rotate sheets by 90 degrees
                    cad.rotate(sheet, 0, 0, -90, False)

            sheets = nest.get_sheets()
            total_area = 0        # m2
            total_garbage = 0     # %
            total_reusable = 0    # %
            for sheet in sheets:
                #name of the jpg
                img_path = f"{folder}//{sheet}{img_ext}"
                area = nest.get_sheet_property(sheet, nest.SheetProperties.AREA)                 #_NSheetArea
                curr1 = nest.get_sheet_property(sheet, nest.SheetProperties.RATE_LEFT_OVER)      #_NSheetRateLeftOver  % of sheet   garbage not reusable material
                curr2 = nest.get_sheet_property(sheet, nest.SheetProperties.RATE_REUSABLE)       #_NSheetRateReusable  % of sheet   reusable material
                area = (area / 1000000)   # m2
                total_area += area        # m2
                total_garbage += curr1    # %
                total_reusable += curr2   # %

                object_path = groups.get_current()

                #get_project_path
                if not os.path.isfile(img_path):
                    os.makedirs(os.path.dirname(img_path), exist_ok=True) 

                view.zoom_on_object(object_path, ratio=1)
                nest.get_sheet_preview(sheet, img_path, 0.35) # 0.35, so lines will be thicker

                #all the html, incl. efficiency
                write_html(html_file, logo, project_name, sheet, sheets, count, total_area, curr1, curr2, total_reusable, total_garbage, img_path, area)
                count += 1
                
                if rotate:
                    cad.rotate(sheet, 0, 0, 90, False)

            insert_java_script(html_file, count)
            #close HTML
            line = '</BODY></HTML>'
            html_file.write(line)


        try:
            to_pdf(report_file, folder)
        except Exception as e:
            dlg.output_box(f" :C {e}")

    except IOError as e:
        dlg.output_box(f"Ein Fehler ist beim Schreiben der Datei '{report_file}' aufgetreten: {e}")



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
    <SCRIPT type="text/javascript" src="js/code39.js"></SCRIPT>
    """)


def write_css(file):   #here I could do a fancier design, if that would be a good idea
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
            display: inline-block;
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


def write_html(file, logo, project_name, sheet, sheets, count, total_area, curr1, curr2, total_reusable, total_garbage, img_path, area):
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
        line = '<DIV style="display: flex; align-items: center;">'
        line += f'<IMG src="file:///{logo}" alt="DDX Logo" width="70" height="70">'
        line += f'<SPAN style="font-size: 35px; margin-left: 20px; align-self: auto;">Projekt: {os.path.splitext(project_name)[0]}</SPAN>'
        line += '</DIV>'
    else:
        line = '<DIV style="page-break-before:always;"></DIV>'

    file.write(line)

   #table for sheet information
    line = '<DIV class="table-container"> <TABLE id="mainTable">'
    file.write(line)

    #row with sheet name
    line = f'<TR><TD style="font-size:30px" colspan="6">{sheet}</TD>'

    s_count = str(count).zfill(3)

    #adding the date
    line += f'<TD colspan="4" class="right-align">{current_date}</TD></TR>'
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

    count_efficiency(file, sheet, sheets, pieces, total_area, curr1, curr2, total_reusable, total_garbage, area)

    #close table
    line = '</BR>'
    file.write(line)




def count_efficiency(file, sheet, sheets, pieces, total_area, curr1, curr2, total_reusable, total_garbage, area):
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
        if sheet == sheets[-1] and sheets[0] != sheet: #if current sheet is the last sheet in list of sheets; AND it's not a single sheet in a list
            number_of_sheets = len(sheets)
            efficiency_sheets_total(file, number_of_sheets, total_area, total_reusable, total_garbage)


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


def efficiency_sheets_total(html_file, number_of_sheets, total_area, total_reusable, total_garbage):
    """
    :param number_of_sheets: the total number of sheets
    :type number_of_sheets: int
    :param total_area: the total area of all sheets
    :type total_area: float
    :param total_reusable: the total percentage of reusable material
    :type total_reusable: float
    :param total_garbage: the total percentage of not reusable material
    :type total_garbage: float
    :return: string that will be written into HTML
    :rtype: str
    """

    average_reusable = total_reusable / number_of_sheets #  %  -  I need to not only add % from every sheet but also get the Durchschnitt per sheet
    average_garbage = total_garbage / number_of_sheets   #  %

    html_file.write(f"""<BR/>
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
        <BR/>
        """)


def insert_java_script(file, count):
    line = """
    <SCRIPT type="text/javascript">
    /* <![CDATA[ */
    function get_object(id) {
        var object = null;
        if (document.layers) {
            object = document.layers[id];
        } else if (document.all) {
            object = document.all[id];
        } else if (document.getElementById) {
            object = document.getElementById(id);
        }
        return object;
    }
    
    window.onload = function() {
        var mainTable = document.getElementById('mainTable');
        if (mainTable) {
            var mainTableWidth = mainTable.offsetWidth;
            var adjustableTables = document.querySelectorAll('.adjustable-table');
            adjustableTables.forEach(function(table) {
                table.style.width = mainTableWidth + 'px';
            });
        }
    };

    /* ]]> */
    </SCRIPT>
    """
    file.write(line)






def to_pdf(report_file, folder):
    """
    Convert HTML to PDF

    :param report_file: the name of the HTML report file to be converted to PDF
    :type report_file: str
    :param folder: the directory where the PDF will be saved
    :type folder: str
    """
    do_debug()
    subprocess.call([path_to_wkhtmltopdf])
    #sclcore.execute_file(report_file, folder)
    path_to_wkhtmltopdf



if __name__ == '__main__':
    nesting_report()
