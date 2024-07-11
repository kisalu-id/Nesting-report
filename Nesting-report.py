import os
from datetime import datetime
import configparser
import ewd
#import pdfkit
from ewd import groups
from company import dlg
from company import nest
from company.gdb import cad
from company.gdb import view
from sclcore import do_debug



def main():
    do_debug()
    ini_path = ewd.explode_file_path(r'C:\Users\...\Nesting_report\config_nr.ini') #wo liegt die ini Datei von Nest settings
    config = configparser.ConfigParser()
    config.read(ini_path)
    rotate = config.get('Einstellung', 'turn')
    rotate = False if rotate == "0" else True

    #?? save the project automatically here? line 129
    # if project is only saved, then the scl is aborted
    #   if( $NLS) return ;
    #   dodebug() ;
    # if file is not saved, then abort here...
    # if( STRSTR( GetName(), ".ewd") < 1) {
    project_name = ewd.get_project_name() # get the name of the opened ewd project
    if not project_name.endswith (".ewd"):
        dlg.output_box( "Projekt bitte speichern. Die Schnittpläne und Labels wurden nicht erstellt.")
    
    #path for the print files
    # = ExplodeFilePath( "%SETTINGPATH%" + "\Nest\Print") ;
    folder =  config.get('Pfad', 'report_pfad')

    #if folder doesn't exist, create
    if not os.path.exists(folder):
        os.makedirs(folder, exist_ok=False)
    
    #iterate through each file and remove it
    file_in_folder = next(os.walk(folder))[2][0] if os.listdir(folder) else ""
    while (file_in_folder != ""):
            os.remove(os.path.join(folder, file_in_folder))
            file_in_folder = next(os.walk(folder))[2][0] if os.listdir(folder) else ""

    report_file = f'{folder}\\report.html' #--- again, need to do something better than this
    if not os.path.isfile(report_file):
        os.makedirs(os.path.dirname(report_file), exist_ok=True)

    n_sheets = 0
    if not folder.endswith("\\"):
        folder += "\\"
    img_ext = ".jpg"

    logo = (f"file:///{report_file}\\logo1.png") #URL

    #switch to top view; wireframe
    view.set_std_view_eye()
    #Ok = bOk  AND  SetWireFrame() ???? how to make it

    with open(report_file, 'w', encoding='utf-8') as html_file:
            #open status bar            Ok = bOk  AND  ProgressCreate( "Print") ;

            #HTML header and file name
            html_file.write(html_header_write(project_name))
            nest.activate()
            sheets = nest.get_sheets()

            for sheet in sheets: #call sheets list, delete number of sheets and empty sheets
                if sheet == "": #how can it see that its empty or not?
                    sheet_to_delete = list(sheet) #NestDeleteSheets( sSheet)
                    #mb change modul??
                    nest.delete_sheets(sheet_to_delete) # ??? I'm very unsure what im doing; I didn't find how to delete a single sheet
                else:
                    n_sheets += 1

                if rotate: #rotate sheets by 90 degrees
                    cad.rotate(sheet, 0, 0, -90, False)

            sheets = nest.get_sheets()
            count = 0

            for sheet in sheets:
                #name of the jpg
                img_path = f"{folder}{sheet}{img_ext}"

                nC = 1
                total_area = 0
                total_garbage = 0
                total_reusable = 0
                nSheets = 0
                width = nest.get_sheet_property(sheet, nest.SheetProperties.WIDTH)               #_NSheetWidth
                height = nest.get_sheet_property(sheet, nest.SheetProperties.HEIGHT)             #_NSheetHeight
                thickness = nest.get_sheet_property(sheet, nest.Properties.SHEET_THICKNESS)      #_NSheetRateReusable
                material = nest.get_sheet_property(sheet, nest.SheetProperties.MATERIAL)         #_NSheetMaterial
                area = nest.get_sheet_property(sheet, nest.SheetProperties.AREA)                 #_NSheetArea
                pieces = nest.get_sheet_property(sheet, nest.SheetProperties.PIECES_NUMBER)      #_NSheetNumPieces
                curr1 = nest.get_sheet_property(sheet, nest.SheetProperties.RATE_LEFT_OVER)      #_NSheetRateLeftOver  % of sheet   garbage not reusable material
                curr2 = nest.get_sheet_property(sheet, nest.SheetProperties.RATE_REUSABLE)       #_NSheetRateReusable  % of sheet   reusable material
                area = (area / 1000000)   # m2
                total_area += area        # m2
                total_garbage += curr1    # %
                total_reusable += curr2   # %
                name = sheet

                object_path = groups.get_current()

                #get_project_path
                if not os.path.isfile(img_path):
                    os.makedirs(os.path.dirname(img_path), exist_ok=True)  #do i for sure need that?

                view.zoom_on_object(object_path, ratio=1)
                nest.get_sheet_preview(sheet, img_path, 1.)
                # bOk = bOk  AND  ExportImage( sImgPath, nHeightImg, nWidthImg, 24) ;

                write_css(html_file)

                #html efficiency for each sheet
                write_html(html_file, logo, project_name, sheet, sheets, count, thickness, name, material, pieces, total_area, curr1, curr2, total_reusable, total_garbage, img_path)
    
                nSheets += 1
                nC += 1

                if rotate:
                    cad.rotate(sheet, 0, 0, 90, False)
                
                # ProgressClose() ;    close progressbar
                # SetShading() ;      Set the view as shaded

                #to_pdf(report_file, folder)



#def change_label_into_id(sheet):
    #??? I don't know; I'm just trying to do the same thing as in SCL
    #label = gdb.get_label(sheet)
    # pieces = nest.get_pieces(sheet)
    # label_pieces = []
    # for piece in pieces:
    #     if piece.upper().endswith("_LABEL"):
    #         label_pieces.append(piece)
    
    # for label_piece in label_pieces:
    #     geo = 



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
    """
    file.write(line)

    while(count > 0):
        s_count = str(count).zfill(3)
        line = f'get_object("code{s_count}").innerHTML = DrawHTMLBarcode_Code39(get_object("code{s_count}").innerHTML, 0, "yes", "in", 0, 3, 0.4, 3, "top", "center", "", "black", "white");\n'
        file.write(line)
        count -= 1

    line = """
    /* ]]> */
    </SCRIPT>
    """
    file.write(line)




def write_html(file, logo, project_name, sheet, sheets, count, thickness, name, material, pieces, total_area, curr1, curr2, total_reusable, total_garbage, img_path):  ####maaaybe do that differently?
    width = nest.get_sheet_property(sheet, nest.SheetProperties.WIDTH)               #_NSheetWidth
    height = nest.get_sheet_property(sheet, nest.SheetProperties.HEIGHT)             #_NSheetHeight
    #bOk = bOk  AND  NestGetPieceProp( sPiece, _NPieceLabel, &sLabel) ;
    line = """
    </HEAD>
    <BODY>
    """


    if count != 0:
        line += ' style="page-break-before:always"'
    file.write(line)

    #table for customer information - header row
    line = f"""
    <DIV>
        <IMG src="{logo}">
        <SPAN style="font-size: 35px; margin-left: 10px;">Projekt: {os.path.splitext(project_name)[0]}</SPAN>
    </DIV>
    """
    file.write(line)

    #table for sheet information
    line = '<TABLE id="mainTable">'
    file.write(line)

    #row with sheet name
    line = f'<TR> <TD style="font-size:30px" colspan="6">{sheet}</TD>'

    count += 1
    s_count = str(count).zfill(3)

    #adding the barcode
    line += f'<TD id="code{s_count}" colspan="4" class="barcode">{sheet.upper()}</TD>'
    line += '</TR>'
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

    count_efficiency(file, sheet, sheets, pieces, total_area, curr1, curr2, total_reusable, total_garbage)

    #close table
    line = '</TABLE></BR>'
    file.write(line)

    #update status
    #perc = count / n_sheets
    #bOk = bOk  AND  ProgressUpdate( perc) ;   ????????
    count += 1

    insert_java_script(file, count)

    #close HTML
    line = '</BODY></HTML>'
    file.write(line)
    #   bOk = bOk  AND  CloseFile( nFile) ;


def count_efficiency(file, sheet, sheets, pieces, total_area, curr1, curr2, total_reusable, total_garbage):
    do_debug()
    #    autocam_aktivieren = config.get('SETTINGS', 'autocam_aktivieren')
    #    if autocam_aktivieren:
    sheets = nest.get_sheets() #do i need that?.. or then delete an arfgument to this function

    if sheet:
        file.write(efficiency_for_sheet(pieces, total_area, curr1, curr2))

        #here we do total_efficiency after every sheet was dealt with
        if sheet == sheets[-1] and sheets[0] != sheet: #if current sheet is the last sheet in list of sheets; AND it's not a single sheet in a list
            number_of_sheets = len(sheets)
            file.write(efficiency_sheets_total(number_of_sheets, total_area, total_reusable, total_garbage))



def efficiency_for_sheet(pieces, total_area, curr1, curr2):
    return f"""<BR/>
        <BR/>
<TABLE class="adjustable-table">
        <TH colspan="3" id="center-text">Effizienzbericht</TH>
        <TR>
            <TD>Guteile</TD>
            <TD>{int(pieces)}</TD>
        </TR>
        <TR>
            <TD>Gesamtfläche</TD>
            <TD>{round(total_area, 2)}  m²</TD>
        </TR>
        <TR>
            <TD>Nicht wiederverwendbares Material</TD>
            <TD>{round(curr1, 2)}% des Blattes   </TD>
        </TR>
        <TR>
            <TD>Wiederverwendbares Material</TD>
            <TD>{round(curr2, 2)}% des Blattes   </TD>
        </TR>
    </TABLE>
        """


def efficiency_sheets_total(number_of_sheets, total_area, total_reusable, total_garbage):
    total_garbage /= number_of_sheets #bc I need to not only add % from every sheet but also get the Durchschnitt
    total_reusable /= number_of_sheets
    return f"""<BR/>
        <TABLE id="thick-border" class="adjustable-table">
            <TH colspan="3" id="center-text">Gesamtwirkungsgradbericht</TH>

            <TR>
                <TH align="left">Anzahl Sheets</TH>
                <TH colspan="2" align="left">{number_of_sheets}</TH>
            </TR>


            <TR>
                <TD>Gesamtfläche</TD>
                <TH colspan="2" align="left">{round(total_area, 2)} m²</TH>
            </TR>

            <TR>
                <TD>Wiederverwendbares Material pro Blatt im Durchschnitt:</TD>
                <TD>{round(total_reusable, 2)}%</TD>
                <TD>{round(total_area * total_reusable / 100, 2)} m²</TD>
            </TR>

            <TR>
                <TD>Nicht wiederverwendbares Material pro Blatt im Durchschnitt</TD>
                <TD>{round(total_garbage, 2)}%</TD>
                <TD>{round(total_area * total_garbage / 100, 2)} m²</TD>
            </TR>
        </TABLE>
        <BR/>
        """


def html_header_write(project_name):
    return f"""<!DOCTYPE HTML>
<HTML lang="de">
<HEAD>
    <META charset="UTF-8">
    <META name="viewport" content="width=device-width, initial-scale=1.0">
    <TITLE>{project_name}</TITLE>
    <SCRIPT type="text/javascript" src="js/code39.js"></SCRIPT>
    """


# def to_pdf(report_file, folder):
#     path_to_wkhtmltopdf = r'C:\Program Files\...\Bundles\wkhtmltopdf.exe'
#     config = pdfkit.configuration(wkhtmltopdf = path_to_wkhtmltopdf)
#     pdfkit.from_file(report_file, output_path = f'{folder}\report.pdf', configuration = config)





def write_css(file):
    line = """

    <STYLE>
        table {
            border: 1px solid black;
            margin: 10px 0;
        }
        th, td {
            border: 1px solid black;
            text-align: left;
            font-weight: normal;
            padding: 2px;
        }
        #thick-border {
            border-width: 3px;
        }
        #center-text {
            font-size: 20px;
            padding: 3px;
            text-align: center;
        }
        .adjustable-table {

        }

    </STYLE>
"""
    file.write(line)


if __name__ == '__main__':
    main()
