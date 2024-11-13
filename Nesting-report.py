import os
import datetime
import configparser
import ewd
import subprocess
from ewd import groups
from company import gdb
from company import dlg
from company import nest
from company import cad
from company import view
import sclcore
from sclcore import do_debug
from sclcore import execute_command_bool as exec_bool
import config



class ReportSheet():
    """
    Represents a report entry for a specific sheet, containing key metrics and an associated image.

    :param sheet: the name of the sheet
    :type sheet: str
    :param mat_leftover: the percentage of material that is not reusable for this sheet
    :type mat_leftover: float
    :param mat_reusable: the percentage of material that is reusable for this sheet
    :type mat_reusable: float
    :param area: the total area of the sheet in square meters
    :type area: float
    :param counter_sheet_in_sheets: the index of this sheet among sheets in sheets_to_report (sheets_to_report are values of the key [material, thickness])
    :type counter_sheet_in_sheets: int
    :param img_path: the file path to an image of the sheet
    :type img_path: str
    """
    def __init__(self, sheet, mat_leftover, mat_reusable, area, counter_sheet_in_sheets, img_path):
        self.sheet = sheet
        self.mat_leftover = mat_leftover
        self.mat_reusable = mat_reusable
        self.area = area
        self.counter_sheet_in_sheets = counter_sheet_in_sheets   #index of this sheet among sheets_to_report
        self.img_path = img_path


class MaterialStats:
    """
    Represents the statistics for sheets from a specific material-thickness pair used in a project.

    Attributes:
        material (str): the name of material used
        thickness (float): the thickness of the material in millimeters
        number_of_sheets (int): the number of sheets in sheets_to_report (sheets_to_report are values of the key [material, thickness])
        total_area (float): the total area of the sheets in sheets_to_report in square meters
        total_reusable (float): the total percentage of reusable material across sheets in sheets_to_report
        total_garbage (float): the total percentage of non-reusable (garbage) material across sheets in sheets_to_report
        average_reusable (float): the average percentage of reusable material per sheet
        average_garbage (float): the average percentage of non-reusable material per sheet
        total_reusable_material (float): the total area of reusable material per sheet (in m²)
        total_non_reusable_material(float): the total area of non-reusable material per sheet (in m²)
    """

    def __init__(self, material, thickness, number_of_sheets, total_area, total_reusable, total_garbage):
        """
        Initializes the MaterialStats object with basic information about the sheets from a specific material-thickness pair
        and calculates derived values, such as reusable and garbage material percentages.

        :param material: the name of material
        :type material: str
        :param thickness: the thickness of the material in millimeters
        :type thickness: float
        :param number_of_sheets: the number of sheets in sheets_to_report (sheets_to_report are values of the key [material, thickness])
        :type number_of_sheets: int
        :param total_area: the total area of the sheets in square meters
        :type total_area: float
        :param total_reusable: the total percentage of reusable material across sheets in sheets_to_report
        :type total_reusable: float
        :param total_garbage: the total percentage of non-reusable (garbage) material across sheets in sheets_to_report
        :type total_garbage: float
        """
        self.material = material
        self.thickness = thickness
        self.number_of_sheets = number_of_sheets
        self.total_area = total_area
        self.total_reusable = total_reusable
        self.total_garbage = total_garbage
        
        self.average_reusable = total_reusable / number_of_sheets
        self.average_garbage = total_garbage / number_of_sheets
        self.total_reusable_material = total_area * total_reusable / 100 / number_of_sheets
        self.total_non_reusable_material= total_area * total_garbage / 100 / number_of_sheets


    def GEB_to_html(self, html_file_object, project_name, logo, nice_design, i):
        """
        Generates HTML for one material-thickness total efficiency report with statistics.

        :param html_file_object: the file object to which the HTML content will be written
        :type html_file_object: file-like object
        :param project_name: the name of the project to be included in the report
        :type project_name: str
        :param logo: path to the company logo file to be included in the report
        :type logo: str
        :param nice_design: specifies if the report should have a nice colorful design (True) or a simple, black-and-white design (False) -- it's not used in this function for now, but it may be used later
        :type nice_design: bool
        :param i: index used to determine when to insert a page break, to iterate through each html_file_object in a list
        :type i: int
        """

        if i == 0 or i % 4 == 0:
            line = '\n<DIV class="page-break-after"></DIV>\n'
            html_file_object.write(line)

        if i == 0 or i % 4 == 0:
            line = '<HEADER style="display: block; width: 100%; text-align: left;">\n'
            line += f'<IMG src="file:///{logo}" alt="company Logo" style="vertical-align: middle; width: 60px; height: 60px; margin: 0 10px 15px 0;">\n'
            line += f'<SPAN style="font-size: 35px; padding: 0 0 8px 0; ">Projekt: {os.path.splitext(project_name)[0]} </SPAN>\n'
            line += '</HEADER>\n'
            html_file_object.write(line)

        line = f"""
    <TABLE id="thick-border" class="adjustable-table">
        <TH colspan="3" class="center-text">Gesamtwirkungsgradbericht</TH>

        <TR>
            <TH align="left">Material und Dicke</TH>
            <TD colspan="2" align="left">{self.material}   {self.thickness} mm</TD>
        </TR>

        <TR>
            <TH align="left">Anzahl Sheets</TH>
            <TH colspan="2" align="left">{self.number_of_sheets}</TH>
        </TR>

        <TR>
            <TD>Gesamtfläche</TD>
            <TH colspan="2" align="left">{round(self.total_area, 2)} m²</TH>
        </TR>

        <TR>
            <TD class="green">Gesamt wiederverwendbares Material</TD>
            <TD class="green">{round(self.total_reusable_material, 2)} m²</TD>
            <TD class="green td-right">{round(self.total_reusable_material / self.total_area * 100, 2)} %</TD>
        </TR>            
        <TR>
            <TD class="grey">Gesamt nicht wiederverwendbares Material</TD>
            <TD class="grey">{round(self.total_non_reusable_material, 2)} m²</TD>
            <TD class="grey td-right">{round(self.total_non_reusable_material/ self.total_area * 100, 2)} %</TD>
        </TR>

        <TR>
            <TD class="green">Wiederverwendbares Material pro Platte im Durchschnitt:</TD>
            <TD class="green">{round(self.total_area * self.average_reusable / 100 / self.number_of_sheets, 2)} m²</TD>
            <TD class="green td-right">{round(self.total_reusable / self.number_of_sheets, 2)}%</TD>
        </TR>

        <TR>
            <TD class="grey">Nicht wiederverwendbares Material pro Platte im Durchschnitt</TD>
            <TD class="grey">{round(self.total_area * self.average_garbage / 100 / self.number_of_sheets, 2)} m²</TD>
            <TD class="grey td-right">{round(self.total_garbage / self.number_of_sheets, 2)}%</TD>
        </TR>
    </TABLE>
            """
        html_file_object.write(line)



def nesting_report():
    """
    The program loads settings from an .ini file, creates a new folder, and generates an HTML report with applied CSS. 
    It includes detailed information about each sheet, calculates individual and total efficiency metrics, and then converts the final HTML report(s) into a PDF format.
    """

    #do_debug()
    rotate, general_folder, nice_design, reports_pdfs_together, only_measures_BW, divide_material = read_config_ini()

    ###if project is not saved --> save it in temp folder or in a temp dircetory in the config file
    project_name = get_or_create_project_name()

    folder = make_or_delete_folder(general_folder, project_name)

    report_file_path = create_report_file_path(folder, project_name)

    img_ext = ".jpg"
    logo = "C:\...\company_logo.png"

    set_view_and_shading(nice_design)

    #try:
    #what i do:
    #sort sheets depending on material and thickness
    #for each materials_dict[key] generate html

    materials_dict, total_sheets_amount = sort_for_material() #basically sorting and saving as a dictionary

    materials_stats_list = []

    counter_for_full_pdf = 0
    counter_sheet_in_sheets = 0   #index of this sheet among sheets_to_report

    for material_and_thickness, sheets_values in materials_dict.items():
        material, thickness = material_and_thickness  #extract material and thickness from the key

        material_stats_obj = create_object_material_stats(material_and_thickness, sheets_values)
        materials_stats_list.append(material_stats_obj)

        #html with a new name
        if divide_material:
            project_name_mat_thick = f"{project_name}_{material}_{thickness}"    #.replace('.ewd', '')
            report_file_path = os.path.join(folder, f"{project_name_mat_thick}.html")

            try:
                os.makedirs(os.path.dirname(report_file_path), exist_ok=True)
            except OSError as e:
                dlg.output_box(f"Fehler beim Ordner erstellen in {os.path.dirname(report_file_path)}")

            with open(report_file_path, 'w', encoding='utf-8') as html_file:
                try:
                    create_report(report_file_path, html_file, project_name_mat_thick, folder, img_ext, logo, nice_design, rotate, reports_pdfs_together, divide_material, sheets_values, total_sheets_amount, materials_dict, counter_for_full_pdf, 0)
                    test = 3
                    
                    if reports_pdfs_together:
                        material_stats_obj.GEB_to_html(html_file, project_name, logo, nice_design, 0)
                        close_html(html_file)

                    else: #if not reports_pdfs_together
                        close_html(html_file)

                except Exception as e:
                    raise e
            output_pdf = os.path.join(folder, f'{project_name_mat_thick}.pdf')
            to_pdf(report_file_path, output_pdf)
        
        else: #if not divide_material:    - _in_ the loop of material_and_thickness
            if counter_for_full_pdf == 0: 
                with open(report_file_path, 'w', encoding='utf-8') as html_file:
                    create_report(report_file_path, html_file, project_name, folder, img_ext, logo, nice_design, rotate, reports_pdfs_together, divide_material, sheets_values, total_sheets_amount, materials_dict, counter_for_full_pdf, counter_sheet_in_sheets)
                    test = 6

            else:  #if counter_for_full_pdf != 0: 
                with open(report_file_path, 'a', encoding='utf-8') as html_file:
                    create_report(report_file_path, html_file, project_name, folder, img_ext, logo, nice_design, rotate, reports_pdfs_together, divide_material, sheets_values, total_sheets_amount, materials_dict, counter_for_full_pdf, counter_sheet_in_sheets)
                    test = 5
            counter_for_full_pdf +=1

    # _after_ the loop of material_and_thickness
    if not divide_material:

        if reports_pdfs_together: #write GEB in the same big PDF at the end
            with open(report_file_path, 'a', encoding='utf-8') as html_file:
                for i, material_stats_obj in enumerate(materials_stats_list):
                    material_stats_obj.GEB_to_html(html_file, project_name, logo, nice_design, i)
                close_html(html_file)

            output_pdf = os.path.join(folder, f'{project_name}.pdf')
            to_pdf(report_file_path, output_pdf)

        else:   #if not reports_pdfs_together:     #write GEB in the separate PDF at the end
            output_pdf = os.path.join(folder, f'{project_name}.pdf')
            to_pdf(report_file_path, output_pdf)

            report_file_path = os.path.join(folder, f'Gesamteffizienbericht_{project_name}.html')

            with open(report_file_path, 'w', encoding='utf-8') as html_file_GEB:

                html_header_and_css(html_file_GEB, project_name, nice_design)     
                for i, material_stats_obj in enumerate(materials_stats_list):
                    material_stats_obj.GEB_to_html(html_file_GEB, project_name, logo, nice_design, i)
                close_html(html_file_GEB)

                output_pdf = os.path.join(folder, f'Gesamteffizienbericht_{project_name}.pdf')
            to_pdf(report_file_path, output_pdf)

    if divide_material and not reports_pdfs_together:
        report_file_path = os.path.join(folder, f'Gesamteffizienbericht_{project_name}.html')
        with open(report_file_path, 'w', encoding='utf-8') as html_file_GEB:

            html_header_and_css(html_file_GEB, project_name, nice_design)     
            for i, material_stats_obj in enumerate(materials_stats_list):
                material_stats_obj.GEB_to_html(html_file_GEB, project_name, logo, nice_design, i)
            close_html(html_file_GEB)

        output_pdf = os.path.join(folder, f'Gesamteffizienbericht_{project_name}.pdf')
        to_pdf(report_file_path, output_pdf)



def create_object_material_stats(material_and_thickness, sheets_values):
    """
    Creates a MaterialStats object for a specific material-thickness pair, calculating total area, reusable material,
    and garbage material from the provided sheet data. Extracts the area, reusable material percentage, and leftover 
    material percentage for each sheet.

    :param material_and_thickness: a tuple containing the material and thickness.
    :type material_and_thickness: tuple (str, float)
    :param sheets_values: a list of sheets, where properties such as area, reusable material, and garbage material are extracted.
    :type sheets_values: list
    :return: a MaterialStats object representing the statistics for the material-thickness pair.
    :rtype: MaterialStats
    """

    material, thickness = material_and_thickness
    number_of_sheets = len(sheets_values)

    total_area = 0
    total_reusable = 0
    total_garbage = 0
    
    for sheet in sheets_values:
        area = nest.get_sheet_property(sheet, nest.SheetProperties.AREA)
        mat_reusable = nest.get_sheet_property(sheet, nest.SheetProperties.RATE_REUSABLE)  # % of sheet reusable material
        mat_leftover = nest.get_sheet_property(sheet, nest.SheetProperties.RATE_LEFT_OVER)  # % of sheet garbage not reusable material
        area = area / 1000000  #to m²

        total_area += area
        total_reusable += mat_reusable
        total_garbage += mat_leftover

    return MaterialStats(
        material=material,
        thickness=thickness,
        number_of_sheets=number_of_sheets,
        total_area=total_area,
        total_reusable=total_reusable,
        total_garbage=total_garbage
    )


def run_config():
    """
    Executes the configuration process by calling the run_config() method from the config module.
    This is triggered when a user opens nesting and clicks on the 'Report config' button.
    """
    config.run_config()


def read_config_ini():
    """
    Returns a set of configuration parameters used for generating reports.

    :return: a tuple containing:
        - rotate: whether the pages should be 90° rotated (bool)
        - general_folder: where the folder where reports will be stored (str)
        - nice_design: whether the report will have a visually appealing design (bool)
        - reports_pdfs_together: whether sheet report and material efficiency report will be combined into a single PDF (bool)
        - only_measures_BW: whether the report will contain only measurements in black and white (bool)
        - divide_material: whether the materials should be divided into separate sections (bool)
    :rtype: tuple
    """

    try:
        ini_path = ewd.explode_file_path(r'%MACHPATH%\script\config.ini')

        config = configparser.ConfigParser()
        config.read(ini_path)

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
        
        rotate = config.get('Druckeinstellungen', 'rotate') #if True, rotate
        rotate = False if rotate == "0" or rotate =="False" else True

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
    ###if project is not saved --> save it in temp folder in EW or in a temp dircetory in the config file
    project_name = ewd.get_project_name() #get the name of the opened ewd project
    if not project_name.endswith (".ewd"):
        project_name = datetime.datetime.now().strftime("%Y%m%d_%H-%M")
        ewd.save_project(ewd.explode_file_path(f"%TEMPPATH%//{project_name}.ewd")) # C:\...\Temp
    else:
        project_name = project_name.replace('.ewd', '')
    return project_name



def make_or_delete_folder(general_folder, project_name):
    """
    Creates a folder for the project if it doesn't exist, or deletes and recreates it if it already exists.

    - If a folder already exists at the specified path, it is deleted and recreated to ensure a clean directory.
    - If the folder does not exist, it will be created using the provided `general_folder` and `project_name`.

    :param general_folder: The base directory where the project folder will be created or recreated.
    :type general_folder: str
    :param project_name: The name of the project, which will be used as the folder name.
    :type project_name: str
    :return: The full path to the created (or recreated) project folder.
    :rtype: str
    """
    delete_folder = 1
    #subfolder with the project name: it will be deleted, if it already exists, then the new folder will be created
    #(so the date of creation of this folder on user's pc will be fresh -> easy to sort)
    folder = os.path.join(general_folder, f'{os.path.splitext(project_name)[0]}')

    #if folder exists (i set this setting as always on), delete
    if delete_folder and os.path.exists(folder):
        #maybe later add "ok" and "cancel"
        remove_existing_folder_with_same_name(folder)

    else: #if folder doesn't exist, create
        try:
            os.makedirs(folder, exist_ok=False)
        except OSError as e:
            dlg.output_box(f"Fehler beim Ordner erstellen in {folder}")
    return folder


def remove_existing_folder_with_same_name(folder):
    """
    If delete_folder is True, remove the existing folder with the same name.

    :param folder: directory where the generated HTML and PDF reports will be saved
    :type folder: str
    """
    try:
        if os.path.exists(folder):
            #dlg.output_box(f"Der Ordner '{folder}' und sein Inhalt werden gelöscht")
            # maybe add ok / cancel in the config
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


def create_report_file_path(folder, poject_name):
    """
    Generates the file path for the HTML report file and ensures the folder exists.

    :param folder: the directory where the report will be saved
    :type folder: str
    :param project_name: the name of the project, used to create the report filename
    :type project_name: str
    :return: the full file path for the HTML report
    :rtype: str
    """
    report_file_path = os.path.join(folder, f'{poject_name}.html')
    try:
        os.makedirs(os.path.dirname(report_file_path), exist_ok=True)
    except OSError as e:
        dlg.output_box(f"Fehler beim Ordner erstellen in {os.path.dirname(report_file_path)}")
    return report_file_path


def set_view_and_shading(remove_color_fill):
    """
    Sets the viewing perspective and shading options for the design environment.
    This function switches to a top view in wireframe mode. If `nice_design` is True, it applies shading to the view.

    :param nice_design: specifies whether to use a nicer design with shading (True) or a simple wireframe view (False)
    :type nice_design: bool
    """
    #switch to top view;
    view.set_std_view_eye()

    if remove_color_fill:
        exec_bool("SetWireFrame")
    else:
        exec_bool("SetShading")



def sort_for_material():
    """
    Iterates through each sheet and creates a dictionary with key-value pairs to sort the sheets by material.
    The function collects information about each sheet and organizes it into a dictionary where the keys 
    are tuples of (material, thickness) and the values are lists of sheets corresponding to those keys. 
    Additionally, it calculates the total number of sheets in a project.

    :return: a tuple containing:
        - materials_dict (dict): a dictionary with materials as keys and lists of sheets as values
        - total_sheets_amount (int): the total number of sheets processed
    """
    materials_dict = {}
    sheets = nest.get_sheets()
    total_sheets_amount = len(sheets)
    for sheet in sheets:
        material = nest.get_sheet_property(sheet, nest.SheetProperties.MATERIAL)
        thickness = nest.get_sheet_property(sheet, nest.SheetProperties.THICKNESS)
        key = (material, thickness) #tuple
        if key in materials_dict:
            materials_dict[key].append(sheet)
        else:
            #materials_dict[(material, thickness)] = []   #tuple = list
            materials_dict[key] = [sheet]
    return materials_dict, total_sheets_amount


def create_report(report_file_path, html_file, project_name, folder, img_ext, logo, nice_design, rotate, reports_pdfs_together, divide_material, sheets_to_report, total_sheets_amount, materials_dict, counter_for_full_pdf, counter_sheet_in_sheets):
    """
    Creates an HTML report by calling necessary functions and handles sheet rotation if required.
    This function is intended to be called exclusively from the nesting_report() function, allowing for different parameters 
    to be passed based on the specific situation.
    
    :param report_file_path: the file path where the report will be saved
    :type report_file_path: str
    :param html_file: the HTML file object to which content will be written
    :type html_file: file-like object
    :param project_name: the name of the project for inclusion in the report
    :type project_name: str
    :param folder: the folder where the report will be located
    :type folder: str
    :param img_ext: the file extension for images included in the report (e.g., '.jpg')
    :type img_ext: str
    :param logo: the file path to the logo image to be included in the report
    :type logo: str
    :param nice_design: specifies if the report should use a nice design (True) or a simple design (False)
    :type nice_design: bool
    :param rotate: indicates whether to rotate the sheets (True) or not (False)
    :type rotate: bool
    :param reports_pdfs_together: whether sheet report and material efficiency report will be combined into a single PDF
    :type reports_pdfs_together: bool
    :param divide_material: indicates if the report should be divided on separate PDFs by material types, or should it be written in a single PDF
    :type divide_material: bool
    :param sheets_to_report: a collection of sheets to be included in the report
    :type sheets_to_report: list
    :param total_sheets_amount: the total number of sheets from the project
    :type total_sheets_amount: int
    :param materials_dict: a dictionary with keys as tuples of (material, thickness) and values as lists of corresponding sheets
    :type materials_dict: dict
    :param counter_for_full_pdf: a counter, that indicates to write the HTML header and CSS at the beginning of the PDF, if counter_for_full_pdf = 0
    :type counter_for_full_pdf: int
    :param counter_sheet_in_sheets: an index for tracking the current sheet within sheets_to_report
    :type counter_sheet_in_sheets: int
    """
    

    try:
        # #if reports_pdfs_together and divide_material:
        # #I ALWAYS count stuff separately for material_thickness
        #i need to have these variables as 0 for each key only if reports are split for material, 
        #     or if that's the first key (pdf) for reports_pdfs_together=True
        #it could be if divide_material or counter_for_full_pdf == 0: , 
        #     but it's the same thing since in nesting_report if divide_material  I'm always passing 0 for counter_for_full_pdf
        if counter_for_full_pdf == 0: 

            #HTML header, file name and css
            html_header_and_css(html_file, project_name, nice_design)

        if rotate: #rotate sheets by 90 degrees
            for sheet in sheets_to_report:
                cad.rotate(sheet, 0, 0, -90, False)


        for sheet in sheets_to_report:
            sheet_obj = get_sheet_obj(folder, sheet, counter_sheet_in_sheets, img_ext)

            counter_sheet_in_sheets = write_html(folder, html_file, logo, project_name, sheets_to_report, reports_pdfs_together, nice_design, divide_material, sheet_obj, total_sheets_amount)

            if rotate:
                cad.rotate(sheet, 0, 0, 90, False)

    except IOError as e:
        dlg.output_box(f"Ein Fehler ist beim Schreiben der Datei '{report_file_path}' aufgetreten: {e}")
    except Exception as e:
        dlg.output_box(f"Ein unerwarteter Fehler ist aufgetreten: {e}")



def get_sheet_obj(folder, sheet, counter_sheet_in_sheets, img_ext):
    """
    Creates a ReportSheet object for a given sheet by extracting its properties 
    and generating the corresponding image path.

    :param folder: the folder where the sheet images are stored
    :type folder: str
    :param sheet: the name of the sheet
    :type sheet: str
    :param counter_sheet_in_sheets: the index of this sheet among sheets_to_report
    :type counter_sheet_in_sheets: int
    :param img_ext: the file extension for the sheet's image (e.g., '.jpg')
    :type img_ext: str

    :return: a ReportSheet object containing the sheet's details
    :rtype: ReportSheet
    """
    img_path = f"{folder}\{sheet}{img_ext}"
    area = nest.get_sheet_property(sheet, nest.SheetProperties.AREA)
    mat_leftover = nest.get_sheet_property(sheet, nest.SheetProperties.RATE_LEFT_OVER)      # % of sheet   garbage not reusable material
    mat_reusable = nest.get_sheet_property(sheet, nest.SheetProperties.RATE_REUSABLE)       # % of sheet   reusable material
    area = (area / 1000000)   # m2

    if not os.path.isfile(img_path):
        os.makedirs(os.path.dirname(img_path), exist_ok=True)

    view.zoom_on_object(sheet, ratio=1)
    nest.get_sheet_preview(sheet, img_path, 0.3) # 0.3, so the lines will be thicker

    return ReportSheet(
        sheet=sheet,
        mat_leftover=mat_leftover,
        mat_reusable=mat_reusable,
        area=area,
        counter_sheet_in_sheets=counter_sheet_in_sheets,   #index of this sheet among sheets_to_report
        img_path=img_path,
    )


def html_header_and_css(html_file_object, project_name, nice_design):
    """
    Write HTML header and chosen CSS style. 

    :param html_file_object: the file object to which the HTML content will be written
    :type html_file_object: file-like object (e.g., obtained via open() in write mode)
    :param project_name: the name of the project
    :type project_name: str
    :param nice_design: specifies if the report should use a nice design (True) or a simple design (False)
    :type nice_design: bool
    """

    line = f"""<!DOCTYPE HTML>
<HTML lang="de">
<HEAD>
    <META charset="UTF-8">
    <META name="viewport" content="width=device-width, initial-scale=1.0">
    <TITLE>Nesting report for {project_name}</TITLE>
    """
    html_file_object.write(line)

    if nice_design:
        write_nice_css(html_file_object)
    else:
        write_css_printing(html_file_object)



def write_nice_css(html_file_object):
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
            display:inline-block;
            page-break-inside: avoid;
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
            display:inline-block;
        }

        .mainTable {
            display: table;
            width: 100%;
        }

        .adjustable-table {
            width: 100%;
        }

        .page-break-after {
            page-break-after: always;
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


def write_html(folder, html_file_object, logo, project_name, sheets_to_report, reports_pdfs_together, nice_design, divide_material, sheet_obj, total_sheets_amount):
    """
    Gets the sheet properties, writes HTML file, includes piece properties, 
    and counts the efficiency for each sheet as well as the total efficiency.

    :param folder: the folder where the report will be saved
    :type folder: str
    :param html_file_object: the file object to which the HTML content will be written
    :type html_file_object: file-like object
    :param logo: the path to the logo to be included in the report
    :type logo: str
    :param project_name: the name of the project for inclusion in the report
    :type project_name: str
    :param sheets_to_report: a collection of sheets to be included in the report
    :type sheets_to_report: list
    :param reports_pdfs_together: whether the sheet report and material efficiency report will be combined into a single PDF (bool)
    :type reports_pdfs_together: bool
    :param nice_design: specifies if the report should use a nice design (True) or a simple design (False)
    :type nice_design: bool
    :param divide_material: indicates if the report should be divided into separate PDFs by material types, or if it should be written in a single PDF
    :type divide_material: bool
    :param sheet_obj: the ReportSheet object containing details for the current sheet
    :type sheet_obj: ReportSheet
    :param total_sheets_amount: the total number of sheets being reported on
    :type total_sheets_amount: int

    :return: the updated counter of the current sheet in sheets_to_report
    :rtype: int
    """
    sheet = sheet_obj.sheet   #count 0 for 1st sheet
    mat_leftover = sheet_obj.mat_leftover
    mat_reusable = sheet_obj.mat_reusable
    area = sheet_obj.area
    counter_sheet_in_sheets = sheet_obj.counter_sheet_in_sheets     # the number of sheet in sheets_to_report
    img_path = sheet_obj.img_path


    pieces = nest.get_sheet_property(sheet, nest.SheetProperties.PIECES_NUMBER)

    write_sheet_info_and_picture(sheet, html_file_object, logo, counter_sheet_in_sheets, img_path, project_name, sheets_to_report, reports_pdfs_together, divide_material, total_sheets_amount)

    #write down the individual information about the pieces on a sheet
    write_pieces_info(sheet, html_file_object)

    line = '</TABLE>\n' #closing mainTable
    html_file_object.write(line)


    efficiency_for_sheet(html_file_object, pieces, area, mat_leftover, mat_reusable)

    #dlg.output_box(f"Ein Fehler ist beim Schreiben der Datei '{total_report_path}' aufgetreten: {e}")

    line = '    </DIV>\n' #closing <DIV class="table-container">

    html_file_object.write(line)
    counter_sheet_in_sheets += 1  #the number of sheet in sheets_to_report
    return counter_sheet_in_sheets


def write_sheet_info_and_picture(sheet, html_file_object, logo, counter_sheet_in_sheets, img_path, project_name, sheets_to_report, reports_pdfs_together, divide_material, total_sheets_amount):
    #write logo, project name, sheet picture, sheet stats (material, thickness, width, height, current_date)
    """
    Writes the individual sheet's information, including logo, project name, sheet picture, 
    and various statistics (material, thickness, width, height, current date) 
    to the HTML file.

    :param sheet: the name of the sheet
    :type sheet: str
    :param html_file_object: the file object to which the HTML content will be written
    :type html_file_object: file-like object
    :param logo: the path to the logo to be included in the report
    :type logo: str
    :param counter_sheet_in_sheets: the index of this sheet among sheets_to_report
    :type counter_sheet_in_sheets: int
    :param img_path: the file path to the image of the sheet
    :type img_path: str
    :param project_name: the name of the project for inclusion in the report
    :type project_name: str
    :param sheets_to_report: a collection of sheets to be included in the report
    :type sheets_to_report: list
    :param reports_pdfs_together: whether the sheet report and material efficiency report will be combined into a single PDF (bool)
    :type reports_pdfs_together: bool
    :param divide_material: indicates if the report should be divided into separate PDFs by material types, or if it should be written in a single PDF
    :type divide_material: bool
    :param total_sheets_amount: the total number of sheets being reported on
    :type total_sheets_amount: int
    """

    material = nest.get_sheet_property(sheet, nest.SheetProperties.MATERIAL)
    thickness = nest.get_sheet_property(sheet, nest.SheetProperties.THICKNESS)
    width = nest.get_sheet_property(sheet, nest.SheetProperties.WIDTH)
    height = nest.get_sheet_property(sheet, nest.SheetProperties.HEIGHT)
    current_date = datetime.datetime.now().strftime("%d.%m.%Y")

    if not divide_material:
        amount_of_sheets_to_report = total_sheets_amount
    if divide_material:
        amount_of_sheets_to_report = len(sheets_to_report)
    
    line = ' '

    if not divide_material and not reports_pdfs_together and (counter_sheet_in_sheets + 1) < amount_of_sheets_to_report:    # and reports_pdfs_together and last_sheet == sheet (or if counter == amount_of_sheets_to_report):   # and reports_pdfs_together:
        line += '\n<DIV class="page-break-after"></DIV>\n'
    elif not divide_material and reports_pdfs_together:
        line += '\n<DIV class="page-break-after"></DIV>\n'
    #elif not divide_material and not reports_pdfs_together:
    elif divide_material:
        if counter_sheet_in_sheets < amount_of_sheets_to_report:
            line += '\n<DIV class="page-break-after"></DIV>\n'
            

    line += '    <HEADER style="display: inline-block; width: 100%; text-align: left;">\n'
    line += f'        <IMG src="file:///{logo}" alt="company Logo" style="vertical-align: middle; width: 60px; height: 60px; margin: 0 10px 15px 0;">\n'
    line += f'        <SPAN style="font-size: 35px; padding: 0 0 8px 0;">Projekt: {os.path.splitext(project_name)[0]} </SPAN>\n'
    line += '    </HEADER>\n'
    
    #table for sheet information
    line += '\n    <DIV class="table-container">\n'
    line += '    <TABLE class="mainTable">\n'
    html_file_object.write(line)

    line = f'        <TR>\n            <TD style="font-size:30px" colspan="6">{sheet}</TD>\n'

    line += f'            <TD colspan="4" class="right-align">{current_date}</TD>\n        </TR>\n'
    html_file_object.write(line)

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
    html_file_object.write(line)

    #picture from the sheet
    size_img = "width=\"1200pt\"" if width > 3 * height else "height=\"400pt\""
    line = f'   <TR>\n      <TD colspan="10">\n         <IMG src="file:///{img_path}" {size_img}>\n     </TD>\n        </TR>\n'
    html_file_object.write(line)


def write_pieces_info(sheet, html_file_object):
    #write down the individual information about the pieces on the sheet (n_piece_count, piece_label, piece_width,piece_height)
    """
    Writes individual information about the pieces on the specified sheet 
    to the HTML file.

    :param sheet: the name of the sheet containing the pieces
    :type sheet: str
    :param html_file_object: the file object to which the HTML content will be written
    :type html_file_object: file-like object
    """
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
        html_file_object.write(line)
        n_piece_count += 1


















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
    path_to_wkhtmltopdf = r'C:\...\wkhtmltopdf.exe'
    subprocess.call([path_to_wkhtmltopdf, report_html_path, output_pdf_path])




if __name__ == '__main__':
    nesting_report()
