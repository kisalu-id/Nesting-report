import configparser
from confdlg import ConfigHelperINI, ConfigParamType
import ewd


def run_config():
    ini_path = ewd.explode_file_path(r"%SETTINGPATH%\script")
    ini_path += "\\config.ini"

    config = configparser.ConfigParser()
    config.read(ini_path)
    cfg_file = open(ini_path, 'w')
    config.write(cfg_file)
    cfg_file.close()
    cfg = ConfigHelperINI(ini_path, width=920, height=400, category_size=140)

    cfg.add_parameter('Druckeinstellungen', 'Druckeinstellungen\\do_report', 'Möchten Sie den Nesting-Report durchführen?', ConfigParamType.BOOLEAN, False)
    cfg.add_parameter('Druckeinstellungen', 'Druckeinstellungen\\nice_design', 'Deaktivieren für schwarz-weiß drucken, aktivieren für ausgefallenes Design', ConfigParamType.BOOLEAN, False)
    cfg.add_parameter('Druckeinstellungen', 'Druckeinstellungen\\remove_color_fill', 'Farbfüllung des Details entfernen (um Druckertinte zu sparen)', ConfigParamType.BOOLEAN, False)
    cfg.add_parameter('Druckeinstellungen', 'Druckeinstellungen\\reports_pdfs_together', 'Effizienzbericht zusammen mit Gesamteffizienzbericht in einem PDF', ConfigParamType.BOOLEAN, False)
    cfg.add_parameter('Druckeinstellungen', 'Druckeinstellungen\\divide_material', 'Teile den Bericht nach Material und Dicke auf', ConfigParamType.BOOLEAN, False)
    cfg.add_parameter('Druckeinstellungen', 'Druckeinstellungen\\rotate', 'Möchten Sie die Platten hochkant drehen?', ConfigParamType.BOOLEAN, False)

    cfg.add_parameter('Pfad', 'Pfad\\report_pfad', 'Report Pfad wählen', ConfigParamType.DIRECTORY, ewd.explode_file_path('%TEMPPATH%'))


    cfg.run()

