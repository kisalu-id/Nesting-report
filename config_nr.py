import configparser
from confdlg import ConfigHelperINI, ConfigParamType
import ewd


def run_config():
    ini_path = ewd.explode_file_path(r"%MACHPATH%\script") #!
    ini_path += "\\config_nr.ini"

    config = configparser.ConfigParser()
    config.read(ini_path)
    cfg_file = open(ini_path, 'w')
    config.write(cfg_file)
    cfg_file.close()
    cfg = ConfigHelperINI(ini_path)

    cfg.add_parameter('Druckeinstellungen', 'Druckeinstellungen\\nice_design', 'Deaktivieren für schwarz-weiß drucken, aktivieren für ausgefallenes Design', ConfigParamType.BOOLEAN, False)
    cfg.add_parameter('Druckeinstellungen', 'Druckeinstellungen\\reports_pdfs_together', 'Effizienzbericht zusammen mit Gesamteffizienzbericht in einem PDF', ConfigParamType.BOOLEAN, False)
    cfg.add_parameter('Druckeinstellungen', 'Druckeinstellungen\\only_measures_BW', '*Funktion in Entwicklung* Nur Maßnahmenbericht', ConfigParamType.BOOLEAN, False)
    cfg.add_parameter('Druckeinstellungen', 'Druckeinstellungen\\divide_material', 'Teile den Bericht nach Material und Dicke auf', ConfigParamType.BOOLEAN, False)
    cfg.add_parameter('Druckeinstellungen', 'Druckeinstellungen\\rotate', 'Möchten Sie die Platten hochkant drehen?', ConfigParamType.BOOLEAN, False)


    cfg.add_parameter('Druckeinstellungen', 'Druckeinstellungen\\mini_report', 'Mini-report', ConfigParamType.BOOLEAN, False)

    cfg.add_parameter('Pfad', 'Pfad\\report_pfad', 'Report Pfad wählen', ConfigParamType.DIRECTORY, ewd.explode_file_path('%TEMPPATH%'))


    cfg.run()


if name == 'main':
    run_config()
