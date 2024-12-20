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

def run_config():
    ini_path = ewd.explode_file_path(r"%SETTINGPATH%\script")
    ini_path += "\\config.ini"

    config = configparser.ConfigParser()
    config.read(ini_path)
    cfg_file = open(ini_path, 'w')
    config.write(cfg_file)
    cfg_file.close()
    cfg = ConfigHelperINI(ini_path, width=920, height=400, category_size=140)

    cfg.add_parameter('Druckeinstellungen', 'Druckeinstellungen\\do_report', 'Nesting-Report durchführen', ConfigParamType.BOOLEAN, False)
    cfg.add_parameter('Druckeinstellungen', 'Druckeinstellungen\\nice_design', 'Aktivieren für ausgefallenes Design, deaktivieren für schwarz-weiß drucken', ConfigParamType.BOOLEAN, False)
    cfg.add_parameter('Druckeinstellungen', 'Druckeinstellungen\\remove_color_fill', 'Farbfüllung des Details entfernen (um Druckertinte zu sparen)', ConfigParamType.BOOLEAN, False)
    cfg.add_parameter('Druckeinstellungen', 'Druckeinstellungen\\reports_pdfs_together', 'Effizienzbericht zusammen mit Gesamteffizienzbericht in einem PDF', ConfigParamType.BOOLEAN, False)
    cfg.add_parameter('Druckeinstellungen', 'Druckeinstellungen\\divide_material', 'Teile den Bericht nach Material und Dicke auf', ConfigParamType.BOOLEAN, False)
    cfg.add_parameter('Druckeinstellungen', 'Druckeinstellungen\\rotate', 'Die Platten hochkant drehen', ConfigParamType.BOOLEAN, False)
    cfg.add_parameter('Druckeinstellungen', 'Druckeinstellungen\\show_warning_delete_folder', 'Meldung anzeigen, wenn der bestehende Ordner gelöscht wird', ConfigParamType.BOOLEAN, True)
    
    cfg.add_parameter('Pfad', 'Pfad\\report_pfad', 'Report Pfad wählen', ConfigParamType.DIRECTORY, ewd.explode_file_path('%TEMPPATH%'))

    cfg.add_parameter('Automatisch öffnen', 'Automatisch öffnen\\auto_open', 'Die PDF-Datei(en) automatich in Browser öffnen', ConfigParamType.BOOLEAN, False)
    cfg.add_parameter('Automatisch öffnen', 'Automatisch öffnen\\open_all', 'Alle Reports automatisch in Browser öffnen (sonst - nur den Gesamteffizienzbericht(e))', ConfigParamType.BOOLEAN, False)
    cfg.add_parameter('Automatisch öffnen', 'Automatisch öffnen\\browser_path', 'Browser Pfad wählen', ConfigParamType.FILE, ewd.explode_file_path('%ROOTPATH%'))

    cfg.add_parameter('Programm wählen', 'Programm wählen\\ewd_file', 'Aktivieren für .EWD Dateien, sonst .EWB', ConfigParamType.BOOLEAN, True)

    cfg.run()


if __name__ == '__main__':
    run_config()
