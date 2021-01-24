import os
import re
from pprint import pprint
import os


def get_all_xml_files(rom_folder):
    # rom_folder = r"E:\RaspberryPi\ISO\Galisteo_Cobaltov3_64GB(DragonBlaze_V6.1)\rom"
    xml_files = []

    for root, dirs, files in os.walk(rom_folder):
        for file in files:
            if file.endswith(".xml"):
                full_path = os.path.join(root, file)
                if "jzjz"  in full_path:
                    xml_files.append(os.path.join(root, file))
                    # print(os.path.join(root, file))

    return xml_files


def export_game_list(xml_files):
    # xml_files = [r"E:\RaspberryPi\ISO\Galisteo_Cobaltov3_64GB(DragonBlaze_V6.1)\rom\n64\gamelist.xml"]

    game_dict = {}

    for xml_file in xml_files:
        f = open(xml_file, "r", encoding='utf-8')

        found = re.search('<System>(.*)</System>', f.read())
        if found:
            system_name = found.group(1)
            if system_name =="Family Computer":
                system_name="NES"
            if system_name =="Super Famicom":
                system_name="SNES"
            game_dict[system_name] = []

        f.close()

        # game_list = []
        f = open(xml_file, "r", encoding='utf-8')

        for line in f:
            # print (line)
            # print(line)

            found = re.search('<name>(.*)</name>', line)

            if found:
                # print(xml_file)
                if found.group(1) not in game_dict.get(system_name):
                    if "#" not in found.group(1):
                        if "notgame" not in found.group(1):
                            game_dict.get(system_name).append(found.group(1))
                # print(found.group(1))

    # pprint(game_dict)
    return game_dict

def export_to_sheet(game_dict):

    import xlsxwriter
    workbook = xlsxwriter.Workbook('E:\RaspberryPi\myfile.xlsx')
    worksheet = workbook.add_worksheet()
    row = 0
    col = 0

    print ("ASdf")
    order = sorted(game_dict.keys())
    for key in order:
        worksheet = workbook.add_worksheet(key[:30])
        worksheet.set_column(0, 0, 20)  # Game name column size
        worksheet.set_column(1, 1, 40)  # Game name column size
        worksheet.set_column(1, 2, 40)  # Game name column size
        worksheet.set_column(1, 3, 40)  # Game name column size
        worksheet.set_column(1, 4, 40)  # Game name column size
        worksheet.set_column('G:XFD', None, None, {'hidden': True})
        row = 0
        col = 0
        row += 1
        worksheet.write(row, col, key)
        worksheet.write(row, 1, "Name")
        worksheet.write(row, 2, "32 GB")
        worksheet.write(row, 3, "64 GB")
        worksheet.write(row, 4, "128 GB")
        row += 1

        for item in game_dict[key]:
            worksheet.write(row, col + 1, item)
            row += 1

    workbook.close()


def main():
    rom_folder = r"E:\RaspberryPi\ISO\Galisteo_Cobaltov3_128GB(DragonBlaze_V6.1)\rom"

    # export_game_list(rom_folder)
    xml_files = get_all_xml_files(rom_folder)
    game_dict = export_game_list(xml_files)

    pprint(game_dict)

    # export_to_sheet(game_dict)

if __name__ == '__main__':
    main()
