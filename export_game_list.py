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
                if "jzjz" in full_path:
                    xml_files.append(os.path.join(root, file))
                    # print(os.path.join(root, file))

    return xml_files


def export_game_list(xml_files, system_size,game_dict={}):
    # xml_files = [r"E:\RaspberryPi\ISO\Galisteo_Cobaltov3_64GB(DragonBlaze_V6.1)\rom\n64\gamelist.xml"]

    # game_dict = {}

    for xml_file in xml_files:
        f = open(xml_file, "r", encoding='utf-8')

        found = re.search('<System>(.*)</System>', f.read())
        if not found:
            continue

        system_name = found.group(1)
        if system_name == "Family Computer":
            system_name = "NES"
        if system_name == "Super Famicom":
            system_name = "SNES"
        if system_name == "Game &amp; Watch":
            system_name = "Game and Watch"
        if system_name == "Famiri Konpyuta Disuku Shisutemu":
            system_name = "fds"

        if system_name not in game_dict.keys():
            game_dict[system_name] = {}

        f.close()

        # game_list = []
        f = open(xml_file, "r", encoding='utf-8')

        for line in f:

            found = re.search('<name>(.*)</name>', line)

            if not found:
                continue

            if found.group(1) not in game_dict.get(system_name):
                if "#"  in found.group(1):
                    continue
                if "notgame"  in found.group(1):
                    continue
                game_dict.get(system_name)[found.group(1)] = []

            if system_size not in game_dict.get(system_name)[found.group(1)]:
                game_dict.get(system_name)[found.group(1)].append(system_size)

    # pprint(game_dict)
    return game_dict


def export_to_sheet(game_dict):
    import xlsxwriter
    workbook = xlsxwriter.Workbook('E:\RaspberryPi\myfile.xlsx')
    fmt_bold_2 = workbook.add_format({'bold': True, 'bg_color': 'cyan', 'color': 'black'})
    fmt_green = workbook.add_format({'bg_color': '99ff66', 'color': 'black'})

    worksheet = workbook.add_worksheet("ALL")
    row = 0
    col = 0

    print("ASdf")
    order = sorted(game_dict.keys())
    for system_name in order:
        worksheet = workbook.add_worksheet(system_name[:30])
        worksheet.set_column(0, 0, 20)  # Game name column size
        worksheet.set_column(1, 1, 40)  # Game name column size
        worksheet.set_column(1, 2, 40)  # Game name column size
        worksheet.set_column(1, 3, 40)  # Game name column size
        worksheet.set_column(1, 4, 40)  # Game name column size
        worksheet.set_column('G:XFD', None, None, {'hidden': True})
        row = 0
        col = 0
        row += 1

        worksheet.write(row, col, system_name)
        worksheet.write(row, 1, "Name",fmt_bold_2)
        worksheet.write(row, 2, "32 GB",fmt_bold_2)
        worksheet.write(row, 3, "64 GB",fmt_bold_2)
        worksheet.write(row, 4, "128 GB",fmt_bold_2)
        row += 1

        for game_name,console_sizes in game_dict[system_name].items():
            worksheet.write(row, col + 1, game_name,fmt_bold_2)

            col = 0
            for console_size in console_sizes:
                if console_size =="32":
                    worksheet.write(row, col + 2, console_size,fmt_green)
                if console_size == "64":
                    worksheet.write(row, col + 3, console_size,fmt_green)
                if console_size == "128":
                    worksheet.write(row, col + 4, console_size,fmt_green)
            #     col= col+1
            row += 1

    workbook.close()


def main():
    game_dict = {}

    rom_folder = r"E:\RaspberryPi\ISO\Galisteo_Cobaltov3_32GB(DragonBlaze_V6.1)\rom"
    xml_files_32 = get_all_xml_files(rom_folder)
    game_dict = export_game_list(xml_files_32,"32")

    rom_folder = r"E:\RaspberryPi\ISO\Galisteo_Cobaltov3_64GB(DragonBlaze_V6.1)\rom"
    xml_files_64 = get_all_xml_files(rom_folder)
    game_dict = export_game_list(xml_files_64,"64",game_dict)


    rom_folder = r"E:\RaspberryPi\ISO\Galisteo_Cobaltov3_128GB(DragonBlaze_V6.1)\rom"
    xml_files_128 = get_all_xml_files(rom_folder)
    game_dict = export_game_list(xml_files_128, "128",game_dict)

    pprint(game_dict)

    export_to_sheet(game_dict)


if __name__ == '__main__':
    main()
