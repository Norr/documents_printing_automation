import os.path
import time
import mouse
import pyautogui
import pandas as pd
import io
import holidays
import datetime
import comtypes.client
import re

from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from tkinter import Tk
from email import policy
from email.parser import BytesParser
from PyPDF2 import PdfFileWriter, PdfFileReader

from pathlib import Path

from Printing import DocumentToPrint as dtp

pl_holidays = holidays.PL()


# def search_and_save(doc_id: str, files_dir: str, file_name: str):
#     file_ext = file_type(file_name)
#     mouse.move(x=mouse_coords[0][0], y=mouse_coords[0][1])  # przesunięcie się na lupę
#     time.sleep(1)
#     mouse.move(x=mouse_coords[1][0], y=mouse_coords[1][1])  # przesunięcie się do pola wyszukiwania
#     mouse.click()  # kliknięcie
#     pyautogui.write(doc_id)  # wpisanie numeru
#     time.sleep(1)
#     email = os.path.join(DOCS_DIR, file_name)
#
#     pyautogui.hotkey('enter')  # naciśnięcie enter do wyszukania sprawy
#     time.sleep(25)
#     if file_ext == "rtf":
#         if check_watek_dokumentu_block():
#             mouse.move(x=mouse_coords[34][0], y=mouse_coords[34][1])
#         else:
#             mouse.move(x=mouse_coords[41][0], y=mouse_coords[41][1])
#     else:
#         if check_watek_dokumentu_block():
#             mouse.move(x=mouse_coords[3][0], y=mouse_coords[3][1])  # przesunięcie się na nazwę pliku
#         else:
#             mouse.move(x=mouse_coords[42][0], y=mouse_coords[42][1])
#     mouse.click()  # kliknięcie w nazwę pliku
#     time.sleep(3)
#     if file_ext == 'pdf':
#         filename = change_filename(file_name)
#         source_file = os.path.join(DOCS_DIR, filename)
#         destination_file = os.path.join(DOCS_DIR, "prepared", filename)
#         save_or_open_file()
#         save_pdf_eml_file(files_dir, filename)
#         add_data_to_pdf("In-II.131.46.2018", "Nr dz.: " + doc_id, source_file=Path(source_file),
#                         destination_file=Path(destination_file))
#     elif file_ext == 'rtf':
#         source_file = os.path.join(DOCS_DIR, doc_id + ".pdf")
#         destination_file = os.path.join(DOCS_DIR, "prepared", doc_id + ".pdf")
#         save_or_open_file()
#         save_rtf_file(doc_id)
#         rtf_to_pdf(doc_id + ".rtf")
#         add_data_to_pdf("In-II.131.46.2018", "", source_file=Path(source_file),
#                         destination_file=Path(destination_file))
#         if not file_name.startswith("In_DOK"):
#             print_decrets()
#
#     elif file_ext == "eml":
#         source_file = os.path.join(DOCS_DIR, f"{file_name[0:-4]}.pdf")
#         destination_file = os.path.join(DOCS_DIR, "prepared", f"{file_name[0:-4]}.pdf")
#         open_file(True, DOCS_DIR, file_name)
#         eml_to_pdf(files_dir, file_name[0:-4])
#         add_data_to_pdf("In-II.131.46.2018", "Nr dz.: " + doc_id, source_file=Path(source_file),
#                         destination_file=Path(destination_file))
#         if Path(email).exists():
#             check_from_in_email(os.path.join(email))
#         print(source_file)
#
#
# def save_or_open_file():
#     mouse.move(x=mouse_coords[4][0], y=mouse_coords[4][1])
#     mouse.click()
#
#
# def open_file(save_eml=False, files_dir: str = "", filename: str = ""):
#     if save_eml:
#         save_or_open_file()
#         save_pdf_eml_file(files_dir, filename)
#         open_file()
#
#         return False
#     mouse.move(x=mouse_coords[7][0], y=mouse_coords[7][1])  # zapisanie emaila
#     time.sleep(1)
#     mouse.click()
#
#
# def save_pdf_eml_file(files_dir: str, filename: str):
#     time.sleep(2)
#     mouse.move(x=mouse_coords[5][0], y=mouse_coords[5][1])
#     mouse.click()
#     pyautogui.write(files_dir)
#     pyautogui.hotkey('enter')
#     time.sleep(2)
#     mouse.move(x=mouse_coords[17][0], y=mouse_coords[17][1])  # przesunięcie na pozycję wpisania nazwy pliku
#     mouse.click()
#     time.sleep(1)
#     pyautogui.write(filename)
#     time.sleep(1)
#     mouse.move(x=mouse_coords[6][0], y=mouse_coords[6][1])
#     mouse.click()
#
#
# def save_rtf_file(doc_id: str):
#     mouse.move(x=mouse_coords[32][0], y=mouse_coords[32][1])  # przesunięcie na pozycję wpisania nazwy pliku
#     time.sleep(1)
#     pyautogui.write(doc_id)
#     mouse.move(x=mouse_coords[33][0], y=mouse_coords[33][1])  # przesunięcie na pozycję guzika "zapisz"
#     mouse.click()
#
#

#
#
# def file_type(file_name: (str | Path)):
#     if isinstance(file_name, Path):
#         file_name = str(file_name)
#     return file_name.split(".")[-1]
#
#
# def check_from_in_email(file_name: str):
#     time.sleep(3)
#     wgn_email_adresses = [
#         'grazyna_zyber@um.poznan.pl',
#         'magda_albinska@um.poznan.pl',
#         'alicja_andrzejewska@um.poznan.pl',
#         'henryka_heigelmann@um.poznan.pl',
#         'justyna_buzuk@um.poznan.pl',
#         'wojciech_slocinski@um.poznan.pl',
#         'tomasz_borowski@um.poznan.pl',
#         'agnieszka_wierzbinska@um.poznan.pl',
#         'katarzyna_szefer@um.poznan.pl'
#     ]
#     with open(file_name, 'rb') as eml:
#         msg = BytesParser(policy=policy.default).parse(eml)
#     if msg['from'] in wgn_email_adresses:
#         print_decrets()
#
#
# def eml_to_pdf(files_dir: str, filename: str):
#     time.sleep(5)
#     mouse.move(x=mouse_coords[8][0], y=mouse_coords[8][1])  # przesunięcie na pozycję "Plik"
#     time.sleep(1)
#     mouse.click()
#     mouse.move(x=mouse_coords[9][0], y=mouse_coords[9][1])  # przesunięcie na pozycję "Drukuj"
#     time.sleep(1)
#     mouse.click()
#     time.sleep(1)
#     mouse.move(x=mouse_coords[10][0], y=mouse_coords[10][1])  # przesunięcie na pozycję wyboru listy drukarek
#     time.sleep(1)
#     mouse.click()
#     time.sleep(1)
#     mouse.move(x=mouse_coords[11][0], y=mouse_coords[11][1])  # przesunięcie na pozycję "Microsoft to PDF"
#     time.sleep(1)
#     mouse.click()
#     time.sleep(1)
#     mouse.move(x=mouse_coords[12][0], y=mouse_coords[12][1])  # przesunięcie na pozycję "Drukuj"
#     time.sleep(1)
#     mouse.click()
#     time.sleep(1)
#     mouse.move(x=mouse_coords[13][0], y=mouse_coords[13][1])  # przesunięcie na pozycję wpisania ścieżki
#     time.sleep(1)
#     mouse.click()
#     time.sleep(1)
#     pyautogui.write(files_dir)
#     time.sleep(1)
#     pyautogui.hotkey('enter')
#     time.sleep(2)
#     mouse.move(x=mouse_coords[14][0], y=mouse_coords[14][1])  # przesunięcie na pozycję wpisania nazwy pliku
#     mouse.click()
#     time.sleep(1)
#     pyautogui.write(filename)
#     time.sleep(1)
#     mouse.move(x=mouse_coords[15][0], y=mouse_coords[15][1])  # przesunięcie na pozycję zapisu pliku
#     time.sleep(1)
#     mouse.click()
#     time.sleep(1)
#     mouse.move(x=mouse_coords[16][0], y=mouse_coords[16][1])  # przesunięcie na pozycję zamknięcia outlooka
#     mouse.click()
#     return filename
#
#
# def add_data_to_pdf(case_number: str, registration_number: str, source_file: Path, destination_file: Path):
#     time.sleep(5)
#
#     packet = io.BytesIO()
#     can = canvas.Canvas(packet, pagesize=A4)
#     can.drawString(35, 820, case_number)
#     can.drawString(35, 808, registration_number)
#     can.save()
#
#     packet.seek(0)
#     new_pdf = PdfFileReader(packet)
#     existing_pdf = PdfFileReader(open(str(source_file), "rb"))
#     output = PdfFileWriter()
#     page = existing_pdf.getPage(0)
#     page.mergePage(new_pdf.getPage(0))
#     output.addPage(page)
#     for exiting_page in range(1, existing_pdf.numPages):
#         output.addPage(existing_pdf.getPage(exiting_page))
#     outputStream = open(str(destination_file), "wb")
#     output.write(outputStream)
#     outputStream.close()
#
#
# def print_decrets():
#     time.sleep(1)
#     mouse.move(x=mouse_coords[18][0], y=mouse_coords[18][1])  # przesunięcie na pozycję ikony excel w mDOK
#     time.sleep(1)
#     mouse.move(x=mouse_coords[19][0], y=mouse_coords[19][1])  # przesunięcie na pozycję raportu dekretacji
#     time.sleep(1)
#     mouse.click()
#     time.sleep(1)
#     open_file()
#     time.sleep(1)
#     mouse.click()
#     time.sleep(15)
#     pyautogui.hotkey('win', 'up')
#     mouse.move(x=mouse_coords[20][0], y=mouse_coords[20][1])  # przesunięcie na włącz edytowanie
#     mouse.click()
#     time.sleep(2)
#     mouse.move(x=mouse_coords[21][0], y=mouse_coords[21][1])  # przesunięcie na nr dz.
#     mouse.click()
#     pyautogui.hotkey('ctrl', 'c')
#     doc_id = Tk().clipboard_get()
#     new_date = parse_date(doc_id)
#     time.sleep(1)
#     mouse.move(x=mouse_coords[22][0], y=mouse_coords[22][1])  # przesunięcie na datę
#     mouse.click()
#     time.sleep(1)
#     pyautogui.write(new_date)
#     mouse.move(x=mouse_coords[23][0], y=mouse_coords[23][1])  # przesunięcie na plik
#     mouse.click()
#     time.sleep(1)
#     mouse.move(x=mouse_coords[24][0], y=mouse_coords[24][1])  # przesunięcie na drukuj
#     mouse.click()
#     time.sleep(1)
#     mouse.move(x=mouse_coords[25][0], y=mouse_coords[25][1])  # przesunięcie na wybór drukarki
#     mouse.click()
#     time.sleep(1)
#     mouse.move(x=mouse_coords[26][0], y=mouse_coords[26][1])  # wybór drukarki microsoft pdf
#     mouse.click()
#     time.sleep(1)
#     mouse.move(x=mouse_coords[27][0], y=mouse_coords[27][1])  # kliknięcie drukuj
#     mouse.click()
#     time.sleep(1)
#     mouse.move(x=mouse_coords[28][0], y=mouse_coords[28][1])  # kliknięcie pola ścieżki
#     mouse.click()
#     pyautogui.write(os.path.join(DOCS_DIR, "prepared"))
#     mouse.move(x=mouse_coords[29][0], y=mouse_coords[29][1])  # kliknięcie w pole nazwy pliku
#     mouse.click()
#     pyautogui.write(os.path.join(f"Raport dekretacji dla dokumentu {doc_id}"))
#     time.sleep(1)
#     mouse.move(x=mouse_coords[30][0], y=mouse_coords[30][1])  # zamknięcie okna
#     mouse.click()
#     time.sleep(1)
#     mouse.move(x=mouse_coords[31][0], y=mouse_coords[31][1])  # niezapisywanie zmian
#     mouse.click()
#     time.sleep(3)
#
#
# def parse_date(doc_id: str):
#     get_date_in_str = doc_id[:6]
#     date_object = datetime.datetime.strptime(get_date_in_str, "%d%m%y")
#
#     # while True:
#     #     date_object = date_object + datetime.timedelta(days=1)
#     #     if date_object.isoweekday() not in (6, 7) and (not bool(pl_holidays.get(date_object))):
#     #         break
#     return date_object.strftime("%Y.%m.%d")
#
#
# def rtf_to_pdf(rtf_file: str):
#     time.sleep(5)
#     in_file = os.path.join(DOCS_DIR, rtf_file)
#     if file_type(rtf_file) == "rtf" and Path(in_file).exists():
#         wdFormatPDF = 17
#         output_file = os.path.join(DOCS_DIR, rtf_file[:-4] + ".pdf")
#         word = comtypes.client.CreateObject('Word.application')
#         doc = word.Documents.Open(in_file)
#         doc.SaveAs(output_file, FileFormat=wdFormatPDF)
#         doc.Close()
#         word.Quit()
#
#
# def check_watek_dokumentu_block() -> bool:
#     mouse.move(x=mouse_coords[35][0], y=mouse_coords[35][1])  # kliknięcie w pole nazwy pliku
#     mouse.click()
#     pyautogui.hotkey("ctrl", "shiftleft", "i")
#     time.sleep(2)
#     mouse.move(x=mouse_coords[36][0], y=mouse_coords[36][1])
#     mouse.move(x=mouse_coords[37][0], y=mouse_coords[37][1])
#     mouse.right_click()
#     mouse.move(x=mouse_coords[38][0], y=mouse_coords[38][1])
#     time.sleep(1)
#     mouse.click()
#     mouse.move(x=mouse_coords[9][0], y=mouse_coords[39][1])
#     mouse.click()
#     time.sleep(1)
#     pyautogui.hotkey("ctrl", "a")
#     pyautogui.hotkey("ctrl", "c")
#     source_code = Tk().clipboard_get()
#     mouse.move(x=mouse_coords[40][0], y=mouse_coords[40][1])
#     mouse.click()
#     time.sleep(1)
#     pyautogui.loc
#     return bool(re.findall('Wątek\sdokumentu', source_code))
#
#
#
#
# mouse_coords = [(14, 263),  # 0
#                 (138, 320),  # 1
#                 (251, 427),  # 2
#                 (303, 370),  # 3
#                 (1667, 149),  # 4
#                 (1127, 48),  # 5
#                 (1347, 436),  # 6
#                 (1567, 149),  # 7
#                 (251, 199),  # 8
#                 (260, 406),  # 9
#                 (631, 410),  # 10
#                 (589, 545),  # 11
#                 (460, 307),  # 12
#                 (1016, 459),  # 13
#                 (826, 812),  # 14
#                 (1358, 913),  # 15
#                 (1173, 168),  # 16
#                 (345, 332),  # 17
#                 (6, 179),  # 18 - dekretacja
#                 (185, 521),  # 19
#                 (1165, 76),  # 20
#                 (257, 261),  # 21
#                 (169, 244),  # 22
#                 (24, 50),  # 23
#                 (70, 358),  # 24
#                 (402, 257),  # 25
#                 (390, 389),  # 26
#                 (245, 171),  # 27
#                 (1276, 523),  # 28
#                 (1026, 861),  # 29
#                 (1882, 19),  # 30
#                 (965, 552),  # 31
#                 (316, 330),  # 32
#                 (1390, 432),  # 33
#                 (257, 360),  # 34
#                 (636, 720),  # 35 - sprawdzenie czy jest div wątek dokumentu
#                 (229, 686),  # 36
#                 (89, 680),  # 37
#                 (190, 128),  #38
#                 (270, 724),  # 39
#                 (1893, 649),  # 40
#                 (255, 257),  # 41 - nie istnieje blok wątek dokumentu - rtf
#                 (303, 271),  # 42 - nie istnieje wątek dokumentu - pozostałe
#                 ]
#
#
# def change_filename(filename: str):
#     bad_chars = {
#         "ą": "a",
#         "ę": "e",
#         "ń": "n",
#         "ł": "l",
#         "ż": "z",
#         "ź": "z",
#         "ó": "o",
#         "ś": "s",
#         "ć": "c",
#         " ": "_",
#         "[": "",
#         "]": "",
#         ",": "",
#     }
#     replaced_filename = ""
#     for char in filename:
#         replaced_filename += bad_chars.setdefault(char, char)
#
#     return replaced_filename

def read_file(file: str):
    file_path = Path(file)
    if Path.exists(file_path) and Path(file).is_file() and file_type(file_path) == 'xls' or 'xlsx':
        FILE = str(Path(file_path))
        df = pd.read_excel(FILE, dtype='str')
    return df.to_dict()


def file_type(file_name: (str | Path)):
    if isinstance(file_name, Path):
        file_name = str(file_name)
    return file_name.split(".")[-1]

DOCS_DIR = r"C:\Users\adawis\Documents\test_spraw"

time.sleep(3)

dane = read_file(r'C:\Users\adawis\Documents\test_spraw\dane_sprawy.xls')
for key, nrdz in enumerate(dane['Numer'].values()):
    file = dane['Treść'][key]
    uid = dane['UID'][key]
    dtp(doc_id=nrdz, file_name=file, case_number='In-II.131.46.2018', uid=uid).print()

