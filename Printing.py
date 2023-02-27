import collections
import dataclasses
import configparser
import datetime
import locale
import os.path
import time
from dateutil.parser import parse
from email import policy
from tkinter import Tk
import openpyxl
import xlsxwriter
import xlrd
import xlwt
import pyautogui
import mouse
import re
import io
import comtypes.client
from collections import namedtuple
import psutil
import cv2
import fitz
import extract_msg
from bs4 import BeautifulSoup

from pathlib import Path
from email.parser import BytesParser
from PyPDF2 import PdfFileReader, PdfFileWriter, PdfReader
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from selenium import webdriver
from selenium.common import ElementNotVisibleException, ElementNotSelectableException
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from PIL import Image
from PIL.ExifTags import TAGS
from doctest import testmod
from xlutils.save import save
from MailProcess import MailProcess as mail_process
from chardet import detect

parser = configparser.ConfigParser()
parser.read(Path('config.ini'))
config = dict(parser.items(parser.sections()[0]))

acceptable_img_format = tuple(config['img_file_ext'].split(","))


def to_float(number: (str | int)) -> float:
    """
    Function that try convert `str` or `int` to `float`. Execute exception when failed.

    :param number: number in `str` or `int`
    :return: number as `float`

    >>> to_float("3")
    3.0
    >>> to_float(2)
    2.0
    >>> to_float("aaaa")
    Traceback (most recent call last):
        ...
    ValueError: could not convert string to float: 'aaaa'
    """
    try:
        if type(number) in (str, int):
            return float(number)
    except TypeError:
        print("Nie można przekonwertować podanej wartości do float")


def move_mouse_to_point(icon: str = None, x: int = 0, y: int = 0,
                        time_sleep: int = 1,
                        confidence: float = 0.0,
                        grayscale: bool = False) -> None:
    """
    Function that returns named tuple with coords of image document.png on screen. We can add to left and top values
    additional values analogical x and y to set correct coordinates for mouse move.

    :param grayscale: convert image to grayscale
    :param confidence: confidence level
    :param time_sleep: waiting time befor function run
    :param icon: path to image for recognize position
    :param x: number of additionl pixels for left as `int`
    :param y: number of additionl pixels for top as `int`
    :return: named tuple with correct coordinates

    >>> move_mouse_to_point(r'img/szukaj.png', "aaa", "bbb")
    Traceback (most recent call last):
        ...
    TypeError: x i y muszą być int

    >>> move_mouse_to_point(r'img/szukaj.png',1.0, 3.0)
    Traceback (most recent call last):
        ...
    TypeError: x i y muszą być int

    >>> move_mouse_to_point(r'img/szukaj.png',"aaa", 1)
    Traceback (most recent call last):
        ...
    TypeError: x i y muszą być int

    >>> move_mouse_to_point(r'img/szukaj.png',1, "bbb")
    Traceback (most recent call last):
        ...
    TypeError: x i y muszą być int

    """
    if icon is not None:
        if not Path(icon).exists():
            for ext in acceptable_img_format:
                if Path(icon + "." + ext).is_file():
                    icon = icon + "." + ext
                    break
                elif ext == acceptable_img_format[-1]:
                    raise NotADirectoryError("Nie znaleziono pliku")
        time.sleep(time_sleep)

        if confidence > 0:
            default_coordinates = pyautogui.locateOnScreen(icon, confidence=confidence, grayscale=grayscale)
        else:
            default_coordinates = pyautogui.locateOnScreen(icon, grayscale=grayscale)
        if default_coordinates is None:
            default_coordinates = pyautogui.locateOnScreen(icon, confidence=confidence, grayscale=grayscale)
            if default_coordinates is None:
                raise Exception('Nie znaleziono pola')
        if not isinstance(x, int) or not isinstance(y, int):
            raise TypeError("x i y muszą być int")
        mouse.move(default_coordinates.left + x, default_coordinates.top + y)
    else:
        mouse.move(x=x, y=y)


def change_filename(filename: str) -> str:
    """
    Change characters in file name to delete special characters.
    :param filename: original file name
    :return: renamed file name

    >>> change_filename('[Zażółć, gęślą jaźń]')
    'zazolc_gesla_jazn'

    >>> change_filename('[123456] Zażółć, gęślą jaźń')
    '123456_zazolc_gesla_jazn'
    """
    bad_chars = {
        "ą": "a",
        "ę": "e",
        "ń": "n",
        "ł": "l",
        "ż": "z",
        "ź": "z",
        "ó": "o",
        "ś": "s",
        "ć": "c",
        " ": "_",
        "[": "",
        "]": "",
        ",": "",
    }
    replaced_filename = ""
    for char in filename:
        if char.isalpha():
            char = char.lower()
        replaced_filename += bad_chars.setdefault(char, char)

    return replaced_filename


def is_pdf_is_bitmap(path_to_file: str) -> bool:
    """
    Function checking that pdf is scaned (as bitmap) or not. If is scanned that returns `True` otherwise `False`.
    :param path_to_file: Path to pdf file as `str`
    :return: `True` if is scanned. `False` if not.
    """
    pdf_file = fitz.open(path_to_file)
    print(pdf_file)
    for page in pdf_file:
        if len(page.get_textpage().extractText()) == 0:
            return True
    return False


def parse_date(doc_id: str) -> str:
    """
    Function that extract date from doc id. Part of the doc id number is the date of its creation.
    :param doc_id: doc id as `str'
    :return: date as 'str' in desired format.
    """
    get_date_in_str = doc_id[:6]
    date_object = datetime.datetime.strptime(get_date_in_str, "%d%m%y")

    # Functionality which is not needed at the moment
    # while True:
    #     date_object = date_object + datetime.timedelta(days=1)
    #     if date_object.isoweekday() not in (6, 7) and (not bool(pl_holidays.get(date_object))):
    #         break
    return date_object.strftime("%Y.%m.%d")


class DocumentToPrint:
    """
    Class used for print single document.

    Attributes
    ----------
    HOME_DIR: User home dir.
    DOC_OUTPUT_DIRECTORY: Directory where prepared document will be saved.
    TEMP_DOC_OUTPUT_DIRECTORY: Directory where document will be saved for processing.
    """
    _EDGE_DRIVER_PATH_DRIVER = os.path.join(os.getcwd(), "edgedriver_win64", "msedgedriver.exe")
    HOME_DIRECTORY = Path.home()
    DOC_OUTPUT_DIRECTORY = None
    TEMP_DOC_OUTPUT_DIRECTORY = None
    doc_id: str = ""
    doc_filename: str = ""
    tmp_file_path = None
    INVALID_DIR_CHARACTERS = r'[\/:*?"<>|]"]'
    _EDGE_SERVICE = Service(_EDGE_DRIVER_PATH_DRIVER)
    _OPT = Options()
    _OPT.add_experimental_option("debuggerAddress", "localhost:9222")
    driver = webdriver.Edge(service=_EDGE_SERVICE, options=_OPT)
    wgn_email_adresses = [
        'grazyna_zyber@um.poznan.pl',
        'magda_albinska@um.poznan.pl',
        'alicja_andrzejewska@um.poznan.pl',
        'henryka_heigelmann@um.poznan.pl',
        'justyna_buzuk@um.poznan.pl',
        'wojciech_slocinski@um.poznan.pl',
        'tomasz_borowski@um.poznan.pl',
        'agnieszka_wierzbinska@um.poznan.pl',
        'katarzyna_szefer@um.poznan.pl',
        'alina_arnold@um.poznan.pl',
    ]

    def __init__(self, doc_id: str, file_name: str, case_number: str, uid: (int | str)):
        """
        Initial method
        :param doc_id: Number of document.
        :param file_name: File name.
        :param case_number: Case number where document is assigned.
        :param uid: Number of document uid in `str` or `int`.
        :rtype: Any
        """
        self.doc_id = doc_id
        self.uid = uid
        self.doc_filename = change_filename(file_name)
        self.case_number = case_number
        self.DOC_OUTPUT_DIRECTORY = os.path.join(self.HOME_DIRECTORY, "documents", "sprawy",
                                                 re.sub(self.INVALID_DIR_CHARACTERS, '', self.case_number))
        self.TEMP_DOC_OUTPUT_DIRECTORY = os.path.join(self.DOC_OUTPUT_DIRECTORY, "_temp")
        if not Path(self.DOC_OUTPUT_DIRECTORY).exists():
            self.make_case_dir(self.DOC_OUTPUT_DIRECTORY, self.TEMP_DOC_OUTPUT_DIRECTORY)
        self.tmp_file_path = os.path.join(self.TEMP_DOC_OUTPUT_DIRECTORY, self.doc_filename)

    @staticmethod
    def make_case_dir(dir_path: str, tmp_dir_path) -> None:
        """
        Method that creates a dictionary with the name of the case number and creates a temporary dictionary inside it.

        :param dir_path: Path to dictionary named of case number.
        :param tmp_dir_path: Path to temporary directory.
        :return: None
        """
        try:
            os.mkdir(dir_path)
            os.mkdir(tmp_dir_path)
        except BaseException:
            raise SystemError(f"Nie udało się utworzyć katalogów {dir_path} oraz {tmp_dir_path}")

    def search_document(self) -> bool:
        """
        Method that moves the mouse cursor to search icon, next moves the mouse cursor to search input field ant type
        doc_id.  Method use pyautogui library to locate image coordinates (like search icon image, document image,
        etc.) and add values to x,y coords for forced clickable place. After that method call pyautogui.hotkey method
        to simulate pressing enter key on keyboard and waiting 20 sec for the page to load.

        :return: bool
        """
        search_icon = pyautogui.locateOnScreen(r'img/szukaj.png')
        if search_icon is None:
            search_icon = pyautogui.locateOnScreen(r'img/szukaj_2.png')
        if search_icon is None:
            raise Exception("Nie znaleziono ikony wyszukiwania.")
        mouse.move(search_icon.left, search_icon.top)
        time.sleep(1)
        search_input_field = pyautogui.locateOnScreen(r'img/wyszukiwanie.png')
        if search_input_field is None:
            raise Exception("Nie znaleziono ikony pola wyszukiwania")
        mouse.move(search_input_field.left + 100, search_input_field.top + 46)
        mouse.click()
        pyautogui.write(self.doc_id)
        pyautogui.hotkey('enter')
        if self.wait_until_load():
            return True
        else:
            raise TimeoutError("Operacja przerwana, przeglądarka za długo nie odpowiada")

    def document_type(self) -> bool:
        """
        Method checks that file is permitted to analyze.

        :return: bool
        >>> rtf = DocumentToPrint("27111902687", "sgn SEKRETARZ 2.rtf", 'test', "2382882")
        >>> rtf.document_type()
        True
        >>> pdf = DocumentToPrint("27111902687", "sgn SEKRETARZ 2.pdf", 'test', "2382882")
        >>> pdf.document_type()
        True
        >>> eml = DocumentToPrint("27111902687", "sgn SEKRETARZ 2.eml", 'test', "2382882")
        >>> eml.document_type()
        True
        >>> msg = DocumentToPrint("27111902687", "sgn SEKRETARZ 2.msg", 'test', "2382882")
        >>> eml.document_type()
        True
        >>> xls = DocumentToPrint("27111902687", "sgn SEKRETARZ 2.xls", 'test', "2382882")
        >>> xls.document_type()
        Traceback (most recent call last):
            ...
        TypeError: Niepoprawny format pliku
        """
        if self.file_type in ('pdf', 'eml', 'rtf', 'msg'):
            return True
        else:
            raise TypeError("Niepoprawny format pliku")

    @property
    def file_type(self) -> str:
        """
        Method that checks file extension.

        :return: File extension as string

        >>> obj = DocumentToPrint("27111902687", "sgn SEKRETARZ 2.msg", 'test', "2382882")
        >>> obj.file_type
        'msg'
        >>> obj = DocumentToPrint("27111902687", "sgn SEKRETARZ 2.pdf", 'test', "2382882")
        >>> obj.file_type
        'pdf'
        """
        if isinstance(self.doc_filename, Path):
            self.doc_filename = str(self.doc_filename)
        return self.doc_filename.split(".")[-1]

    def wait_until_load(self):
        """
        Method that check if element of HTML document, contains id attribute equals to DOK_{dok_id}
        has become clickable.

        :return: return True if document is clickable, False if page is not loaded after 60 sec.
        """
        try:
            ignore_error_list = [ElementNotVisibleException, ElementNotSelectableException]
            wait = WebDriverWait(driver=self.driver, timeout=to_float(config['timeout']),
                                 poll_frequency=to_float(config['freq']), ignored_exceptions=ignore_error_list)
            wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, f"[id*='DOK_{self.uid}']")))
            time.sleep(2)
            return True
        except TimeoutError:
            print("Przeglądarka nie odpowiada")
            return False

    def print_document_to_pdf(self) -> None:
        """
        A method that, based on the file extension, calls thr appropriate methods for generating a pdf file.
        :return: None
        """
        if self.document_type():
            match self.file_type():
                case "pdf":
                    self.process_pdf_file()
                case "rtf":
                    self.process_rtf_file()
                case ("msg" | "eml"):
                    self.process_messages_file()

    def process_pdf_file(self):
        move_mouse_to_point(r'img/document.png', x=90, y=105, grayscale=True, confidence=0.8)
        mouse.click()
        self.save_or_open_file()
        time.sleep(5)
        if Path(self.tmp_file_path).exists():
            self.add_data_to_pdf()
        else:
            raise FileNotFoundError("Nie znaleziono pliku pdf do dodania danych")

    def process_rtf_file(self):
        move_mouse_to_point(r'img/cert.png', x=3, y=3, grayscale=True, confidence=0.8)
        mouse.click()
        self.save_or_open_file()
        time.sleep(5)
        if Path(self.tmp_file_path).exists():
            pdf_file_name_from_rtf = self.rtf_to_pdf()
            print(pdf_file_name_from_rtf)
            self.add_data_to_pdf(filename=pdf_file_name_from_rtf, print_doc_id=False)
            if not self.doc_filename.startswith("in_doc"):
                self.print_decrets()
        else:
            raise FileNotFoundError("Nie znaleziono pliku pdf do dodania danych")

    def process_messages_file(self):
        pdf_file_name = self.doc_filename[:-3] + "pdf"
        # if self.file_type == "msg":
        #     #move_mouse_to_point(r'img/document.png', x=90, y=105, grayscale=True, confidence=0.8)
        #     x, y = self.driver.find_element(By.CSS_SELECTOR, f"a[id*='DOK_{self.uid}']").location.values()
        #     move_mouse_to_point(x=x+10, y=y+73)
        # elif self.file_type == "eml":
        #     move_mouse_to_point(r'img/document.png', x=90, y=105, grayscale=True, confidence=0.8)
        # mouse.click()
        # self.save_or_open_file()
        # time.sleep(1)
        # try:
        #     move_mouse_to_point(r'img/save_email_confirmation.png', x=90, y=55, grayscale=True, confidence=0.8)
        #     mouse.click()
        # except Exception:
        #     pass
        #time.sleep(1)
        #move_mouse_to_point(r'img/open_email.png', x=5, y=2, grayscale=True, confidence=0.8)
        #mouse.click()
        time.sleep(5)
        if self.file_type == "msg":
            message = extract_msg.openMsg(os.path.join(self.TEMP_DOC_OUTPUT_DIRECTORY, self.tmp_file_path))
            html_message = message.htmlBodyPrepared.decode('utf-8')
            print(message.headerDict['Cc'].decode('ascii'))


            # print(str(message.headerDict['Cc']).replace("=?ISO-8859-2?Q?", "")
            #       .replace("?=", "").encode("ISO-8859-2").decode("utf-8"))
            # msg_dict = {dict_key: dict_value.replace("<", '&lt;').replace(">", "&gt;")
            #             for dict_key, dict_value
            #             in message.headerDict.items()}
            # date = parse(msg_dict['Date'])
            # locale.setlocale(locale.LC_ALL, 'pl_PL')
            # msg_dict['name_and_surname'] = 'Adam Wiśniewski'
            # msg_dict['date'] = date.strftime("%A, %d %B %Y %H:%M")
            # msg_dict['subject'] = message.subject
            # msg_dict['body'] = html_message
            # print(message.headerFormatProperties)
            # if message.attachments is not None:
            #     msg_dict['attachments'] = "; ".join([attachment.name for attachment in message.attachments])
            #
            # email = mail_process(output_file_path=os.path.join(self.TEMP_DOC_OUTPUT_DIRECTORY, pdf_file_name),
            #                      **msg_dict)
            #email.render_pdf()
        #time.sleep(3)
       #self.add_data_to_pdf(filename=pdf_file_name)

    def save_or_open_file(self, save: bool = True, dir_path: str = None, filename: str = None) -> str:
        if save:
            move_mouse_to_point(r'img/save_or_open.png', confidence=0.5, x=160, y=7, time_sleep=3)
            mouse.click()
            time.sleep(1)
            self.save_file_as(dir_path=dir_path, filename=filename)

            return 'save'
        else:
            move_mouse_to_point(r'img/save_or_open.png', confidence=0.5, x=10, y=5, time_sleep=2)
            mouse.click()
            return 'open'

    def save_file_as(self, filename: str = None, dir_path: str = None, choose_pdf_format: bool = False):
        if filename is None:
            filename = self.doc_filename
        if dir_path is None:
            dir_path = self.TEMP_DOC_OUTPUT_DIRECTORY

        move_mouse_to_point(r'img/download_3.png', x=700, y=30, confidence=0.8, grayscale=True)
        mouse.click()
        pyautogui.write(dir_path)
        pyautogui.hotkey("enter")
        move_mouse_to_point(r'img/file_name.png', x=850, y=2, confidence=0.8, grayscale=True)
        mouse.click()
        pyautogui.hotkey("ctrl", "a")
        pyautogui.press("del")
        pyautogui.write(filename)
        time.sleep(1)

        if choose_pdf_format:
            time.sleep(1)
            pyautogui.press("tab", presses=2)
            pyautogui.press("p", presses=2)
        pyautogui.press("enter")

    def add_data_to_pdf(self, filename=None, print_doc_id: bool = True) -> None:
        """
        Method that adds case number and document id to first page of document. Method checks that is any text in
        document. If not that suggests scanned document which containing image itself. In this scenario image will be
        extract from PDF document and save as separated file. In next step image will be rotated if necessary and
        cropped from top. Image will be added to new blank page that already contains case number and document id. The
        last part is marge all page together to one, output PDF file.
        :param print_doc_id: If parameter is set to `True` doc id will be printed on document, otherwise not.
        :param filename: filename as `str` - different that orginal filename from object_found.
        :return: None
        """
        image_file_name = ""

        if filename is None:
            file_path = self.tmp_file_path
            destination_pdf_file = os.path.join(self.DOC_OUTPUT_DIRECTORY, self.doc_filename)
        else:
            file_path = os.path.join(self.TEMP_DOC_OUTPUT_DIRECTORY, filename)
            destination_pdf_file = os.path.join(self.DOC_OUTPUT_DIRECTORY, filename)

        existing_pdf = PdfFileReader(open(str(file_path), "rb"))
        output = PdfFileWriter()
        packet = io.BytesIO()
        page = existing_pdf.getPage(0)
        new_canvas = canvas.Canvas(packet, pagesize=A4)
        new_canvas.drawString(35, 820, "Nr sprawy: " + self.case_number)

        if print_doc_id:
            new_canvas.drawString(35, 808, "Nr dz.: " + self.doc_id)

        if is_pdf_is_bitmap(file_path):
            image_file_name = os.path.join(self.TEMP_DOC_OUTPUT_DIRECTORY, existing_pdf.pages[0].images[0].name)
            img_data = existing_pdf.pages[0].images[0].data
            with open(image_file_name, "wb", encoding='utf-8') as image_file:
                image_file.write(img_data)
            img = Image.open(image_file_name)
            width = img.width
            height = img.height

            if width > height:
                img = img.rotate(angle=90.0, expand=True)
                img.save(image_file_name)
            left = 0
            top = 150
            right = img.width
            bottom = img.height
            img = img.crop(box=(left, top, right, bottom))
            img.save(image_file_name)
            new_canvas.drawImage(image_file_name, x=0, y=0, width=595, height=787)
        new_canvas.save()
        packet.seek(0)
        new_pdf = PdfFileReader(packet)

        if is_pdf_is_bitmap(file_path):
            output.addPage(new_pdf.getPage(0))
        else:
            page.mergePage(new_pdf.getPage(0))
            output.addPage(page)

        for existing_page in range(1, existing_pdf.numPages):
            output.addPage(existing_pdf.getPage(existing_page))
        output_stream = open(str(destination_pdf_file), "wb")
        output.write(output_stream)
        output_stream.close()

        if image_file_name != "" and Path(image_file_name).exists():
            os.remove(image_file_name)

    def rtf_to_pdf(self) -> str:
        time.sleep(5)
        pdf_file_name = self.doc_id + "_" + self.doc_filename[:-4] + ".pdf"

        if self.file_type() == "rtf" and Path(self.tmp_file_path).exists():
            wd_format_pdf = 17
            output_file = os.path.join(self.TEMP_DOC_OUTPUT_DIRECTORY, pdf_file_name)
            word = comtypes.client.CreateObject('Word.application')
            doc = word.Documents.Open(self.tmp_file_path)
            doc.SaveAs(output_file, FileFormat=wd_format_pdf)
            doc.Close()
            word.Quit()
            os.remove(self.tmp_file_path)
            return pdf_file_name

    def print_decrets(self):
        ctypes = {"string": 1}
        filename = f"raport_dekretacji_{self.doc_id}.xls"
        move_mouse_to_point(r'img/decretation_img.png', x=2, y=2)
        time.sleep(1)  # waiting for the menu to appear
        move_mouse_to_point(r'img/decretation_menu_option.png', x=10, y=2)
        mouse.click()
        time.sleep(2)  # waiting to save/cancel dialog appear
        self.save_or_open_file(filename=filename)
        time.sleep(2)
        workbook = xlrd.open_workbook(filename=os.path.join(self.TEMP_DOC_OUTPUT_DIRECTORY, filename),
                                      formatting_info=True)

        dictionary_id = workbook.sheet_by_index(0).cell_value(rowx=2, colx=1)
        dictionary_id_cell_format = workbook.sheet_by_index(0).cell_xf_index(rowx=2, colx=1)
        refers_to = f"DOTYCZY:{workbook.sheet_by_index(0).cell_value(rowx=0, colx=0)}"
        refers_to_cell_format = workbook.sheet_by_index(0).cell_xf_index(rowx=3, colx=0)
        uid = f"UID: {int(workbook.sheet_by_index(0).cell_value(rowx=0, colx=2))}"
        uid_cell_format = workbook.sheet_by_index(0).cell_xf_index(rowx=2, colx=2)

        workbook.sheet_by_index(0).put_cell(rowx=1, colx=1, ctype=ctypes.get("string"), value=parse_date(dictionary_id),
                                            xf_index=dictionary_id_cell_format)
        workbook.sheet_by_index(0).put_cell(rowx=3, colx=0, ctype=ctypes.get("string"), value=refers_to,
                                            xf_index=refers_to_cell_format)
        workbook.sheet_by_index(0).put_cell(rowx=2, colx=2, ctype=ctypes.get("string"), value=uid,
                                            xf_index=uid_cell_format)
        save(workbook, os.path.join(self.DOC_OUTPUT_DIRECTORY, filename))

        # os.remove(os.path.join(self.TEMP_DOC_OUTPUT_DIRECTORY, filename))
        # time.sleep(3)  # waiting for Excel to open
        # try:
        #     move_mouse_to_point(r'img/enable_editing_excel.png', x=15, y=5, grayscale=True, confidence=0.5)
        # except Exception:
        #     time.sleep(2)
        #     move_mouse_to_point(r'img/enable_editing_excel.png', x=15, y=5, grayscale=True, confidence=0.5)
        # mouse.click()
        # time.sleep(1)  # waiting for Excel to switch to edit mode
        # move_mouse_to_point(r'img/doc_id.png', x=150, y=5, grayscale=True, confidence=0.5)  # move mouse cursor to
        # # doc_id cell in Excel
        # mouse.click()
        # pyautogui.hotkey('ctrl', 'c')
        # doc_id = Tk().clipboard_get()
        # new_date = parse_date(doc_id)
        # move_mouse_to_point(r'img/report_date.png', x=150, y=5, grayscale=True, confidence=0.5)
        # mouse.click()
        # pyautogui.write(new_date)
        # pyautogui.hotkey('enter')
        # pyautogui.hotkey('f12')
        # time.sleep(2)  # waiting to save as dialog appear
        # self.save_file_as(filename=f"raport_dekretacji_{self.doc_filename[:-4]}_{self.doc_id}.pdf",
        #                   dir_path=self.DOC_OUTPUT_DIRECTORY,
        #                   choose_pdf_format=True)
        #
        # time.sleep(4)  # wating for Adobe to open pdf
        # pyautogui.hotkey("alt", "f4")
        # time.sleep(2)
        # pyautogui.hotkey("alt", "f4")
        # pyautogui.hotkey("right")
        # pyautogui.hotkey("enter")

    def check_from_in_email(self):
        time.sleep(3)
        with open(self.tmp_file_path, 'rb') as eml:
            msg = BytesParser(policy=policy.default).parse(eml)
            if str(msg['from']).lower() in self.wgn_email_adresses:
                self.print_decrets()

    def print(self):
        if self.search_document():
            self.print_document_to_pdf()


# d = DocumentToPrint("10101800730", "07684920181008104457.pdf", 'test', "1303361")
# d = DocumentToPrint("17082102410", 'In_DOK_WEW.rtf', "test", "4176440")
# d = DocumentToPrint("24092101303", "[160920211446] Email od Dyrektora Słocińskiego w sprawie funkcjonalności"
#                                    " przekształceń.pdf", 'test', "4275558")
d = DocumentToPrint("03022301599", "Odp W sprawie obrotu wtórnego po ubruttowieniu sprawy.msg", "test",
                    "5932196")
d.process_messages_file()
# d.print_decrets()
# d.is_pdf_is_bitmap(r'C:\Users\adawis\documents\sprawy\test\_temp\email_od_dyrektora_slocinskiego_w_sprawie_funkcjonalnosci_przeksztalcen.pdf')
testmod()
