import os

import jinja2
import pdfkit
import wkhtmltopdf
from pathlib import Path

config = pdfkit.configuration(wkhtmltopdf=r'C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe')


def _html2pdf(html_file, output_file):
    """
    Function that create pdf file from rendered html template.
    :return: PDF file
    """
    options = {
        'page-size': 'A4',
        'margin-top': '2cm',
        'encoding': "UTF-8",
        'margin-right': '1.5cm',
        'margin-bottom': '1.5cm',
        'margin-left': '1.5cm',
        'no-outline': None,
        'enable-local-file-access': None
    }
    with open(file=html_file) as f:
        pdfkit.from_file(f, output_file, options=options, configuration=config)


class MailProcess:
    def __init__(self, output_file_path: str, **kwargs):
        self.output_file_path = output_file_path
        self.data = dict()
        self.data.update(kwargs)

    def render_pdf(self):
        """
        Method that creating html page with email contents like Microsoft Outlook print style.
        :return: file rendered.html
        """
        template_loader = jinja2.FileSystemLoader(searchpath=os.path.join(os.getcwd(), 'html_template'))
        template_env = jinja2.Environment(loader=template_loader)
        template_file = 'template.html'
        rendered_file = 'rendered.html'
        template = template_env.get_template(template_file)
        output_data = template.render(self.data)
        with open(rendered_file, mode='w', encoding='utf8') as rendered:
            rendered.write(output_data)
        _html2pdf(rendered_file, self.output_file_path)