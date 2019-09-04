from io import BytesIO
from typing import BinaryIO

from PyPDF4.pdf import PageObject

from reportlab.pdfgen import canvas
from PyPDF4 import PdfFileWriter, PdfFileReader


class Color:
    """A font color used in text drawing.

        Accepts color values as standard RGB values in the range from `0` to `255`, but internally stores
        the values as float values in the range from `0.0` to `1.0` respectively.
    """

    def __init__(self, r: int = 0, g: int = 0, b: int = 0) -> None:
        self.r = self._convert_from_255_format(r)
        self.g = self._convert_from_255_format(g)
        self.b = self._convert_from_255_format(b)

    @staticmethod
    def _convert_from_255_format(value):
        return value / 255

    def as_tuple(self):
        return self.r, self.g, self.b


class Location:
    """A location that denotes the actual place on a particular page of the PDF document where the data should be inserted."""

    def __init__(self, x: int, y: int, page: int = 0) -> None:
        self.x = x
        self.y = 841.9 - y
        self.page = page


class PdfDataConfig:
    """A configuration for the PDF data to be merged.

       font_size may be negative. In this case, the text will be upside down.
    """

    defaults_settings = dict(font="Helvetica", font_size=12, color=Color(0, 0, 0))

    def __init__(
        self,
        font: str = None,
        font_size: int = None,
        color: Color = None,
        location: Location = None,
        text: str = None,
    ) -> None:
        self.font = font
        self.font_size = font_size
        self.color = color
        self.location = location
        self.text = text


class PdfData:
    """PDF data to be merged."""

    def __init__(
        self,
        general_setting: PdfDataConfig = PdfDataConfig(),
        configs: [PdfDataConfig] = None,
    ) -> None:
        self.general_setting = general_setting
        self.configs = configs


class PdfCreationFailed(Exception):
    pass


def _create_pdf_page_with_data(pdf_data: PdfDataConfig) -> PageObject:
    bytes_data = BytesIO()
    can = canvas.Canvas(bytes_data)

    try:
        can.setFont(pdf_data.font, pdf_data.font_size)
    except KeyError:
        raise PdfCreationFailed(f"Invalid font: {pdf_data.font}")

    can.setFillColorRGB(*pdf_data.color.as_tuple())
    can.drawString(pdf_data.location.x, pdf_data.location.y, text=pdf_data.text)
    can.save()

    bytes_data.seek(0)
    return PdfFileReader(bytes_data).getPage(0)


def _get_pdf_page(file: BinaryIO, page: int) -> PageObject:
    existing_pdf = PdfFileReader(file)
    return existing_pdf.getPage(page)


def _write_pdf_to_file(filename: str, pdf_writer: PdfFileWriter) -> None:
    with open(filename, "wb") as result:
        pdf_writer.write(result)


def _insert_page_to_pdf(
    original: BinaryIO, page: PageObject, page_index: int
) -> PdfFileWriter:
    input = PdfFileReader(original)
    output = PdfFileWriter()
    for i in range(input.getNumPages()):
        if i != page_index:
            p = input.getPage(i)
            output.addPage(p)
        else:
            output.addPage(page)
    return output


def _merge_data(pdf_data: PdfDataConfig, original: BinaryIO) -> BytesIO:
    new_page_with_text = _create_pdf_page_with_data(pdf_data)
    page = pdf_data.location.page
    try:
        source_page = _get_pdf_page(original, page)
    except IndexError:
        raise PdfCreationFailed(
            f"No page with current index: {page}. "
            f"Please note that page numbers start at 0"
        )

    source_page.mergePage(new_page_with_text)
    result_pdf_writer = _insert_page_to_pdf(original, source_page, page)
    output = BytesIO()
    result_pdf_writer.write(output)
    return output


def _combine_configs(
    general_settings: PdfDataConfig, config: PdfDataConfig = None
) -> PdfDataConfig:
    """Combines PDF configs.

      Priority for combining from high to low: single configuration(config) -> general settings -> default values.
    """

    processed_pdf_data = PdfDataConfig()
    for field_name in vars(processed_pdf_data):
        value = (
            getattr(config, field_name, None)
            or getattr(general_settings, field_name)
            or PdfDataConfig.defaults_settings.get(field_name)
        )
        if not value:
            raise PdfCreationFailed(
                f"Can't get {field_name} from configs: "
                f"general_settings={vars(general_settings)}; config="
                f"{vars(config) if config else None}"
            )
        setattr(processed_pdf_data, field_name, value)
    return processed_pdf_data


def _prepare_configs(pdf_data_list: [PdfData]) -> [PdfDataConfig]:
    """Creates a list of PdfDataConfig with a set of all the necessary data to merge
       with the original pdf from the list of PdfData.
    """
    prepared_data_list = []
    for pdf_data in pdf_data_list:
        if pdf_data.configs:
            for config in pdf_data.configs:
                prepared_data = _combine_configs(
                    pdf_data.general_setting, config=config
                )
                prepared_data_list.append(prepared_data)
        else:
            prepared_data_list.append(_combine_configs(pdf_data.general_setting))
    return prepared_data_list


def merge_pdf_with_data(original: BinaryIO, pdf_data_list: [PdfData]) -> BytesIO:
    """Merges PDF data with original PDF

    Supported fonts = ('Courier', 'Courier-Bold', 'Courier-Oblique', 'Courier-BoldOblique',
    'Helvetica', 'Helvetica-Bold', 'Helvetica-Oblique', 'Helvetica-BoldOblique',
    'Times-Roman', 'Times-Bold', 'Times-Italic', 'Times-BoldItalic',
    'Symbol','ZapfDingbats')

    Parameters:
        original (BinaryIO): original source PDF as a byte array
        pdf_data_list ([PdfData]): List of data to merge with the original PDF

    Returns:
        (BytesIO): merged PDF file as a byte array
    """
    result = original
    prepared_pdf_data = _prepare_configs(pdf_data_list)
    for pdf_data in prepared_pdf_data:
        result = _merge_data(pdf_data, result)
    return result
