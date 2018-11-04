import io
import os
from datetime import date
from typing import Union, List, Callable, Any, Iterable
from PyPDF2 import PdfFileReader, PdfFileWriter
from PyPDF2.pdf import PageObject
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from cerberus import Validator
import yaml


PDF_FONT = 'HeiseiMin-W3'


class DrawingPosition:
    def __init__(self, x: int, y: int, font_size: int = 10, char_space: int = 0):
        self.x = x
        self.y = y
        self.font_size = font_size
        self.char_space = char_space

    def __str__(self):
        return 'DrawingPosition(x={}, y={}, font_size={}, char_space={})'\
            .format(self.x, self.y, self.font_size, self.char_space)


class TextOnPage:
    def __init__(self, text: str, position: DrawingPosition):
        self.text = text
        self.position = position

    def __str__(self):
        return 'TextOnPage(text={}, position={})'.format(self.text, self.position)


BOOKING_DATA_SCHEMA = {
    '依頼日': {'type': 'date', 'default': date.today()},
    '利用代表者': {
        'type': 'dict',
        'required': True,
        'schema': {
            '利用代表者名': {'type': 'string', 'required': True},
            'フリガナ': {'type': 'string', 'required': True},
            '性別': {'type': 'string', 'required': True, 'allowed': ['男', '女']},
            '勤務先名': {'type': 'string', 'required': True},
            '代表利用者の方の保険証': {
                'type': 'dict',
                'required': True,
                'schema': {
                    '記号': {'type': 'integer', 'required': True, 'min': 0},
                    '番号': {'type': 'integer', 'required': True, 'min': 0}
                }
            },
            '連絡先電話番号': {
                'type': 'list',
                'required': True,
                'minlength': 2,
                'maxlength': 2,
                'schema': {
                    'type': 'dict',
                    'required': True,
                    'schema': {
                        '番号': {'type': 'string', 'required': True, 'regex': r'^\d+-\d+-\d+$', 'maxlength': 13},
                        '種別': {'type': 'string', 'required': True, 'allowed': ['携帯', '自宅', '勤務先']}
                    }
                }
            }

        }
    },
    '請求書・旅行書類送付先': {
        'type': 'dict',
        'required': True,
        'schema': {
            '請求書送付先住所': {
                'type': 'dict',
                'required': True,
                'schema': {
                    '種別': {'type': 'string', 'required': True, 'allowed': ['自宅', '勤務先']},
                    '郵便番号': {'type': 'string', 'required': True, 'regex': r'^\d{3}-\d{4}$'},
                    '住所': {'type': 'string', 'required': True}
                }
            },
            '旅行書類等送付先住所': {
                'type': 'dict',
                'required': True,
                'schema': {
                    '種別': {'type': 'string', 'required': True, 'allowed': ['自宅', '勤務先']},
                    '郵便番号': {'type': 'string', 'required': True, 'regex': r'^\d{3}-\d{4}$'},
                    '住所': {'type': 'string', 'required': True}
                }
            }
        }
    }
}


BOOKING_DATA_POSITIONS = {
    '依頼日': {
        'heisei_year': DrawingPosition(62, 747, 9),
        'month': DrawingPosition(84, 747, 9),
        'day': DrawingPosition(104, 747, 9)
    },
    '利用代表者': {
        '利用代表者名': DrawingPosition(180, 712, 10),
        'フリガナ': DrawingPosition(180, 732, 10),
        '性別': {
            '男': DrawingPosition(501, 711, 16),
            '女': DrawingPosition(535, 711, 16),
        },
        '勤務先名': DrawingPosition(180, 690, 9),
        '代表利用者の方の保険証': {
            '記号': DrawingPosition(490, 690, 10, 2),
            '番号': DrawingPosition(540, 690, 10, 2)
        },
        '連絡先電話番号': [
            {
                '番号': [DrawingPosition(225, 668, 10, 2), DrawingPosition(290, 668, 10, 2), DrawingPosition(360, 668, 10, 2)],  # '-' 区切り
                '種別': {
                    '携帯': DrawingPosition(441, 666, 16),
                    '自宅': DrawingPosition(465, 666, 16),
                    '勤務先': DrawingPosition(493, 666, 16)
                }
            },
            {
                '番号': [DrawingPosition(225, 647, 10, 2), DrawingPosition(290, 647, 10, 2), DrawingPosition(360, 647, 10, 2)],
                '種別': {
                    '携帯': DrawingPosition(441, 645, 16),
                    '自宅': DrawingPosition(465, 645, 16),
                    '勤務先': DrawingPosition(493, 645, 16)
                }
            }
        ]
    },
    '請求書・旅行書類送付先': {
        '請求書送付先住所': {
            '種別': {
                '自宅': DrawingPosition(144, 606, 16),
                '勤務先': DrawingPosition(177, 606, 16)
            },
            '郵便番号': DrawingPosition(217, 609, 8, 2),
            '住所': DrawingPosition(266, 609, 8)
        },
        '旅行書類等送付先住所': {
            '種別': {
                '自宅': DrawingPosition(144, 585, 16),
                '勤務先': DrawingPosition(177, 585, 16)
            },
            '郵便番号': DrawingPosition(217, 589, 8, 2),
            '住所': DrawingPosition(266, 589, 8)
        }
    }
}


def add_info_on_booking_request_paper(in_paper_file: str,
                                      out_paper_file: str,
                                      booking_data: Union[dict, str],
                                      form_page_num: int = 2) -> None:
    if isinstance(booking_data, str):
        booking_data_dict = parse_booking_data(booking_data)
    else:
        booking_data_dict = booking_data
    booking_data_dict = validate_booking_data(booking_data_dict)
    texts = booking_data_dict_to_texts(booking_data_dict)
    edit_booking_request_paper(in_paper_file, out_paper_file, texts, form_page_num)


def parse_booking_data(file_or_doc: str) -> dict:
    # to dict
    if os.path.isfile(file_or_doc):
        with open(file_or_doc, 'r') as f:
            dict_booking_data = yaml.load(f)
    else:
        dict_booking_data = yaml.load(file_or_doc)
    return dict_booking_data


def validate_booking_data(booking_data: dict) -> dict:
    validator = Validator(BOOKING_DATA_SCHEMA)
    validated_data = validator.validated(booking_data)
    if validated_data is None:
        raise RuntimeError('Validation for booking data was failed. : ' + str(validator.errors))
    return validated_data


def booking_data_dict_to_texts(booking_data: dict) -> List[TextOnPage]:
    ret = []

    # 依頼日
    ret += create_texts(booking_data, ['依頼日'], create_booking_date_text)

    # 利用代表者
    ret += create_texts(booking_data, ['利用代表者', '利用代表者名'])
    ret += create_texts(booking_data, ['利用代表者', 'フリガナ'])
    ret += create_texts(booking_data, ['利用代表者', '性別'], generate_selection_creator(['男', '女']))
    ret += create_texts(booking_data, ['利用代表者', '勤務先名'])
    ret += create_texts(booking_data, ['利用代表者', '代表利用者の方の保険証', '記号'])
    ret += create_texts(booking_data, ['利用代表者', '代表利用者の方の保険証', '番号'])
    ret += create_texts(booking_data, ['利用代表者', '連絡先電話番号', 0, '番号'], create_phone_number_text)
    ret += create_texts(booking_data, ['利用代表者', '連絡先電話番号', 0, '種別'], generate_selection_creator(['携帯', '自宅', '勤務先']))
    ret += create_texts(booking_data, ['利用代表者', '連絡先電話番号', 1, '番号'], create_phone_number_text)
    ret += create_texts(booking_data, ['利用代表者', '連絡先電話番号', 1, '種別'], generate_selection_creator(['携帯', '自宅', '勤務先']))

    # 請求書・旅行書類送付先
    ret += create_texts(booking_data, ['請求書・旅行書類送付先', '請求書送付先住所', '種別'], generate_selection_creator(['自宅', '勤務先']))
    ret += create_texts(booking_data, ['請求書・旅行書類送付先', '請求書送付先住所', '郵便番号'])
    ret += create_texts(booking_data, ['請求書・旅行書類送付先', '請求書送付先住所', '住所'])
    ret += create_texts(booking_data, ['請求書・旅行書類送付先', '旅行書類等送付先住所', '種別'], generate_selection_creator(['自宅', '勤務先']))
    ret += create_texts(booking_data, ['請求書・旅行書類送付先', '旅行書類等送付先住所', '郵便番号'])
    ret += create_texts(booking_data, ['請求書・旅行書類送付先', '旅行書類等送付先住所', '住所'])


    return ret


def get_deep_element(d: dict, keys: List[str]) -> Any:
    ret = d
    for key in keys:
        ret = ret[key]
    return ret


def create_text(text: Any, position: DrawingPosition) -> List[TextOnPage]:
    return [TextOnPage(str(text), position)]


def create_booking_date_text(booking_date: date, positions: dict) -> List[TextOnPage]:
    ret = []
    heisei_year = booking_date.year - 1988  # 西暦 -> 和暦(平成)
    ret.append(TextOnPage(str(heisei_year), positions['heisei_year']))
    ret.append(TextOnPage(str(booking_date.month), positions['month']))
    ret.append(TextOnPage(str(booking_date.day), positions['day']))
    return ret


def generate_selection_creator(selectable_items: list) -> Callable[[str, dict], List[TextOnPage]]:
    def ret(text: str, positions: dict):
        for item in selectable_items:
            if text == item:
                return [TextOnPage('○', positions[item])]
        raise RuntimeError('Selectable items are {}, but got "{}"'.format(str(selectable_items), text))
    return ret


def create_phone_number_text(text: str, positions: list) -> List[TextOnPage]:
    """電話番号を '-' 区切りにする。"""
    phone_number_parts = text.split('-')
    assert len(phone_number_parts) == 3
    return [TextOnPage(phone_number_parts[i], positions[i]) for i in range(0, 3)]


def create_texts(booking_data: dict, target_keys: List[str], creator: Callable[[Any, Union[DrawingPosition, dict, list]], List[TextOnPage]] = create_text) -> List[TextOnPage]:
    data_element = get_deep_element(booking_data, target_keys)
    drawing_element = get_deep_element(BOOKING_DATA_POSITIONS, target_keys)
    return creator(data_element, drawing_element)


def edit_booking_request_paper(in_paper_file: str, out_paper_file: str, texts: List[TextOnPage], form_page_num: int = 2) -> None:
    # 読み込み
    with open(in_paper_file, 'rb') as f_in:
        in_pdf = PdfFileReader(f_in)

        # フォームになっているページのみを抜き出す
        form_page = in_pdf.getPage(form_page_num)

        # テキストを付与
        edited_page = add_text_on_page(form_page, texts)

        # 書き出し
        out_pdf = PdfFileWriter()
        out_pdf.addPage(edited_page)

        with open(out_paper_file, 'wb') as f_out:
            out_pdf.write(f_out)


def add_text_on_page(pdf_page: PageObject, texts: Iterable[TextOnPage]) -> PageObject:
    buf = io.BytesIO()

    # create a new PDF with Reportlab
    pdfmetrics.registerFont(UnicodeCIDFont(PDF_FONT))   # 日本語表示のためにフォントを登録する必要がある
    can = canvas.Canvas(buf, pagesize=A4)

    # 与えられたテキスト情報を canvas に追加していく
    for text in texts:
        text_obj = can.beginText(text.position.x, text.position.y)
        text_obj.setFont(PDF_FONT, text.position.font_size)
        text_obj.setCharSpace(text.position.char_space)
        text_obj.textLine(text.text)    # textLine() する前に他の setXxx() を終わらせる必要がある
        can.drawText(text_obj)

    can.save()

    buf.seek(0)
    temp_pdf = PdfFileReader(buf)
    pdf_page.mergePage(temp_pdf.getPage(0))

    return pdf_page


# 実験用
if __name__ == '__main__':
    in_paper_file = './data/201106_tehaisho.pdf'
    out_paper_file = './data/out.pdf'

    booking_data = {
        '利用代表者': {
            '利用代表者名': '健保 太郎',
            'フリガナ': 'ケンポ タロウ',
            '性別': '男',
            '勤務先名': '株式会社○△□',
            '代表利用者の方の保険証': {
                '記号': 1234,
                '番号': 56
            },
            '連絡先電話番号': [
                {'番号': '090-1234-5678', '種別': '携帯'},
                {'番号': '0123-45-6789', '種別': '自宅'}
            ]
        },
        '請求書・旅行書類送付先': {
            '請求書送付先住所': {
                '種別': '自宅',
                '郵便番号': '123-4567',
                '住所': '東京都港区○△□1-2-3'
            },
            '旅行書類等送付先住所': {
                '種別': '勤務先',
                '郵便番号': '345-6789',
                '住所': '東京都港区○△□4-5-6'
            }
        }
    }

    add_info_on_booking_request_paper(in_paper_file, out_paper_file, booking_data)
