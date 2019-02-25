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

    @staticmethod
    def is_drawable() -> bool:
        return True


class NoPosition(DrawingPosition):
    def __init__(self):
        super().__init__(-1, -1, -1)

    @staticmethod
    def is_drawable() -> bool:
        return False


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
    },
    '利用希望コース': {
        'type': 'dict',
        'required': True,
        'schema': {
            '旅行期間': {
                'type': 'dict',
                'required': True,
                'schema': {
                    '開始日': {'type': 'date', 'required': True},
                    '終了日': {'type': 'date', 'required': True}
                }
            },
            'ツアーコード': {'type': 'string', 'required': True},
            'ツアー名': {'type': 'string', 'required': True},
            'パンフレット名／頁': {'type': 'string', 'default': ''},
            '利用ホテル': {
                'type': 'list',
                'required': True,
                'minlength': 1,
                'maxlength': 3,
                'schema': {
                    'type': 'dict',
                    'required': True,
                    'schema': {
                        '宿泊開始日': {'type': 'date', 'required': True},
                        '泊数': {'type': 'integer', 'required': True},
                        'ホテル名': {'type': 'string', 'required': True},
                        '食事': {'type': 'string', 'required': True, 'allowed': ['朝食', '朝夕食', '食事なし']},
                        'タバコ': {'type': 'string', 'default': 'どちらでも', 'allowed': ['禁煙', '喫煙', 'どちらでも']}
                    }
                }
            },
            '参加人数': {
                'type': 'dict',
                'required': True,
                'schema': {
                    '大人': {'type': 'integer', 'default': 1},
                    '子供': {'type': 'integer', 'default': 0},
                    '幼児': {'type': 'integer', 'default': 0}
                }
            },
            # 'ホテル部屋割り': {
            #     'type': 'list',
            #     'required': True,
            #     'minlength': 1,
            #     'maxlength': 3,
            #     'schema': {
            #         '人数': {'type': 'integer', 'required': True},
            #         '部屋数': {'type': 'integer', 'required': True},
            #         '1人あたり旅行代金': {
            #             '大人': {'type': 'integer'},
            #             '子供': {'type': 'integer'},
            #             '幼児': {'type': 'integer'}
            #         }
            #     }
            # }
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
    },
    '利用希望コース': {
        '旅行期間': {
            '開始日': {
                'year': DrawingPosition(87, 556, 10),
                'month': DrawingPosition(135, 556, 10),
                'day': DrawingPosition(170, 556, 10),
                'dow': DrawingPosition(204, 556, 10)
            },
            '終了日': {
                'year': DrawingPosition(242, 556, 10),
                'month': DrawingPosition(290, 556, 10),
                'day': DrawingPosition(322, 556, 10),
                'dow': DrawingPosition(355, 556, 10)
            }
        },
        'ツアーコード': DrawingPosition(445, 556, 10, 2),
        'ツアー名': DrawingPosition(87, 536, 10),
        'パンフレット名／頁': DrawingPosition(440, 536, 10),
        '利用ホテル': [
            {
                '宿泊開始日': {
                    'month': DrawingPosition(85, 516, 9),
                    'day': DrawingPosition(108, 516, 9),
                },
                '泊数': DrawingPosition(143, 516, 9),
                'ホテル名': DrawingPosition(81, 501, 9),
                '食事': {
                    '朝食': DrawingPosition(83, 481, 16),
                    '朝夕食': DrawingPosition(114, 481, 16),
                    '食事なし': DrawingPosition(154, 481, 16)
                },
                'タバコ': {
                    '禁煙': DrawingPosition(192, 481, 16),
                    '喫煙': DrawingPosition(217, 481, 16),
                    'どちらでも': NoPosition()
                }
            },
            {
                '宿泊開始日': {
                    'month': DrawingPosition(249, 516, 9),
                    'day': DrawingPosition(272, 516, 9),
                },
                '泊数': DrawingPosition(307, 516, 9),
                'ホテル名': DrawingPosition(245, 501, 9),
                '食事': {
                    '朝食': DrawingPosition(251, 481, 16),
                    '朝夕食': DrawingPosition(283, 481, 16),
                    '食事なし': DrawingPosition(318, 481, 16)
                },
                'タバコ': {
                    '禁煙': DrawingPosition(361, 481, 16),
                    '喫煙': DrawingPosition(386, 481, 16),
                    'どちらでも': NoPosition()
                }
            },
            {
                '宿泊開始日': {
                    'month': DrawingPosition(415, 516, 9),
                    'day': DrawingPosition(438, 516, 9),
                },
                '泊数': DrawingPosition(473, 516, 9),
                'ホテル名': DrawingPosition(412, 501, 9),
                '食事': {
                    '朝食': DrawingPosition(415, 481, 16),
                    '朝夕食': DrawingPosition(448, 481, 16),
                    '食事なし': DrawingPosition(488, 481, 16)
                },
                'タバコ': {
                    '禁煙': DrawingPosition(526, 481, 16),
                    '喫煙': DrawingPosition(551, 481, 16),
                    'どちらでも': NoPosition()
                }
            }
        ],
        '参加人数': {
            '大人': DrawingPosition(125, 450, 10),
            '子供': DrawingPosition(125, 436, 10),
            '幼児': DrawingPosition(125, 421, 10)
        },

    },
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

    # 利用希望コース
    ret += create_texts(booking_data, ['利用希望コース', '旅行期間', '開始日'], create_duration_text)
    ret += create_texts(booking_data, ['利用希望コース', '旅行期間', '終了日'], create_duration_text)
    ret += create_texts(booking_data, ['利用希望コース', 'ツアーコード'])
    ret += create_texts(booking_data, ['利用希望コース', 'ツアー名'])
    ret += create_texts(booking_data, ['利用希望コース', 'パンフレット名／頁'])
    for i in range(len(booking_data['利用希望コース']['利用ホテル'])):
        ret += create_texts(booking_data, ['利用希望コース', '利用ホテル', i, '宿泊開始日'], create_hotel_date_text)
        ret += create_texts(booking_data, ['利用希望コース', '利用ホテル', i, '泊数'])
        ret += create_texts(booking_data, ['利用希望コース', '利用ホテル', i, 'ホテル名'])
        ret += create_texts(booking_data, ['利用希望コース', '利用ホテル', i, '食事'], generate_selection_creator(['朝食', '朝夕食', '食事なし']))
        ret += create_texts(booking_data, ['利用希望コース', '利用ホテル', i, 'タバコ'], generate_selection_creator(['禁煙', '喫煙', 'どちらでも']))
    ret += create_texts(booking_data, ['利用希望コース', '参加人数', '大人'])
    ret += create_texts(booking_data, ['利用希望コース', '参加人数', '子供'])
    ret += create_texts(booking_data, ['利用希望コース', '参加人数', '幼児'])

    return ret


def get_deep_element(d: dict, keys: List[str]) -> Any:
    ret = d
    for key in keys:
        try:
            ret = ret[key]
        except KeyError as e:
            print("d=" + str(d))
            print("keys=" + str(keys))
            raise e
    return ret


def create_text(text: Any, position: DrawingPosition) -> List[TextOnPage]:
    t = '' if text is None else text
    return [TextOnPage(str(t), position)]


def create_booking_date_text(booking_date: date, positions: dict) -> List[TextOnPage]:
    ret = []
    heisei_year = booking_date.year - 1988  # 西暦 -> 和暦(平成)
    ret.append(TextOnPage(str(heisei_year), positions['heisei_year']))
    ret.append(TextOnPage(str(booking_date.month), positions['month']))
    ret.append(TextOnPage(str(booking_date.day), positions['day']))
    return ret


def date_to_dow(d: date) -> str:
    return ['月', '火', '水', '木', '金', '土', '日'][d.weekday()]


def create_duration_text(duration_date: date, positions: dict) -> List[TextOnPage]:
    return [
        TextOnPage(str(duration_date.year), positions['year']),
        TextOnPage(str(duration_date.month), positions['month']),
        TextOnPage(str(duration_date.day), positions['day']),
        TextOnPage(str(date_to_dow(duration_date)), positions['dow'])
    ]


def create_hotel_date_text(date_to_visit: date, positions: dict) -> List[TextOnPage]:
    return [
        TextOnPage(str(date_to_visit.month), positions['month']),
        TextOnPage(str(date_to_visit.day), positions['day']),
    ]


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


def create_texts(data: dict, target_keys: List[str], creator: Callable[[Any, Union[DrawingPosition, dict, list]], List[TextOnPage]] = create_text) -> List[TextOnPage]:
    data_element = get_deep_element(data, target_keys)
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
        if not text.position.is_drawable():
            continue
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
        },
        '利用希望コース': {
            '旅行期間': {
                '開始日': date(2019, 1, 15),
                '終了日': date(2019, 1, 18)
            },
            'ツアーコード': 'UX5600A',
            'ツアー名': 'ステイ札幌2・3・4日間',
            # 'パンフレット名／頁': 'xxx／23',
            '利用ホテル': [
                {
                    '宿泊開始日': date(2019, 1, 15),
                    '泊数': 1,
                    'ホテル名': 'ホテルAAA',
                    '食事': '食事なし',
                },
                {
                    '宿泊開始日': date(2019, 1, 16),
                    '泊数': 1,
                    'ホテル名': 'ホテルBBB',
                    '食事': '朝夕食',
                    'タバコ': '禁煙'
                },
                {
                    '宿泊開始日': date(2019, 1, 17),
                    '泊数': 1,
                    'ホテル名': 'ホテルCCC',
                    '食事': '朝食',
                    'タバコ': '喫煙'
                }
            ],
            '参加人数': {
                '大人': 3,
                '子供': 1,
                '幼児': 1
            },
            # 'ホテル部屋割り': [
            #     {
            #         '人数': 1,
            #         '部屋数': 1,
            #         '1人あたり旅行代金': {
            #             '大人': 7500
            #         }
            #     },
            #     {
            #         '人数': 2,
            #         '部屋数': 1,
            #         '1人あたり旅行代金': {
            #             '大人': 7500,
            #             '子供': 5000
            #         }
            #     },
            #     {
            #         '人数': 2,
            #         '部屋数': 1,
            #         '1人あたり旅行代金': {
            #             '大人': 7500,
            #             '幼児': 0
            #         }
            #     }
            # ]
        }
    }

    add_info_on_booking_request_paper(in_paper_file, out_paper_file, booking_data)
