from datetime import date
from its_easy.tour.pdf import validate_booking_data


class TestValidateBookingData:
    def test_correct_data(self):
        # -- setup --
        booking_data = {
            '依頼日': date(2018, 10, 1),
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
            }
        }

        # -- exercise --
        actual = validate_booking_data(booking_data)

        # -- verify --
        assert actual == booking_data
