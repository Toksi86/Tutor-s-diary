import locale
import datetime as dt

locale.setlocale(locale.LC_ALL, 'ru_RU.UTF-8')


class Date(dt.datetime):
    """Надстройка над классом datetime"""

    def __new__(cls, year=None, month=None, day=None, hour=0, minute=0, second=0,
                microsecond=0, tzinfo=None, *, fold=0, date_time=None):
        if date_time is not None:
            return super().__new__(cls, date_time.year, date_time.month, date_time.day)
        # Почему убрали значения по умолчанию?
        return super().__new__(cls, year, month, day, hour, minute, second,
                               microsecond, tzinfo, fold=fold)

    def _next_day(self):
        """Получить следующий день от текущего"""
        return self + dt.timedelta(days=1)

    def get_weekday(self):
        """Получить день недели от текущей даты"""
        return self.strftime("%A").capitalize()

    def get_datetime_with_format(self):
        """Получить текущую дату в формате день-месяц-год"""
        return self.strftime('%d-%m-%Y')

    def next_workday(self):
        """Получить следующий рабочий день от текущей даты"""
        if self.weekday() == 4:
            return self + dt.timedelta(days=3)
        if self.weekday() == 5:
            return self + dt.timedelta(days=2)
        return self._next_day()
