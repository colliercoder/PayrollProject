import datetime
from datetime import date
import holidays
from holidays import country_holidays
"""
print(holidays.Colombia(years=2022).items())

colombia = country_holidays('Colombia', years=2022)
custom_holidays = colombia
custom_holidays.update({'2022-07-16': "Dia del Minero"})
print(custom_holidays)
"""
def days_before_holidays():
    days_before_holiday = []
    for p, i in holidays.Colombia(years = 2022).items():
        day = p - datetime.timedelta(days=1)
        days_before_holiday.append(day)
    return days_before_holiday
