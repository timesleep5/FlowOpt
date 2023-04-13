from datetime import timedelta


class Timeutils:
    @staticmethod
    def get_week_dates(base_date, start_day, end_day=None):
        monday = base_date - timedelta(days=base_date.isoweekday() - 1)
        week_dates = [monday + timedelta(days=i) for i in range(7)]
        week_dates = week_dates[start_day - 1:end_day or start_day]
        return [str(date.strftime('%d.%m.%y')) for date in week_dates]
