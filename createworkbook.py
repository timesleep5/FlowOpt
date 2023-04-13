from openpyxl import Workbook

from Utils.workbookutils import Workbookutils as utils


class CreateWorkbook:
    def __init__(self, wbname):
        self.wb = self._create_wb()
        self.week = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
        self.subjects = ['Gaming', 'Java', 'IT-Security', 'Coding', 'Productivity', 'Rest', 'Project']
        self.categories = ['Day', 'Date', 'Planned', 'Done', 'Resumé']
        self.wbname = wbname

    @staticmethod
    def _create_wb():
        return Workbook()

    def get_workheets(self, amount):
        for _ in range(amount):
            self.wb.create_sheet()

    def style_sheets(self):
        self.style_backlog()
        self.style_schedule()

    def rename_sheets(self):
        ws = self.wb['Sheet']
        ws.title = 'Backlog'
        ws = self.wb['Sheet1']
        ws.title = 'Schedule'
        ws = self.wb['Sheet2']
        ws.title = 'History'

    def style_backlog(self):
        ws = self.wb['Backlog']
        self.append_headline(ws, 2, 20)
        utils.style_row_border(ws, 1, len(self.week), 2, 'opentop', 'thin')

    def style_schedule(self):
        ws = self.wb['Schedule']
        self.append_headline(ws, 3, 5)
        utils.style_row_border(ws, 1, len(self.week), 3, 'full', 'thin')

    def append_headline(self, ws, startrow: int, styled_rows: int):
        ws.append(self.week)
        utils.style_row_font(ws, 1, len(self.week), 1, 'bold')
        utils.style_row_border(ws, 1, len(self.week), 1, 'full', 'thick')

        ws.append(self.subjects)
        utils.style_row_font(ws, 1, len(self.subjects), 2, 'italic')
        utils.style_row_border(ws, 1, len(self.subjects), 2, 'full', 'thin')

        utils.style_row_border_mult(ws, 1, len(self.subjects), startrow, startrow + styled_rows-1, 'sides', 'thin')
        utils.style_row_border(ws, 1, len(self.subjects), startrow+styled_rows, 'opentop', 'thin')

    def create(self):
        self.get_workheets(2)
        self.rename_sheets()
        self.style_sheets()
        self.savewb()

    def savewb(self):
        self.wb.save(filename=self.wbname)

    def get_workbook_name(self):
        return self.wbname
