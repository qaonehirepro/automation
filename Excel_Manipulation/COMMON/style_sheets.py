import xlsxwriter


class StyleSheets:

    def create_excel(self, writing_path):
        excel_sheet = xlsxwriter.Workbook(writing_path)
        self.black_color = excel_sheet.add_format()
        self.black_color_bold = excel_sheet.add_format({'font_color': 'black', 'bold': True, 'font_size': 9})
        self.red_color = excel_sheet.add_format({'font_color': 'red', 'font_size': 9})
        self.green_color = excel_sheet.add_format({'font_color': 'green', 'font_size': 9})
        self.green_color_bold = excel_sheet.add_format({'font_color': 'green', 'font_size': 9, 'bold': True})
        self.orange_color = excel_sheet.add_format({'font_color': 'orange', 'font_size': 9})
        self.over_all_status_pass = excel_sheet.add_format({'font_color': 'green', 'bold': True, 'font_size': 9})
        self.over_all_status_failed = excel_sheet.add_format({'font_color': 'red', 'bold': True, 'font_size': 9})
        self.over_all_status_color = excel_sheet.add_format({'font_color': 'green', 'bold': True, 'font_size': 9})
        self.over_all_status = 'Pass'
        return excel_sheet


style_sheet_obj = StyleSheets()
