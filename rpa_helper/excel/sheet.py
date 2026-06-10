import os
from openpyxl import Workbook, load_workbook


class ExcelRow:
    def __init__(self, sheet, row_num, values):
        self.sheet = sheet
        self.row_num = row_num
        self.values = list(values)

    def __getitem__(self, key):
        idx = self.sheet.header_index[key]

        if idx >= len(self.values):
            return None

        return self.values[idx]

    def __setitem__(self, key, value):
        self.sheet.set_value(self.row_num, key, value)

        idx = self.sheet.header_index[key]

        while len(self.values) <= idx:
            self.values.append(None)

        self.values[idx] = value

    def __contains__(self, key):
        return key in self.sheet.header_index

    def get(self, key, default=None):
        try:
            value = self[key]
            return default if value is None else value
        except KeyError:
            return default

    def to_dict(self):
        return {
            header: self.get(header)
            for header in self.sheet.headers
        }

    def __repr__(self):
        return f"<ExcelRow row={self.row_num}>"


class ExcelSheet:
    def __init__(self, file_path):
        self.file_path = file_path

        # 文件不存在 → 创建
        if os.path.exists(file_path):
            self.wb = load_workbook(file_path)
        else:
            self.wb = Workbook()

        self.ws = self.wb.active

        # 缓存表头
        self.headers = [cell.value for cell in self.ws[1]]
        if all(v is None for v in self.headers):
            self.headers = []
            self.header_index = {}
        else:
            self.header_index = {name: idx for idx, name in enumerate(self.headers) if name}

    def has_column(self, name):
        return name in self.header_index

    def add_column(self, name):
        if self.has_column(name):
            return self.header_index[name]

        col_idx = len(self.headers)
        self.ws.cell(row=1, column=col_idx + 1, value=name)
        self.headers.append(name)
        self.header_index[name] = col_idx

        return col_idx

    def ensure_columns(self, *columns):
        for col in columns:
            self.add_column(col)

    def require_columns(self, *columns):
        missing = [
            col
            for col in columns
            if col not in self.header_index
        ]

        return missing

    def get_value(self, row, field, default=None):
        idx = self.header_index.get(field)
        if idx is None or idx >= len(row):
            return default
        return row[idx]

    def set_value(self, row_num, field, value):
        if not self.has_column(field):
            self.add_column(field)
        col_idx = self.header_index[field] + 1
        self.ws.cell(row=row_num, column=col_idx, value=value)

    def append(self, data: dict):
        for key in data.keys():
            if key not in self.header_index:
                self.add_column(key)

        row = []
        for header in self.headers:
            row.append(data.get(header))
        self.ws.append(row)

    def iter_rows(self):
        for row_num, values in enumerate(self.ws.iter_rows(min_row=2, values_only=True), start=2):
            yield row_num, values

    def rows(self):
        for row_num, values in enumerate(self.ws.iter_rows(min_row=2, values_only=True), start=2):
            yield ExcelRow(self, row_num, values)

    def save(self, file_path=None):
        self.wb.save(file_path or self.file_path)

    def close(self):
        self.wb.close()

if __name__ == "__main__":
    excel = ExcelSheet(r"C:\Users\Administrator\Desktop\test.xlsx")
    for row in excel.rows():
        print(row.to_dict())
    # excel.append({"name": "test", "value": 1})
    # excel.save()