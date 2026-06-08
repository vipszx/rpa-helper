from openpyxl import load_workbook


class ExcelSheet:
    def __init__(self, file_path):
        self.file_path = file_path
        self.wb = load_workbook(file_path)
        self.ws = self.wb.active

        # 缓存表头
        self.headers = [cell.value for cell in self.ws[1]]
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
        # 保存一下，确保新增列立即生效
        self.save()
        return col_idx

    def require_columns(self, *columns):
        """
        检查字段是否都存在
        返回缺失字段列表
        """
        return [
            col
            for col in columns
            if col not in self.header_index
        ]

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
        # 每写一行就保存
        self.save()

    def save(self, file_path=None):
        self.wb.save(file_path or self.file_path)

    def iter_rows(self):
        """生成器，每次返回 row_num 和值元组"""
        for row_num, row in enumerate(self.ws.iter_rows(min_row=2, values_only=True), start=2):
            yield row_num, row
