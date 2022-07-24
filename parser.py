import os
import xml.etree.ElementTree as ET
import pandas as pd


class XmlParser:
    namespace = {
        'o': 'urn:schemas-microsoft-com:office:office',
        'x': 'urn:schemas-microsoft-com:office:excel',
        'ss': 'urn:schemas-microsoft-com:office:spreadsheet',
        'html': 'http://www.w3.org/TR/REC-html40',
    }
    header = [
        '№ п/п',
        'Потребитель',
        'Адрес',
        'Номер ИТП',
        'Период',
    ]
    values = []

    new_head = {}
    values_dicts = []

    def __init__(self, path, number):
        self.number = number
        self.rows = self.get_rows(path)
        self.consumer = self.get_basic_info(3, 1, 'Потребитель:')
        self.address = self.get_basic_info(4, 1, 'Адрес:')
        self.itp = self.get_basic_info(7, 1, 'Тепловычислитель:', itp=True)
        self.period = self.get_basic_info(0, 6, 'Отчёт о теплопотреблении за')
        self.head_value = self.get_head_value(12, 89)

    @classmethod
    def form_new_head(cls, file):
        excel_data = pd.read_excel(file, sheet_name='Лист1')
        excel_data_dict = excel_data.to_dict()
        length = len(excel_data_dict['Name'])

        for i in range(length):
            name = excel_data_dict['Name'][i]
            value = excel_data_dict['Value'][i]

            if value in cls.new_head.keys():
                continue

            cls.new_head[value] = name

    @classmethod
    def form_values(cls):
        for value in cls.values_dicts:
            values = []
            for head in cls.header:
                if head in value.keys():
                    values.append(value[head])
                else:
                    values.append(None)
            cls.values.append(values)

    @classmethod
    def update_header(cls):
        for i in range(len(cls.header)):
            head = cls.header[i]
            if head in cls.new_head.keys():
                cls.header[i] = cls.new_head[head]

    @classmethod
    def clear_empty(cls):
        i = 0
        while i < len(cls.header):
            count = False
            for value in cls.values:
                if value[i]:
                    count = True
                    break
            if not count:
                del cls.header[i]
                for value in cls.values:
                    del value[i]
            else:
                i += 1

    @classmethod
    def save_xlsx(cls):
        df = pd.DataFrame(cls.values, columns=cls.header)
        if os.path.exists('data.xlsx'):
            os.remove('data.xlsx')
        df.to_excel('data.xlsx', index=False, encoding='utf-8-sig')

    def get_rows(self, path):
        tree = ET.parse(path)
        worksheet = tree.find('ss:Worksheet', namespaces=self.namespace)
        table = worksheet.find('ss:Table', namespaces=self.namespace)
        rows = table.findall('ss:Row', namespaces=self.namespace)

        return rows

    def get_basic_info(self, i, j, target, itp=False):
        row = self.rows[i]
        cells = row.findall('ss:Cell', namespaces=self.namespace)
        cell = cells[j]
        info = ''
        for elem in cell.iter():
            if elem.text:
                info += elem.text
        info = info.replace(target, '')
        info = info.strip()

        if itp:
            symb = info.find('№')
            number = info[symb+2:]
            info = f'ИТП {number}'

        return info

    def get_head_value(self, head_row, value_row):
        head_value = {}
        row_h = self.rows[head_row]
        cells_h = row_h.findall('ss:Cell', namespaces=self.namespace)[2:]
        row_v = self.rows[value_row]
        cells_v = row_v.findall('ss:Cell', namespaces=self.namespace)[1:]

        for i, cell_h in enumerate(cells_h):
            visible = cell_h.find('ss:NamedCell', namespaces=self.namespace)
            if visible is not None:
                head = cell_h.find('ss:Data', namespaces=self.namespace).text
                head = head.replace('.', '').upper()
                stop = head.find(',')
                head = head[:stop]

                data = cells_v[i].find('ss:Data', namespaces=self.namespace)
                if data is not None:
                    value = data.text
                    if value is not None:
                        if value != '#VALUE!':
                            value = round(float(value), 2)
                else:
                    value = None

                head_value[head] = value

        return head_value

    def append_head(self):
        for head in self.head_value.keys():
            if head not in XmlParser.header:
                XmlParser.header.append(head)

    def append_values(self):
        self.head_value[self.header[0]] = self.number
        self.head_value[self.header[1]] = self.consumer
        self.head_value[self.header[2]] = self.address
        self.head_value[self.header[3]] = self.itp
        self.head_value[self.header[4]] = self.period
        XmlParser.values_dicts.append(self.head_value)


def main():
    dir_name = '10_октябрь'
    xml_list = os.listdir(dir_name)

    for number, file in enumerate(xml_list):
        path = os.path.join(dir_name, file)
        print(path)
        xml_parser = XmlParser(path, number + 1)
        xml_parser.append_head()
        xml_parser.append_values()

    XmlParser.form_values()
    XmlParser.form_new_head('rename.xlsx')
    XmlParser.update_header()
    XmlParser.clear_empty()
    XmlParser.save_xlsx()


if __name__ == '__main__':
    main()
