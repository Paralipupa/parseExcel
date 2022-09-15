from file_readers import get_file_reader
import csv, json


class ExcelImporter:

    def __init__(self, file_name: str, page_name:str=None ):
        self.name = file_name
        self.page_name = page_name
        self.records = list()

    def _get_data_xls(self):
        ReaderClass = get_file_reader(self.name)
        data_reader = ReaderClass(self.name, self.page_name)
        if not data_reader:
            raise Exception(f'file reading error: {self.name}')
        return data_reader

    def _get_names(self, record: list) -> list:
        names = []
        for cell in record:
            if cell:
                nm = dict()
                nm['name'] = str(cell).strip()
                names.append(nm)
        return names

    def _get_record(self, names:list, record:list) -> dict:
        rec = dict()
        index = 0
        for cell in record:
            if index < len(names):
                rec[names[index]['name']] = str(cell).strip()
            index += 1
        return rec


    def read(self) -> bool:
        data_reader = self._get_data_xls()
        index = 0
        names = []
        for record in data_reader:
            if index == 0:
                names = self._get_names(record)
            else:
                self.records.append(self._get_record(names,record))
            index += 1
        return True

    def write(self, file_output) -> bool:
        with open(f'{file_output}.json', mode='w', encoding='utf-8') as file:
            jstr = json.dumps(self.records, indent=4,
                                ensure_ascii=False)
            file.write(jstr)

        with open(f'{file_output}.csv', mode='w', encoding='utf-8') as file:
            names = [x for x in self.records[0]]
            file_writer = csv.DictWriter(file, delimiter=";",
                                            lineterminator="\r", fieldnames=names)
            file_writer.writeheader()
            for rec in self.records:
                file_writer.writerow(rec)

        return True


if __name__ == '__main__':
    parser = ExcelImporter('input.xls')
    parser.read()
    parser.write('output')
