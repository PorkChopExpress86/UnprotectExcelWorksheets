import os
import re
import shutil
import zipfile


def empty_extract_folder() -> None:
    shutil.rmtree('extract')
    with open('extract/.gitkeep', 'w') as f:
        pass


class RemoveExcelPassword:
    def __init__(self, file_name):
        self.file_name = file_name
        self.base_name = None
        self.zip_name = None

    def rename_file(self) -> None:
        self.base_name = self.file_name.split('.')[0]
        self.zip_name = self.base_name + '.zip'
        shutil.copyfile(self.file_name, self.zip_name)
        return None

    def unzip_file(self) -> None:
        empty_extract_folder()
        with zipfile.ZipFile(self.zip_name, 'r') as zip_ref:
            zip_ref.extractall('extract')

    def get_xml_files(self) -> None:
        for root, dirs, files in os.walk('extract/xl/worksheets'):
            for file in files:
                if file.endswith('.xml'):
                    file_path = os.path.join(root, file)
                    self.remove_protection(file_path)

    def remove_protection(self, file_path) -> None:
        search_comp = re.compile(r"<sheetProtection (.*)\/>")
        xml_file = open(file_path, 'r')
        text = xml_file.readlines()
        xml_file.close()

        text = re.sub(search_comp, "", str(text))
        xml_file = open(file_path, 'w')
        xml_file.write(text)
        xml_file.close()


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    passwd_remover = RemoveExcelPassword('Locked Excel Workbook.xlsx')
    passwd_remover.rename_file()
    passwd_remover.unzip_file()
    passwd_remover.get_xml_files()
