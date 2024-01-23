import os
import re
import shutil
import zipfile


def empty_extract_folder(folder_path) -> None:
    shutil.rmtree(folder_path, ignore_errors=True)


class RemoveExcelPassword:
    def __init__(self, file_name):
        self.file_name = file_name
        self.base_name = None
        self.zip_name = None
        empty_extract_folder('extract')
        empty_extract_folder('unlocked')

    def rename_file_to_zip(self) -> None:
        self.base_name = self.file_name.split('.')[0]
        self.zip_name = self.base_name + '.zip'
        shutil.copyfile(self.file_name, self.zip_name)
        return None

    def unzip_file(self) -> None:

        with zipfile.ZipFile(self.zip_name, 'r') as zip_ref:
            zip_ref.extractall('extract')

    def get_xml_files(self, re_method=False) -> None:
        for root, dirs, files in os.walk('extract/xl/worksheets'):
            for file in files:
                if file.endswith('.xml'):
                    file_path = os.path.join(root, file)
                    if re_method:
                        self.remove_protection(file_path)
                    else:
                        self.remove_protection_re(file_path)

    def remove_protection_re(self, file_path) -> None:
        search_comp = re.compile(r"<sheetProtection (.*)scenarios=\"1\"/>")
        xml_file = open(file_path, mode='r', encoding='utf-8')
        text_original = xml_file.read()
        xml_file.close()

        text = re.sub(search_comp, "", str(text_original))
        xml_file = open(file_path, mode='w', encoding='utf-8')
        xml_file.write(text)
        xml_file.close()

    def remove_protection(self, file_path) -> None:
        # Need to find a xml parsing tool that will work
        pass

    def zip_files(self) -> None:
        shutil.make_archive(f'unlocked/{self.base_name}', 'zip', "extract")

    def rename_file_to_xlsx(self) -> None:
        unlocked_zipped = f'unlocked/{self.zip_name}'
        new_excel_name = f'unlocked/{self.base_name}.xlsx'
        shutil.copyfile(unlocked_zipped, new_excel_name)


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    passwd_remover = RemoveExcelPassword('Locked Excel Workbook.xlsx')
    passwd_remover.rename_file_to_zip()
    passwd_remover.unzip_file()
    passwd_remover.get_xml_files(re_method=False)
    passwd_remover.zip_files()
    passwd_remover.rename_file_to_xlsx()
