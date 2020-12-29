###########################################################
#
#   Excelファイルを管理するクラス
#
#       指定されたファイルを読み込んで、
#
###########################################################
import openpyxl
import sys
import os


def make_save_filename(file_name):
    """
    保存するファイル名を作成する
    :param file_name:
    :return:候補ファイル名 (None:諦めた)
    """
    for count in range(0, 10000):
        basename_without_ext = os.path.splitext(os.path.basename(file_name))[0]
        basename = basename_without_ext + "_%d" % count
        dir_name = os.path.dirname(file_name)
        ext_name = os.path.splitext(file_name)[1][1:]
        if ext_name != '':
            basename = basename + '.' + ext_name
        filename_candidacy = os.path.join(dir_name, basename)
        if not os.path.exists(filename_candidacy):
            return filename_candidacy
    return None


class ExcelManager:
    def __init__(self, file_name):
        """
        指定されたExcelファイルを管理する
        :param file_name:ファイル名
        """
        self.file_name = file_name
        try:
            self.work_book = openpyxl.load_workbook(file_name, read_only=True)

        except:
            print("ファイル%sがオープン出来ませんでした" % file_name)
            sys.exit(1)

        save_file_name = make_save_filename(file_name)


if __name__ == "__main__":
    name = make_save_filename("C:\\develop\\python\\pythonProject\\divider\\test.data")
    print(name)
    name = make_save_filename("test.data")
    print(name)
    name = make_save_filename("C:\\develop\\python\\pythonProject\\divider\\test")
    print(name)
    name = make_save_filename("C:\\develop\\python\\pythonProject\\divider\\test.py")
    print(name)
