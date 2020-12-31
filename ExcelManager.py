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
import shutil


def make_save_filename(file_name):
    """
    保存するファイル名を作成する
    指定したファイル名の後ろに"_番号"をつけたファイル名を作る
    但し、既に_番号をつけたファイル名があれば、ユニークな名前になるまで番号をインクリメントする
    :param file_name:
    :return:候補ファイル名
    :note:10000までやっても重複する場合は、exit(1)で抜ける
    """
    basename_without_ext = os.path.splitext(os.path.basename(file_name))[0]  # 拡張子を除いた名前
    for count in range(0, 10000):
        basename = basename_without_ext + "_%d" % count  # "_番号"をつける
        dir_name = os.path.dirname(file_name)  # ディレクトリパス名
        ext_name = os.path.splitext(file_name)[1][1:]  # 拡張子
        if ext_name != '':  # 拡張子があれば、間に'.'をつける
            basename = basename + '.' + ext_name
        filename_candidacy = os.path.join(dir_name, basename)  # パス名を再構築
        if not os.path.exists(filename_candidacy):  # 生成したファイル名が存在しなければ、返す
            return filename_candidacy

    print("%sからファイル名を作れませんでした" % file_name)
    sys.exit(1)


class ExcelManager:
    def __init__(self, file_name):
        """
        指定されたExcelファイルを管理する
        :param file_name:ファイル名
        """
        save_file_name = make_save_filename(file_name)  # 生成するファイル名

        try:
            shutil.copy(file_name, save_file_name)  # ファイルを複写
        except FileNotFoundError as err:
            print(type(err))
            print(err)
            print("ファイル%sがありません" % file_name)
            sys.exit(1)

        except OSError as err:
            print(type(err))
            print(err)
            print("%sをファイル%sに複製出来ませんでした" % (file_name, save_file_name))
            sys.exit(1)

        self.file_name = save_file_name  # ファイル名確定
        try:
            # Excel Bookを開く
            self.work_book = openpyxl.load_workbook(self.file_name)
        except OSError as err:
            print(type(err))
            print(err)
            print("ファイル%sがオープン出来ませんでした" % self.file_name)
            sys.exit(1)

    def close(self):
        """
        ファイルを閉じる
        :return:
        """
        try:
            self.work_book.save(self.file_name)
        except OSError as err:
            print(type(err))
            print(err)
            print("ファイル%sが保存出来ませんでした" % self.file_name)
            sys.exit(1)

        self.work_book.close()

    def get_sheet(self, sheet_name):
        """
        シートオブジェクトを得る
        :param sheet_name:シート名
        :return:シートオブジェクト(無い時はNone)
        """
        if sheet_name in self.work_book.sheetnames:
            work_sheet = self.work_book[sheet_name]
            return work_sheet
        else:
            return None

    def get_column(self, sheet, column_name, column_name_line):
        """
        指定した行番号をラベル名の行とし、column_nameで指定された文字列をラベル名とするカラムを返す
        :param sheet:チェックするシート
        :param column_name:検索するラベル名
        :param column_name_line:検索する行
        :return:カラムオブジェクト(None=一致するものがない）
        """
        for cell in sheet.iter_cols(min_row=column_name_line, max_row=column_name_line):
            if cell[0].value is not None and cell[0].value == column_name:
                return cell[0]
        return None


if __name__ == "__main__":
    name = make_save_filename("C:\\develop\\python\\pythonProject\\divider\\test.data")
    print(name)
    name = make_save_filename("test.data")
    print(name)
    name = make_save_filename("C:\\develop\\python\\pythonProject\\divider\\test")
    print(name)
    name = make_save_filename("C:\\develop\\python\\pythonProject\\divider\\test.py")
    print(name)

    obj = ExcelManager("test1.xlsx")

    sheet = obj.get_sheet("test1")
    print(sheet)
    print(obj.get_sheet("test2"))

    col = obj.get_column(sheet, "品種", 2)
    print(col)
    print(obj.get_column(sheet, "テスト", 2))

    sheet = obj.get_sheet("圃場情報")
    col = obj.get_column(sheet, "品種", 2)
    print(col)
    print(obj.get_column(sheet, "テスト", 2))

    obj.close()
