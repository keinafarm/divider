##########################################################
#
#   指定したファイルの指定したシートの指定したカラムの内容と同じ名前でシートを作って
#   そこに、コピーする
#
#   Usage:
#       > divider  file名 シート名 カラム名 [タイトル行数]
#           file名 : 抽出するExcelファイル名
#           シート名 : 抽出するデータが入っているシート
#           カラム名 : 抽出するキーワードとなるカラム名
#           タイトル行数 : タイトルとして各シートにコピーする行数（省略時は2)
#
#   カラム名で指定されたカラムのタイトル行数以降のセルに入っている文字列を抽出し
#   ソートしてから、重複を省く
#   上記の処理で出来たリストを、シート名として新たにシートを作成する
#   各シートには、それぞれカラム名で指定されたカラムの内容が、シート名と一致する行を
#   シート名で指定されたシートからコピーする
#   但し、各シートの上から「タイトル行数」分には、シート名で指定されたシートから無条件にコピーされる
#
##########################################################

from ExcelManager import ExcelManager
import sys


class Divider:
    def __init__(self, file_name, sheet_name, column_name, title_lines):
        """
        指定したシートの指定したカラムの内容に従って、シートを作って分配する
        :param file_name:ファイル名
        :param sheet_name:抽出元のシート名
        :param column_name:分配するキーにするカラム名
        :param title_lines:カラム名が入っている行の番号
        """
        self.work_book = ExcelManager(file_name)  # Excelファイルを管理するオブジェクト
        self.sheet_name = sheet_name  # 抽出元のシート名
        self.column_name = column_name  # 分配するキーにするカラム名
        self.title_lines = title_lines  # カラム名が入っている行の番号

        source_sheet = self.work_book.get_sheet(self.sheet_name)  # ファイルを読み込んで管理するオブジェクトを作る
        if source_sheet is None:  # 失敗したらエラーを出して終了
            print("元になるシート%sがありません" % sheet_name)
            sys.exit(1)

        # 抽出するシートから指定されたカラムを特定する
        search_column = self.work_book.get_column(source_sheet, column_name, title_lines)
        if search_column is None:  # 特定出来なかったらエラーを出して終了
            print("元になるカラム%sがありません" % column_name)
            sys.exit(1)

        # 特定出来たカラムからデータを取り出す
        columns = self.work_book.get_column_data(search_column, title_lines + 1)
        if columns is None:  # 取り出せなかったらエラーを出して終了
            print("データがありませんでした %s" % column_name)
            sys.exit(1)

        columns = set(columns)  # 取り出したカラムのデータの重複を取り除く
        #        print(columns)

        # 取り出したカラムのデータの文字列を名前にしてシートを作成する
        self.sheets = self.work_book.make_sheet(columns)
        # 抽出元のシートから、コピー先で共通に使用される行を取り出す
        common_rows = self.work_book.get_rows_by_lineNo(source_sheet, 1, title_lines)
        #        print(common_rows)

        # 抽出元のシートから行単位でデータを抽出して、作成したシートにコピーする
        for keyword in self.sheets.keys():
            rows = self.work_book.get_rows_by_searched_column(search_column, keyword, title_lines + 1)
            #            print(self.sheets[keyword])
            #            print(rows)
            self.work_book.append_rows(self.sheets[keyword], common_rows)
            self.work_book.append_rows(self.sheets[keyword], rows)

        self.work_book.close()


if __name__ == "__main__":
    args = sys.argv
    parameters = len(args)
    if parameters < 4 or parameters > 5:
        print("Usage:")
        print("   divider  file名 シート名 カラム名 [タイトル行数]")
        sys.exit(1)

    if parameters == 4:
        lines = 2
    else:
        lines = int(args[4])

    obj = Divider(args[1], args[2], args[3], lines)
