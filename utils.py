import cx_Oracle
import csv
import xlrd
import os
import datetime


class ImportOracle(object):
    def inoracle(self):
        pass

    def connOracle(self):
        os.environ['NLS_LANG'] = 'SIMPLIFIED CHINESE_CHINA.utf8'
        conn = cx_Oracle.connect(self.dbUsername, self.dbPassword, self.dbIp+':1521/' + self.dbInstance)
        cursor = conn.cursor()
        fields = [i + ' varchar2(200)' for i in self.title]
        fields_str = ','.join(fields)
        try:
            sql = 'drop table %s' % (self.table_name)
            print(sql)
            cursor.execute(sql)
        except cx_Oracle.Error as e:
            print(e)

        finally:
            sql = 'create table %s (%s)' % (self.table_name, fields_str)
            print(sql)

            cursor.execute(sql)
            a = [':%s' % i for i in range(len(self.title) + 1)]
            value = ','.join(a[1:])
            sql = 'insert into %s values(%s)' % (self.table_name, value)
            print(sql)

            cursor.prepare(sql)
            cursor.executemany(None, self.data)
            cursor.close()
            conn.commit()
            conn.close()


class ImportOracleCsv(ImportOracle):
    def inoracle(self):
        with open(self.filename, 'rb') as f:
            reader = csv.reader(f)
        contents = [i for i in reader]

        title = contents[0]
        data = contents[1:]

        return (title, data)


class ImportOracleExcel(ImportOracle):
    def inoracle(self):
        wb = xlrd.open_workbook(self.filename)
        sheet1 = wb.sheet_by_index(0)
        title = sheet1.row_values(1)
        data = [sheet1.row_values(row) for row in range(2, sheet1.nrows)]
        return (title, data)


class ImportError(ImportOracle):
    def inoracle(self):
        print('Undefine file type')

        return 0


class ChooseFactory(object):
    choose = {}
    choose['csv'] = ImportOracleCsv()
    choose['xlsx'] = ImportOracleExcel()
    choose['xls'] = ImportOracleExcel()

    def choosefile(self, ch):
        if ch in self.choose:
            op = self.choose[ch]
        else:
            op = ImportError()
        return op


class ImportMain(object):

    def run(self):
        try:
            file_name = self.filePath
            table_name = 'ly_tmp_' + datetime.datetime.now().strftime('%Y%m%d')
            op = file_name.split('.')[-1]
            factory = ChooseFactory()
            cal = factory.choosefile(op)
            cal.filename = file_name
            (cal.title, cal.data) = cal.inoracle()
            cal.table_name = table_name
            cal.dbUsername = self.name;
            cal.dbPassword = self.password;
            cal.dbIp = self.ip
            cal.dbInstance = self.instance
            cal.connOracle()
            return 1
        except Exception as e:
            print(e)
            return 0