import pymssql
from win32com.client import Dispatch,constants

server="."
userlogin="sa"
pwd="123456"
db_name="test"

conn = pymssql.connect(server, userlogin, pwd, db_name)

cursor = conn.cursor()

def get_tables(cursor,db_name):
    sql = "select table_name,TABLE_CATALOG from information_schema.tables where TABLE_CATALOG = '" + db_name + "'"
    cursor.execute(sql)
    result = cursor.fetchall()
    return result

def get_table_desc(cursor,db_name,table_name):
    sql = "select COLUMN_NAME,DATA_TYPE,IS_NULLABLE  from information_schema.columns where TABLE_CATALOG = '" + db_name + "' and table_name = '" + table_name + "'"
    cursor.execute(sql)
    result = cursor.fetchall()
    return result
def CreateWord():
    word = Dispatch('Word.Application')
    word.Visible = 1  # 是否在后台运行word
    word.DisplayAlerts = 0  # 是否显示警告信息
    #doc = word.Documents.Add()  # 新增一个文档
    doc = word.Documents.Open(r"d:\demo.docx")

    tables = get_tables(cursor, db_name)
    for r in tables:
        name=r[0]
        r = doc.Range(0, 0)  # 获取一个范围
        r.Style.Font.Name = u"Verdana"  # 设置字体
        r.Style.Font.Size = "9"  # 设置字体大小
        r.InsertBefore("\n " + name)  # 在这个范围前插入文本

        cols = get_table_desc(cursor, db_name,name)

        table = r.Tables.Add(doc.Range(r.End, r.End), len(cols) + 1, 4)  # 建一张表格

        table.Rows[0].Cells[0].Range.Text = u"列"
        table.Rows[0].Cells[1].Range.Text = u"类型"
        table.Rows[0].Cells[2].Range.Text = u"是否为空"
        table.Rows[0].Cells[3].Range.Text = u"列备注"

        for i in range(len(cols)):
            table.Rows[i+1].Cells[0].Range.Text = cols[i][0]
            table.Rows[i+1].Cells[1].Range.Text = cols[i][1]
            table.Rows[i+1].Cells[2].Range.Text = cols[i][2]
            table.Rows[i+1].Cells[3].Range.Text = ""



CreateWord()
#ols = get_table_desc(cursor, db_name,"user")

#list = ['html', 'js', 'css', 'python']
#for i in range(len(list)):
 #   print ("序号：%s   值：%s" % (i + 1, list[i]))







