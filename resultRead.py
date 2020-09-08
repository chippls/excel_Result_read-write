import xlrd
import xlsxwriter
import xlwt
import openpyxl


def result(n):
    workbook = openpyxl.open("D:/工程热物理大会/换热器/套管换热器/20200907思科换热器Matlab结果/换热器2不同管长.xlsx")
    table = workbook.get_sheet_by_name(n)
    # 读取结果并保存为数组
    t_h_out = table["A41"].value
    t_c_out = table["B1"].value
    temp = [t_h_out, t_c_out]
    return temp


def main():
    # xlsxwriter创建表单并写入
    workbook = xlsxwriter.Workbook("测试结果.xlsx")
    table = workbook.add_worksheet("sheet1")
    for i in range(10):
        # python3.7不支持bytes与string拼接，转化为string
        sheet = "Sheet"+str(i+1)
        num1 = "A"+str(i+1)
        num2 = "B"+str(i+1)
        t_h_out = float("{:.2f}".format(result(sheet)[0]))
        t_c_out = float("{:.2f}".format(result(sheet)[1]))
        # print("{:.2f}".format(result(i)[0])+'\t'+"{:.2f}".format(result(i)[1]))
        # 不识别制表符'\t'，只能指定列
        table.write_number(num1, t_h_out)
        table.write_number(num2, t_c_out)
    workbook.close()
    print("已完成")


main()