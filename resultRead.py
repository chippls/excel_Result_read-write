import xlrd
import xlsxwriter
import xlwt
import openpyxl


def result(n):
    # 读取文件路径
    wb = xlrd.open_workbook("D:/工程热物理大会/换热器/套管换热器/20200816思科换热器测试/换热器2不同管长.xlsx")
    table = wb.sheet_by_index(n)
    # 读取结果并保存为数组
    t_h_out = table.cell(40, 0).value
    t_c_out = table.cell(0, 1).value
    temp = [t_h_out, t_c_out]
    return temp


def main():
    # xlsxwriter创建表单并写入
    workbook = xlsxwriter.Workbook("测试结果.xlsx")
    table = workbook.add_worksheet("sheet1")
    for i in range(5):
        # python3.7不支持bytes与string拼接，转化为string
        num1 = "A"+str(i)
        num2 = "B"+str(i)
        t_h_out = float("{:.2f}".format(result(i)[0]))
        t_c_out = float("{:.2f}".format(result(i)[1]))
        # print("{:.2f}".format(result(i)[0])+'\t'+"{:.2f}".format(result(i)[1]))
        # 不识别制表符'\t'，只能指定列
        table.write_number(num1, t_h_out)
        table.write_number(num2, t_c_out)
    workbook.close()
    print("已完成")


main()