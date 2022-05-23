import xlwings as xw
import time
def main():
    lis = set("abcdefghijklmnopqrstuvwxyz")
    lis2 = set("acegikmoqsuwy")
    lis3 = set("bdfhjlnprtvxz")
    wb = xw.Book.caller()
    for i in lis:
        for j in range(1,58):
            wb.sheets[0].range('%s%d' %(i, j)).color = (49,117+j,159+j)
    for j in range(1,58):
        for i in lis2:
            wb.sheets[0].range('%s%d' %(i, j)).color = (140-2*j,30,50)
    
    for j in range(1,58):
        for i in lis3:
            wb.sheets[0].range('%s%d' %(i, j)).color = (255,133+j,133+j)
    wb.sheets[0].range('C20').value = "北"
    wb.sheets[0].range('C20').api.Font.Size = 72			# 设置字号为15
    wb.sheets[0].range('C20').api.Font.Bold = True
    wb.sheets[0].range('H20').value = "科"
    wb.sheets[0].range('H20').api.Font.Size = 72		# 设置字号为15
    wb.sheets[0].range('H20').api.Font.Bold = True
    wb.sheets[0].range('M20').value = "大"
    wb.sheets[0].range('M20').api.Font.Size = 72		# 设置字号为15
    wb.sheets[0].range('M20').api.Font.Bold = True
    wb.sheets[0].range('R20').value = "牛"
    wb.sheets[0].range('R20').api.Font.Size = 72		# 设置字号为15
    wb.sheets[0].range('R20').api.Font.Bold = True
    wb.sheets[0].range('W20').value = "逼"
    wb.sheets[0].range('W20').api.Font.Size = 72		# 设置字号为15
    wb.sheets[0].range('W20').api.Font.Bold = True
@xw.func
def hello(name):
    return f"Hello {name}!"


if __name__ == "__main__":
    xw.Book("demo.xlsm").set_mock_caller()
    main()

#*(i+1)+2*j  %(i, j)