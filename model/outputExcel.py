from openpyxl import Workbook
from openpyxl.styles import PatternFill, Border, Side
from openpyxl.chart import LineChart, Reference
from openpyxl.styles.colors import Color



def output_excel(mul_version_defects_dict):
    EXCELPATH = ".\\output\\test1.xlsx"
    # 现在传的数据一次传一个版本
    GETDATA = mul_version_defects_dict

    # print(GETDATA[0][1]['ver'][1])
    getDKEY = []
    # 将字典所有键存入列表中
    for i in GETDATA.keys():
        getDKEY.append(i)
    # print(getDKEY)

    wb = Workbook()
    DATASETNU = len(GETDATA)  # 数据集数量
    ws = wb.active
    ws.title = getDKEY[0]  # 将第一个改名
    # 有几个数据集 创建多少个 不需要创建第一个
    for i in range(1, DATASETNU, 1):
        wb.create_sheet(f'{getDKEY[i]}')

    green = Color(rgb="00B050")
    green_fill = PatternFill(start_color=green,
                            end_color=green, fill_type='solid')

    yellow = Color(rgb="FFFF00")
    yellow_fill = PatternFill(start_color=yellow,
                            end_color=yellow, fill_type='solid')

    border_style = Side(style="thin", color='000000')
    black_border = Border(left=border_style, right=border_style,
                        top=border_style, bottom=border_style)


    # * 塞入数据 注意 currentDSNu 下标从 0 开始
    for currentDSNu in range(0, DATASETNU, 1):
        if len(GETDATA[getDKEY[currentDSNu]]) == 0 :
            continue

        ws = wb[f'{getDKEY[currentDSNu]}']  # 选中对应的工作簿

        # 开始插入缺陷名称 注意从第二列开始插入
        for col in range(0, len(GETDATA[getDKEY[currentDSNu]]), 1):
            editCol = int(2 + col)
            ws.cell(
                row=1, column=editCol).value = GETDATA[getDKEY[currentDSNu]][col]['name']
            ws.cell(row=1, column=editCol).fill = green_fill
            ws.cell(row=1, column=editCol).border = black_border

        # 开始插入 第一列数据集
        verNu = 1
        for row in range(0, len(GETDATA[getDKEY[currentDSNu]][0]['ver']), 1):
            editRow = int(2+row)  # 从第二行开始操作
            ws.cell(row=editRow, column=1).value = f'版本{verNu}'
            ws.cell(row=editRow, column=1).fill = yellow_fill
            ws.cell(row=editRow, column=1).border = black_border
            verNu += 1

        # 开始插入数据 按行插入
        verNu = 0  # 对应版本号 从0开始
        # todo GETDATA[currentDSNu][0]['ver'] 中的 0 固定死因为每个括号包含的版本数是一样的
        for row in range(0, len(GETDATA[getDKEY[currentDSNu]][0]['ver']), 1):
            for col in range(0, len(GETDATA[getDKEY[currentDSNu]]), 1):
                editRow = int(2+row)  # 行数从第二行开始操作
                editCol = int(2+col)  # 列从第二列开始操作

                ws.cell(
                    row=editRow, column=editCol).value = GETDATA[getDKEY[currentDSNu]][col]['ver'][verNu]
                # print(GETDATA[getDKEY[currentDSNu]][col]['ver'][verNu])

            verNu += 1


        # 开始画折线图
        lineChart = LineChart()
        chartData = Reference(ws, min_row=1, max_row=ws.max_row,
                            min_col=2, max_col=ws.max_column)
        titleData = Reference(ws, min_row=2, max_row=ws.max_row,
                            min_col=1, max_col=1)
        # todo from_rows=False 默认以没一列为数据系列
        lineChart.title = f'{getDKEY[currentDSNu]}报告'
        lineChart.add_data(chartData, from_rows=False, titles_from_data=True)
        lineChart.set_categories(titleData)
        ws.add_chart(lineChart, f'F{ws.max_row+10}')



    # 画图 https://blog.csdn.net/qq_34699587/article/details/115529544


    wb.save(EXCELPATH)