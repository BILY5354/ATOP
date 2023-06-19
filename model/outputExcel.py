import os
from datetime import datetime
from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference, Series
from tool.stytle import CellColor


""" 
    total_batch_dict 所有批次的缺陷
    total_yield_dur_dict
    total_ver_list
    target_dir
"""


def output_excel(total_batch_dict, total_yield_dur_dict,
                 total_ver_list, target_dir, ver_files):

    output_path = ".\\output"
    now_time = datetime.now().strftime('%Y_%m_%d_%H_%M_%S')
    create_new_folder_name = os.path.join(output_path, now_time)
    output_excel_path = fr"{create_new_folder_name}\test1.xlsx"

    # 将字典所有键存入列表中 存的是数据集名称
    total_batch_names = [i for i in total_batch_dict.keys()]
    total_yield_dur_names = [i for i in total_yield_dur_dict.keys()] 

    wb = Workbook()
    number_of_total_batch = len(total_batch_dict)  # 数据集数量
    number_of_total_yield_dur = len(total_yield_dur_dict)

    ws = wb.active
    ws.title = "汇总表格"  # 将第一个改名
    # 有几个数据集 创建多少个工作簿
    for i in range(0, number_of_total_batch, 1):
        wb.create_sheet(f'{total_batch_names[i]}')

    # * 塞入数据 注意 current_dataset_nu 下标从 0 开始
    for current_dataset_nu in range(0, number_of_total_batch, 1):
        current_batch_name = total_batch_names[current_dataset_nu]  # 遍历的批次号名字
        ws = wb[f'{current_batch_name}']  # 选中对应的工作簿

        if len(total_batch_dict[current_batch_name]) == 0:
            continue

        current_batch = total_batch_dict[current_batch_name]  # 当前(遍历)数据集列表
        current_batch_ver_list = current_batch[0]['ver']  # 当前(遍历)数据集对应版本数列表

        # 开始插入缺陷名称 注意从第二列开始插入
        for col in range(0, len(current_batch), 1):
            edit_col = int(2 + col)
            ws.cell(
                row=1, column=edit_col).value = current_batch[col]['name']
            ws.cell(row=1, column=edit_col).fill = CellColor.green_fill
            ws.cell(row=1, column=edit_col).border = CellColor.black_border

        # 插入 第一列数据集版本号
        current_cell_index = 1
        for row in range(0, len(current_batch_ver_list), 1):
            edit_row = int(2+row)  # 从第二行开始操作
            ws.cell(row=edit_row,
                    column=1).value = f'{list(current_batch_ver_list[row])[0]}'
            ws.cell(row=edit_row, column=1).fill = CellColor.yellow_fill
            ws.cell(row=edit_row, column=1).border = CellColor.black_border
            current_cell_index += 1

        # 开始插入数据 按行插入
        current_cell_index = 0  # 对应版本号 从0开始
        # todo total_batch_dict[current_dataset_nu][0]['ver'] 中的 0 固定死因为每个括号包含的版本数是一样的
        for row in range(0, len(current_batch_ver_list), 1):
            for col in range(0, len(current_batch), 1):
                # 特定数据集版本中的字典(数量为1)
                specificDataSetVer = current_batch[col]['ver'][current_cell_index]
                edit_row = int(2+row)  # 行数从第二行开始操作
                edit_col = int(2+col)  # 列从第二列开始操作

                # 获取对应数据集缺陷版本的数量
                specificKey = [i for i in specificDataSetVer.keys()][0]

                ws.cell(row=edit_row,
                        column=edit_col).value = specificDataSetVer[specificKey]
                # print(specificDataSetVer)

            current_cell_index += 1

        # 开始画折线图
        lineChart = LineChart()
        chartData = Reference(ws, min_row=1, max_row=ws.max_row,
                              min_col=2, max_col=ws.max_column)
        titleData = Reference(ws, min_row=2, max_row=ws.max_row,
                              min_col=1, max_col=1)

        series = Series(chartData)
        series.marker

        # todo from_rows=False 默认以没一列为数据系列
        lineChart.title = f'{current_batch_name}报告'
        lineChart.add_data(chartData, from_rows=False, titles_from_data=True)
        lineChart.set_categories(titleData)
        ws.add_chart(lineChart, f'F{ws.max_row+10}')

        # 插入对应的vrp文件信息
        edit_col = ws.max_column + 1
        ws.cell(row=1, column=edit_col).value = "vrp文件"
        ws.cell(row=1, column=edit_col).fill = CellColor.green_fill
        ws.cell(row=1, column=edit_col).border = CellColor.black_border
        # 新增加入对应版本的vrp信息
        for row in range(0, len(ver_files[current_batch_name]), 1):
            edit_row = int(2 + row)  # 从第二行开始操作
            ws.cell(row=edit_row, column=edit_col).value = ver_files[current_batch_name][row]

    # * 插入汇总表格数据
    ws = wb["汇总表格"]

    ws.cell(row=1, column=1).value = "良率"
    ws.cell(row=1, column=1).fill = CellColor.blue_fill
    ws.cell(row=1, column=1).border = CellColor.black_border

    # 开始插入批次信息 从第二列开始插入
    for col in range(0, number_of_total_yield_dur, 1):
        # link = "test1.xlsx"+f"#{total_batch_names[col]}!A1"
        link = fr"Y:\GeRun\SecondaryWire\4#\VTSimulation\{total_batch_names[col]}"
        edit_col = int(2 + col)
        ws.cell(
            row=1, column=edit_col).value = total_yield_dur_names[col]
        ws.cell(row=1, column=edit_col).style = "Hyperlink"
        ws.cell(row=1, column=edit_col).fill = CellColor.green_fill
        ws.cell(row=1, column=edit_col).border = CellColor.black_border
        ws.cell(row=1, column=edit_col).hyperlink = link

    # 开始插入每个版本信息 从二行开始插入
    for row in range(0, len(total_ver_list), 1):
        edit_row = int(2+row)  # 从第二行开始操作
        # link = fr"\\192.168.0.11\Data\GeRun\SecondaryWire\4#\VT报告文件db3\{total_ver_list[row]}\SaveFile.txt"
        ws.cell(row=edit_row, column=1).value = total_ver_list[row]
        ws.cell(row=edit_row, column=1).fill = CellColor.yellow_fill
        ws.cell(row=edit_row, column=1).border = CellColor.black_border

    # 开始插入每个批次对应的单元格信息
    for col in range(0, number_of_total_yield_dur, 1):
        current_yield_dur = total_yield_dur_dict[total_yield_dur_names[col]]

        insertNu = 0  # 单个数据集中 插入的良率位置 从0开始
        for row in range(0, len(total_ver_list), 1):
            edit_row = int(2+row)
            edit_col = int(2+col)

            # 判断该良率版本是否列相同
            if total_ver_list[row] == current_yield_dur[insertNu][0]:
                value = float(current_yield_dur[insertNu][1].strip('%')) / 100
                ws.cell(row=edit_row, column=edit_col).value = value
                # #! 是否必要 增加超链接功能
                ws.cell(row=edit_row, column=edit_col).style = "Hyperlink"
                link = f'{target_dir[total_yield_dur_names[col]][insertNu]}'.replace(r'\\192.168.0.11\Data', r'Y:', 1)
                ws.cell(row=edit_row, column=edit_col).hyperlink = link
                insertNu += 1
            else:
                # print(f'row{edit_row} col{edit_col}')
                continue

    # 开始画折线图
    lineChart = LineChart()
    chartData = Reference(ws, min_row=1, max_row=ws.max_row,
                          min_col=2, max_col=ws.max_column)
    titleData = Reference(ws, min_row=2, max_row=ws.max_row,
                          min_col=1, max_col=1)

    series = Series(chartData)
    series.marker

    # todo from_rows=False 默认以没一列为数据系列
    lineChart.title = f'汇总良率报告 单位(%)'
    lineChart.add_data(chartData, from_rows=False, titles_from_data=True)
    lineChart.set_categories(titleData)
    lineChart.y_axis.scaling.max = 1  # 设置Y轴  最大值
    ws.add_chart(lineChart, f'F{ws.max_column+6}')

    # * 第二部分 总表
    BottomRow = ws.max_row+2
    ws.cell(row=BottomRow, column=1).value = "耗时(s)"
    ws.cell(row=BottomRow, column=1).fill = CellColor.blue_fill
    ws.cell(row=BottomRow, column=1).border = CellColor.black_border

    # 开始插入批次信息 从第二列开始插入
    for col in range(0, number_of_total_yield_dur, 1):
        # 修改下超链接
        aa = "数据集1"
        link = "test1.xlsx"+f"#{total_batch_names[col]}!A1"
        edit_col = int(2 + col)
        ws.cell(
            row=BottomRow, column=edit_col).value = total_yield_dur_names[col]
        ws.cell(row=BottomRow, column=edit_col).style = "Hyperlink"
        ws.cell(row=BottomRow, column=edit_col).fill = CellColor.green_fill
        ws.cell(row=BottomRow, column=edit_col).border = CellColor.black_border
        ws.cell(row=BottomRow, column=edit_col).hyperlink = link

    # 开始插入每个版本信息 从二行开始插入
    for row in range(0, len(total_ver_list), 1):
        edit_row = int(BottomRow+row+1)  # 从底部加1行插入
        ws.cell(row=edit_row,
                column=1).value = total_ver_list[row]
        ws.cell(row=edit_row, column=1).fill = CellColor.yellow_fill
        ws.cell(row=edit_row, column=1).border = CellColor.black_border

    # 开始插入每个批次对应的单元格信息
    for col in range(0, number_of_total_yield_dur, 1):
        current_yield_dur = total_yield_dur_dict[total_yield_dur_names[col]]

        insertNu = 0  # 单个数据集中 插入的良率位置 从0开始
        for row in range(0, len(total_ver_list), 1):
            edit_row = int(BottomRow+row+1)
            edit_col = int(2+col)

            # 判断该良率版本是否列相同
            if total_ver_list[row] == current_yield_dur[insertNu][0]:

                ws.cell(row=edit_row,
                        column=edit_col).value = float(current_yield_dur[insertNu][2])
                ws.cell(row=edit_row, column=edit_col).style = "Hyperlink"
                db3_file_name = target_dir[total_yield_dur_names[col]][insertNu].split('\\')[-1]
                link = "http://192.168.0.101:8022/db_file_summary/0/" + db3_file_name
                ws.cell(row=edit_row,column=edit_col).hyperlink = link
                insertNu += 1
            else:
                continue

    # 开始画折线图
    dur_line_chart = LineChart()
    chartData2 = Reference(ws, min_row=BottomRow, max_row=ws.max_row,
                           min_col=2, max_col=ws.max_column)
    titleData2 = Reference(ws, min_row=int(BottomRow+1), max_row=ws.max_row,
                           min_col=1, max_col=1)

    series2 = Series(chartData2)
    series2.marker

    # todo from_rows=False 默认以每一列为数据系列
    dur_line_chart.title = f'汇总时间报告 单位(s)'
    dur_line_chart.add_data(chartData2, from_rows=False, titles_from_data=True)
    dur_line_chart.set_categories(titleData2)
    ws.add_chart(dur_line_chart, f'F{ws.max_column+20}')

    os.makedirs(fr'{create_new_folder_name}')
    wb.save(output_excel_path)
    os.startfile(fr'{create_new_folder_name}')
