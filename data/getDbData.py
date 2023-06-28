import sqlite3
import os
import json
import re
from data.getFilePath import get_file_path
from pathlib import Path
from dateutil.parser import parse
import time
from functools import wraps


def get_db_defect_data(targetDir):
    total_files_info = get_file_path(targetDir)
    total_db_file_path_dict = total_files_info['totalBatchVerPath']
    total_batch_name = [i for i in total_db_file_path_dict.keys()]
    total_version_defects_dict = {}  # 总字典 所有版本批次号的缺陷 

    for current_batch_name in total_batch_name:
        db_file_path_list = total_db_file_path_dict[current_batch_name]

        # 在一开始便判断总字典是否存在此批次号
        if current_batch_name not in total_version_defects_dict:
            total_version_defects_dict[current_batch_name] = []
            current_batch_defect_id_list = []  # 保存本批次的所有缺陷id 且该id顺序与字典中相同

        for db_file_path in db_file_path_list:

            db_file_path = Path(db_file_path)
            version_name = re.findall(r'VT报告文件db3\\(.*?)\\', str(db_file_path))[0]  # ver_id

            query_sql = "select defects from REP_INST LIMIT 1"
            query_sql_result = execute_sql(db_file_path, query_sql)[0][0]  # 拿字符
            if query_sql_result is None:
                print('报告非正常结束,defects 信息不全')
            if not query_sql_result.endswith(';'):  # 在最后加分号方便re分组
                query_sql_result += ";"
            one_batch_defects_list = re.findall(r'(.*?;)', query_sql_result)

            for defect in one_batch_defects_list:
                split_defect_list = defect.split(',')
                # 存在大于0的缺陷 
                if int(split_defect_list[1]) == 1 and int(split_defect_list[5]) > 0:

                    # 统一生成本次 {'版本号' : 数量} 用于加入
                    current_ver_nu_of_defect = build_ver_nu_of_defect_unit(version_name, split_defect_list[5])
                    current_defect_id = split_defect_list[2]

                    # 无此缺陷信息 需要构建
                    if current_defect_id not in current_batch_defect_id_list:
                        current_batch_defect_dict = build_one_defect_unit(split_defect_list[2],
                                                                          split_defect_list[3],
                                                                          split_defect_list[0])

                        current_batch_defect_dict['ver'].append(current_ver_nu_of_defect)
                        current_batch_defect_id_list.append(current_defect_id)

                        # 将首次创建单个缺陷的字典加入该批次的列表中 如果已有此id直接append
                        total_version_defects_dict[current_batch_name].append(current_batch_defect_dict)

                    else:  # 有此缺陷字典 直接加入 {'版本号' : 数量} 首先找此缺陷在列表的位置
                        try:
                            # 该索引是当前缺陷在该批次列表中的下标
                            current_index = current_batch_defect_id_list.index(current_defect_id)
                        except:
                            # ! 是否需要加入信息插入Excel表中
                            print("没有此defrct Id信息")
                            break

                        defect_of_specific_id_dict = total_version_defects_dict[current_batch_name][current_index]
                        defect_of_specific_id_dict["ver"].append(current_ver_nu_of_defect)

                else:  # 本缺陷数为0 看下一个
                    continue

    return total_version_defects_dict, total_files_info


def time_it(func):
    @wraps(func)
    def inner(*args, **kwargs):
        start = time.time()
        func(*args, **kwargs)
        end = time.time()
        # print('用时:{}秒'.format(end-start))

    return inner


def execute_sql(db_file_path, sql):
    # start = time.time()
    if (not os.path.isfile(db_file_path)):
        return False
    conn = sqlite3.connect(db_file_path)
    # 设置一个text_factory，告诉decode()忽略此类错误(utf-8无法解读)
    conn.text_factory = lambda b: b.decode(errors='ignore')
    cursor = conn.cursor()
    try:
        cursor.execute(sql)
        sql_result = cursor.fetchall()
    except:
        print('数据库执行 SQL 有误，请确定数据库文件是否有内容')
    # 关闭游标：
    cursor.close()
    # 提交事务
    conn.commit()
    # 关闭连接
    conn.close()

    # end = time.time()
    # print('用时:{}秒'.format(end-start))
    return sql_result


# @time_it
def get_yied_dur_data(targetDir):
    total_files_info = get_file_path(targetDir)
    total_db_file_path_dict = total_files_info['totalBatchVerPath']
    total_yield_dur_dict = {}
    ver_list = []
    for data_set in total_db_file_path_dict:
        dbFile_path_list = total_db_file_path_dict[data_set]
        mul_ver_data_list = []
        for db_file_path in dbFile_path_list:
            lot_yied_dur_list = [0, 0, 0]
            db_file_path_split = db_file_path.split('\\')
            # 版本号
            ver_id = db_file_path_split[-2]
            if ver_id not in ver_list:
                ver_list.append(ver_id)
            db_file_path = Path(db_file_path)
            sql = "select startTime,endTime,runAttrib from REP_INST LIMIT 1"
            # print(db_file_path)
            sql_res = execute_sql(db_file_path, sql)
            result_sql = sql_res[0]
            startTime = parse(result_sql[0])
            endTime = parse(result_sql[1])
            runAttrib_json = json.loads(result_sql[2])
            fail_count = (
                    int(runAttrib_json["FAIL"]) + int(runAttrib_json['ASSISTFAIL']))
            good_count = (
                    int(runAttrib_json["GOOD"]) + int(runAttrib_json['ASSISTPASS']))
            total_count = fail_count + good_count
            yied_float = good_count / total_count
            yied = "%.2f%%" % (yied_float * 100)
            duration_seconds = (endTime - startTime).seconds
            # m,s = divmod(duration_seconds,60)
            # duration = str(m) + '分' + str(s) + '秒'
            lot_yied_dur_list[0] = ver_id
            lot_yied_dur_list[1] = yied
            lot_yied_dur_list[2] = int(duration_seconds)
            mul_ver_data_list.append(lot_yied_dur_list)
            total_yield_dur_dict[data_set] = mul_ver_data_list
    return total_yield_dur_dict, ver_list


# 构建缺陷整体字典 {'缺陷id': '106','缺陷名字': '塌线缺陷(SW)', 'ver': [最小单位(空)]}
# 没有 version_name 因为其是用于 append 的
def build_one_defect_unit(defect_id, defect_name, defect_code):
    return {
        'id': defect_id,
        'name': defect_name + f'({defect_code})' ,
        'ver': []
    }


# 构建缺陷-数量字典 插入最小单位 {'版本号' : 数量}
def build_ver_nu_of_defect_unit(version_name, number):
    return {
        version_name: number
    }