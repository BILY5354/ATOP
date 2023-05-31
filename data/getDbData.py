import sqlite3
import os
from data.getFilePath import get_file_path
from pathlib import Path


def get_dbData(targetDir):
    dbFiles_path_dict = get_file_path(targetDir)
    mul_version_defects_dict = {}
    for data_set in dbFiles_path_dict:
        dbFile_path_list = dbFiles_path_dict[data_set]
        for db_file_path in dbFile_path_list:
            db_file_path_split = db_file_path.split('\\')
            # 版本号
            ver_id = db_file_path_split[-2]
            db_file_path = Path(db_file_path)
            sql = "select defects from REP_INST LIMIT 1"
            if data_set not in mul_version_defects_dict:
                result_sql = execute_sql(db_file_path, sql)[0]
                file_defects_list = []
                defects = result_sql[0]
                if (defects == '' or defects == None):
                    print('报告非正常结束,defects 信息不全')
                    break
                defects_list = defects.split(";")
                for defect in defects_list:
                    defect_list = defect.split(",")
                    if (defect_list[5] != '0'):
                        defect_dict = {}
                        ver_name_count_dict = {}
                        ver_defect_list = []
                        defect_code = defect_list[0]
                        defect_name = defect_list[3]
                        defect_id = defect_list[2]
                        name_code = defect_name+'('+defect_code+')'
                        # {'id': '1', 'name': '基准点(REF)', 'ver': [{ghgh:3}, {gfgff:3}]}
                        defect_dict['id'] = defect_id
                        defect_dict['name'] = name_code
                        defect_count = int(defect_list[5])
                        ver_name_count_dict[ver_id] = defect_count
                        ver_defect_list.append(ver_name_count_dict)
                        defect_dict['ver'] = ver_defect_list
                        file_defects_list.append(defect_dict)
                mul_version_defects_dict[data_set] = file_defects_list
            else:
                file_defect_data_list = []
                db_file_path = Path(db_file_path)
                result_sql = execute_sql(db_file_path, sql)[0]
                file_defects_list = []
                defect_id_list = []
                defects = result_sql[0]
                if (defects == '' or defects == None):
                    print('报告非正常结束,defects 信息不全')
                    break
                defects_list = defects.split(";")
                for defect in defects_list:
                    defect_list = defect.split(",")
                    if (defect_list[5] != '0'):
                        defect_code = defect_list[0]
                        defect_name = defect_list[3]
                        defect_id = defect_list[2]
                        defect_id_list.append(defect_id)
                        name_code = defect_name+'('+defect_code+')'
                        defect_count = defect_list[5]
                        id_name_count = defect_id + ',' + name_code + ',' + defect_count
                        file_defect_data_list.append(id_name_count)
                pre_defects_list = mul_version_defects_dict[data_set]
                for def_dict in pre_defects_list:
                    pre_id = def_dict['id']
                    pre_name = def_dict['name']
                    pre_ver_list = def_dict['ver']
                    if pre_id not in defect_id_list:
                        pre_ver_list.append(0)
                    else:
                        for defect_list in file_defect_data_list:
                            ver_name_count_dict = {}
                            defect_split = defect_list.split(',')
                            id = defect_split[0]
                            name = defect_split[1]
                            count = int(defect_split[2])
                            if (pre_id == id):
                                ver_name_count_dict[ver_id] = count
                                pre_ver_list.append(ver_name_count_dict)

    return mul_version_defects_dict


def execute_sql(db_file_path, sql):
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
    
    return sql_result