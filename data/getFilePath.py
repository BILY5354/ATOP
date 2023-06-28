import os
from os.path import isfile, join,isdir
import re

# 查询文件夹的信息 totalBatchVerPath totalVrpDict numberOfBatches
total_files_info = {}  # 所有批次与版本信息 vrp文件信息 批次数量信息


def get_file_path(target_dir):

    save_files = []
    # todo 当前遍历到的目录名 当前dirpath子目录名列表 当前dirpath所有文件名列表
    for (dir_path, dir_names, filenames) in os.walk(target_dir):
        for fn in filenames:
            fpath = os.path.join(dir_path, fn)
            if "SaveFile.txt" in fpath:
                save_files.append(fpath)

    total_batch_ver_dict = {}  # { 批次号 : [版本1地址，版本2地址] }
    total_vrp_files = {}  # { 批次号 : vrp }
    for sf in save_files:
        # print(f'文件:{sf}')
        with open(sf) as f:
            total_path = [path for path in f.read().splitlines() if path]
        group_path = [total_path[i:i+4] for i in range(0, len(total_path), 4)]

        for oneGroup in group_path:
            if "SimulRecipe1.srp" in oneGroup[1]:
                batch_name = re.findall(
                    r'VTSimulation\\(.*)\\SimulRecipe1', oneGroup[1])[0]
            else:
                batch_name = re.findall(r'([^\\]+)\.srp', oneGroup[1])[0]

            path = re.sub(r'Y:', r'\\\\192.168.0.11\\Data', fr'{oneGroup[2]}')
            # path = re.sub(r'C:\\Users\\AOI\\Desktop', r'\\\\192.168.0.130\\gerun', fr'{oneGroup[2]}')
            correspond_vrp_file = re.findall(r'([^\\]+)\.vrp', fr'{oneGroup[0]}')[0] + ".vrp"
            correspond_version_name = re.findall(r'VT报告文件db3\\(.*?)\\', fr'{oneGroup[2]}')[0]

            if batch_name is not None and path is not None:
                # print(f'{batch_name} {path}')
                if batch_name not in total_batch_ver_dict:
                    total_batch_ver_dict[batch_name] = []
                    total_batch_ver_dict[batch_name].append(path)
                else:
                    total_batch_ver_dict[batch_name].append(path)

            if batch_name is not None and correspond_vrp_file is not None:
                # print(f'{batch_name} {path}')
                # 构建格式 {批次号1: {版本号1: 文件, 版本号2: 文件},批次号2...}
                if batch_name not in total_vrp_files:
                    total_vrp_files[batch_name] = {}
                    total_vrp_files[batch_name][correspond_version_name] = correspond_vrp_file
                else:
                    total_vrp_files[batch_name][correspond_version_name] = correspond_vrp_file

    total_files_info['totalBatchVerPath'] = total_batch_ver_dict
    total_files_info['totalVrpDict'] = total_vrp_files
    return total_files_info


def get_number_of_batches(target_dir):
    number_of_specific_batch = {}
    # 拿所有批次号的名字
    total_batches_name = [f for f in os.listdir(target_dir) if isdir(join(target_dir, f))]
    index = -1
    for (dir_path, dir_names, filenames) in os.walk(target_dir):
        count = 0
        for fn in filenames:
            if '.db3' in fn:
                count += 1

        if index > -1:  # 跳过第一次
            number_of_specific_batch[f'{total_batches_name[index]}'] = (count / 2)
        index += 1

    # total_files_info['numberOfBatches'] = number_of_specific_batch
    return number_of_specific_batch
