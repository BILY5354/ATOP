import os
import re


def get_file_path(target_dir):
    savefiles = []
    # todo 当前遍历到的目录名 当前dirpath子目录名列表 当前dirpath所有文件名列表
    for (dir_path, dir_names, filenames) in os.walk(target_dir):
        for fn in filenames:
            fpath = os.path.join(dir_path, fn)
            if "SaveFile.txt" in fpath:
                savefiles.append(fpath)

    total_batch_ver_dict = {}  # { 批次号 : [版本1地址，版本2地址] }
    for sf in savefiles:
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

            if batch_name != None and path != None:
                # print(f'{batch_name} {path}')
                if batch_name not in total_batch_ver_dict:
                    total_batch_ver_dict[batch_name] = []
                    total_batch_ver_dict[batch_name].append(path)
                else:
                    total_batch_ver_dict[batch_name].append(path)

    return total_batch_ver_dict


def get_specific_ver_vrp(target_dir):
    savefiles = []
    # todo 当前遍历到的目录名 当前dirpath子目录名列表 当前dirpath所有文件名列表
    for (dir_path, dir_names, filenames) in os.walk(target_dir):
        for fn in filenames:
            fpath = os.path.join(dir_path, fn)
            if "SaveFile.txt" in fpath:
                savefiles.append(fpath)

    print(savefiles)

    total_vrp_files = {}  # { 批次号 : vrp }
    for sf in savefiles:
        # print(f'文件:{sf}')
        with open(sf) as f:
            total_path = [path for path in f.read().splitlines() if path]
        group_path = [total_path[i:i + 4] for i in range(0, len(total_path), 4)]

        for oneGroup in group_path:

            if "SimulRecipe1.srp" in oneGroup[1]:
                batch_name = re.findall(
                    r'VTSimulation\\(.*)\\SimulRecipe1', oneGroup[1])[0]
            else:
                batch_name = re.findall(r'([^\\]+)\.srp', oneGroup[1])[0]
            correspond_vrp_file = re.findall(r'([^\\]+)\.vrp', fr'{oneGroup[0]}')[0] + ".vrp"

            if batch_name != None and correspond_vrp_file != None:
                # print(f'{batch_name} {path}')
                if batch_name not in total_vrp_files:
                    total_vrp_files[batch_name] = []
                    total_vrp_files[batch_name].append(correspond_vrp_file)
                else:
                    total_vrp_files[batch_name].append(correspond_vrp_file)

    return total_vrp_files
