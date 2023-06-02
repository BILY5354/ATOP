import os
import re


def get_file_path(targetDir):
    savefiles = []
    # todo 当前遍历到的目录名 当前dirpath子目录名列表 当前dirpath所有文件名列表
    for (dirpath, dirnames, filenames) in os.walk(targetDir):
        for fn in filenames:
            fpath = os.path.join(dirpath, fn)
            if "SaveFile.txt" in fpath:
                savefiles.append(fpath)

    batchVerNu = {}  # { 批次号 : [版本1地址，版本2地址] }
    for sf in savefiles:
        # print(f'文件:{sf}')
        with open(sf) as f:
            pathList = [path for path in f.read().splitlines() if path]
        pathGroupList = [pathList[i:i+4] for i in range(0, len(pathList), 4)]

        for oneGroup in pathGroupList:
            if "SimulRecipe1.srp" in oneGroup[1]:
                batchNu = re.findall(
                    r'VTSimulation\\(.*)\\SimulRecipe1', oneGroup[1])[0]
            else:
                batchNu = re.findall(r'([^\\]+)\.srp', oneGroup[1])[0]
            path = re.sub(r'Y:', r'\\\\192.168.0.11\\Data', fr'{oneGroup[2]}')

            if batchNu != None and path != None:
                # print(f'{batchNu} {path}')
                if batchNu not in batchVerNu:
                    batchVerNu[batchNu] = []
                    batchVerNu[batchNu].append(path)
                else:
                    batchVerNu[batchNu].append(path)

    return batchVerNu
