#-*- coding:utf-8 -*-

import  xlrd
import os
import traceback

import utils

logfile = open("logfile_get_route.log", "w")

def print_info(ofile, route_id, info_list):
    for i in range(len(info_list) - 1):
        for j in range(i+1, len(info_list)):
            # print i j
            ofile.write("{:7s}{:7s}".format(route_id, info_list[i][0])
             + "{:.8f}".format(info_list[i][1]).zfill(11) + "  "
             + "{:.8f}".format(info_list[i][2]).zfill(12) + " "
             + "{:7s}".format(info_list[j][0])
             + "{:.8f}".format(info_list[j][1]).zfill(11) + "  "
             + "{:.8f}".format(info_list[j][2]).zfill(12) + " "
             + "1 10000\n"
             )
            ofile.write("{:7s}{:7s}".format(route_id, info_list[j][0])
             + "{:.8f}".format(info_list[j][1]).zfill(11) + "  "
             + "{:.8f}".format(info_list[j][2]).zfill(12) + " "
             + "{:7s}".format(info_list[i][0])
             + "{:.8f}".format(info_list[i][1]).zfill(11) + "  "
             + "{:.8f}".format(info_list[i][2]).zfill(12) + " "
             + "2 10000\n"
             )

def get_info(old_info):
    info = []
    latitude = old_info[2] + old_info[3] / 60.0 + old_info[4] / 360.0
    if old_info[1].strip() == 'S':
        latitude *= -1
    longitude = old_info[6] + old_info[7] / 60.0 + old_info[8] / 360.0
    if old_info[5].strip() == 'W':
        longitude *= -1
    info.append(old_info[0].strip())
    info.append(latitude)  # 维度
    info.append(longitude) # 经度
    # [   ID    纬度    经度    ]
    return info

def main():
    file = xlrd.open_workbook("./original_data.xlsx")
    ofile = open("awys3.txt", 'w')

    # 建立航线点经纬度数据库  用字典存储
    dic = {}

    for num in range(2,7):
        sheet = file.sheets()[num]
        col = 0
        if num <= 3:
            col = 1
        nrows = sheet.nrows    
        # 开始读取数据
        for i in range(1, nrows):
            row = sheet.row_values(i)    
            old_info = row[col:]
            # 判断数字是否正常
            is_err, err_list = utils.check_old_info(old_info)
            if is_err:
                logfile.write(f"[WARNING] 数字信息错误 sheet{num+1}, row{i+1}, column{err_list} 已经补0\n")
            try:
                info = get_info(old_info)  #TODO: ERROR
            except Exception as e:
                # 记录错误
                logfile.write("#"*10+"\n")
                logfile.writelines(traceback.format_exc())
                logfile.write("#"*10+"\n")
                # 忽略错误
                pass
                # sys.exit(0)
            dic[info[0]] = info

    # 开始读取航线信息
    table = file.sheets()[7]
    nrows = table.nrows
    route_id = table.row_values(1)[0].strip()
    info_list = []

    count = 0

    for i in range(1, nrows):
        row = table.row_values(i)
        if row[0].strip() != route_id:
            # 输出机场对应的航路信息
            print_info(ofile,route_id, info_list)
            info_list.clear()   
            route_id = row[0].strip()
        point = row[2].strip()
        if point in dic.keys():
            # 数据库中找到
            info_list.append(dic[point])
        else:
            count += 1
            pass

    print_info(ofile, route_id, info_list)

    print("共有", count, "个航路点没有找到经纬度信息")
    logfile.write(f"[LOG]共有{count}个航路点没有找到经纬度信息\n")
    ofile.close()

if __name__ == '__main__':
    print("[LOG]程序开始运行......")
    try:
        main()
    except Exception as e:
        # 记录错误
        logfile.write("#"*10+"\n")
        logfile.writelines(traceback.format_exc())
        logfile.write("#"*10+"\n")
        # 忽略错误
        pass
        # sys.exit(0)
    finally:
        logfile.close()
    print("[LOG] 程序执行完毕,为您打印log文件`logfile_get_route.log`")
    print()
    with open("logfile_get_route.log", "r") as logfile:
        print(logfile.read())
    print()
    os.system('pause')
    


