#-*- coding:utf-8 -*-

import  xlrd
import os
import traceback

import utils

logfile = open("logfile_get_airport.log", "w")

def print_info(ofile, infolist):
    
    for info in infolist:
        ofile.write(
            "{:5s}{:6s}".format(info[0], info[1])
            + "{:.6f}".format(info[2]).zfill(9) + "    "
            + "{:.6f}".format(info[3]).zfill(10) + "   "
            + "\n"
        )

def get_info(old_info):
    # print(old_info)
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

# 主函数
def main():

    file = xlrd.open_workbook("./original_data.xlsx")
    ofile = open("Airports_4.txt", 'w')

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


    # 读取机场信息
    airport_dc = {}
    table = file.sheets()[0]
    nrows = table.nrows
    for i in range(1, nrows):
        row = table.row_values(i)
        # 判断数字是否正常
        is_err, err_list = utils.check_old_info(row)
        if is_err:
            logfile.write(f"[WARNING] 数字信息错误 sheet{1}, row{i+1}, column{err_list} 已经补0\n")
        try:
            info = get_info(row)
        except Exception as e:
            # 记录错误
            logfile.write("#"*10+"\n")
            logfile.writelines(traceback.format_exc())
            logfile.write("#"*10+"\n")
            # 忽略错误
            pass
            # sys.exit(0)
        airport_dc[info[0]] = info


    sid_info_dic = {}

    # 开始读取离场信息
    table = file.sheets()[8]
    nrows = table.nrows
    last_airport = table.row_values(1)[0].strip()
    last_point = ""
    last_number = -1000

    # count = 0
    for i in range(1, nrows):
        row = table.row_values(i)
        if row[0].strip() != last_airport or row[4] < last_number:
            # 查找是否存在航线点的信息
            if last_point in dic.keys():
                if last_airport not in sid_info_dic.keys():
                    sid_info_dic[last_airport] = []
                index = True
                for l in sid_info_dic[last_airport]:
                    if l[1] == last_point:
                        index = False 
                if index:
                    sid_info_dic[last_airport].append(["SID", last_point, dic[last_point][1], dic[last_point][2]])
            if row[0].strip() != last_airport:
                last_airport = row[0].strip()
                last_number = -10000
                last_point = ""
                continue
        last_number = row[4]
        last_point = row[5].strip()

    star_info_dic = {}
    # 开始读取进场信息
    table = file.sheets()[9]
    nrows = table.nrows
    last_airport = table.row_values(1)[0].strip()
    last_number = 1000

    # count = 0

    for i in range(1, nrows):
        row = table.row_values(i)
        if row[0].strip() != last_airport or row[4] < last_number:
            # 查找是否存在航线点的信息
            point = row[5].strip()
            airport = row[0].strip()
            if point in dic.keys():
                if airport not in star_info_dic.keys():
                    star_info_dic[airport] = []
                index = True
                for l in star_info_dic[airport]:
                    if l[1] == point:
                        index = False 
                if index:
                    star_info_dic[airport].append(["STAR", point, dic[point][1], dic[point][2]])
            if row[0].strip() != last_airport:
                last_airport = airport
                last_number = 1000
                continue
        last_number = row[4]


    for airport in airport_dc.keys():
        ofile.write("A {:5s}".format(airport)
        + "{:.6f}".format(airport_dc[airport][1]).zfill(9) + " "
        + "{:.6f}".format(airport_dc[airport][2]).zfill(10) + " "
        + "10 5 5 5\n"
        )
        if airport in sid_info_dic.keys():
            print_info(ofile, sid_info_dic[airport])
        if airport in star_info_dic.keys():
            print_info(ofile, star_info_dic[airport])

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
    print("[LOG] 程序执行完毕,为您打印log文件`logfile_get_airport.log`")
    print()
    with open("logfile_get_airport.log", "r") as logfile:
        print(logfile.read())
    print()
    os.system('pause')
