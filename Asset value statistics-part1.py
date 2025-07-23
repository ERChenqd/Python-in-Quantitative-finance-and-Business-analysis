import ExmailLogin
import FileProcess
import datetime
import os
import zipfile

# td = datetime.datetime.now().strftime('%Y-%m-%d')
#path = os.getcwd()+"\\"+td+"\\"
# pyinstaller -F Test3.py
# 设置路径
td = datetime.datetime.now().strftime('%Y-%m-%d')
path = os.getcwd()+"\\"+td+"\\"
# path = "/Users/stu001/Desktop/爬虫/" + td + "/"
# path = "/Users/chenhaoyu/Downloads/NetValue/" + td + "/"
folder = os.path.exists(path)
if not folder:
    os.makedirs(path)
result_path = path + "净值汇总" + td + ".txt"
result = open(result_path, 'w')

uin = "risk@anzhicapital.com"
pwd = input("请输入邮箱密码:")
netval_mail = ExmailLogin.LoginExmail(uin, pwd)
conn = netval_mail.login_mail()
# netval_mail.get_mail(conn, path, "2019-11-26")
pdate = input("请输入需要提取的邮件收件日期（从何时开始，格式 yyyy-mm-dd）:")
mailtitle, incomingdate, attname = netval_mail.get_mail(conn, path, pdate)
assets = {'SCG919': '安值量化母基金',

          'SCP765': '安值福慧量化1号', 'SCV467': '安值福慧量化2号', 'SCV474': '安值福慧量化3号', 'SGJ083': '安值福慧量化5号',
          'SGM242': '安值福慧量化6号', 'SGM270': '安值福慧量化7号', 'SGH894': '安值福慧量化8号', 'SJF166': '安值福慧量化10号',

          'SET659': '安值量化1号',

          'SGJ141': '安值海天量化1号', 'SGW414': '安值海天量化2号', 'CSC_安值海天量化3号': '安值海天量化3号',
          'CSC_安值海天量化7号': '安值海天量化7号', 'SJV073': '安值海天量化5号', 'SJX652': '安值海天量化6号',

          'SLL277': '安值中盛量化1号A期', 'SLL285': '安值中盛量化1号私募',

          # 后面的是新加入的产品
          'SQS246': '安值多策略进取1号',

          'SSL298': '兴建量化1号',

          'SQY121': '安值量化进取1号'

          }
records = {}

# 解压文件夹内压缩包
for filename in os.listdir(path):
    if filename.find("zip") != -1:
        zip_file = zipfile.ZipFile(path + filename)
        for zip_name in zip_file.namelist():  # 解压
            zip_file.extract(zip_name, path)
            attname.append(zip_name)

pdt_list_1=['安值量化母基金','安值中盛量化1号A期','安值中盛量化1号私募','安值福慧量化1号','安值福慧量化2号','安值福慧量化3号','安值量化进取1号',
            '安值福慧量化5号','安值福慧量化6号','安值福慧量化7号','安值福慧量化8号','安值福慧量化10号','安值海天量化5号','安值多策略进取1号']
pdt_list_2=['安值海天量化6号','安值海天量化2号']#安值量化1号暂时有问题
pdt_list_to_be_fixed=['安值量化1号']
pdt_list_3=['安值海天量化3号','安值海天量化7号']
pdt_list_4=['兴建量化1号']
print("资产名称          估值日期      单位净值    累计净值    资产净值")
result.write("资产名称          估值日期      单位净值    累计净值    资产净值\n")
# records={}
# 2023/11/23 新加了兴建1、多策略进取1、量化进取1
for k, v in assets.items():
    for i in range(len(attname)):
        if (attname[i].find(k) != -1) and (attname[i].find("特殊通知") == -1) and (attname[i].find("zip") == -1):
            # print(v)
            if v in pdt_list_1 and (attname[i].find("txt") == -1):
                # 找到累计单位净值、单位净值和估值日期的行列 (cumulative net value, net value, evaluation date)
                try:
                    cumnv_row = FileProcess.get_xlsrow(path + attname[i], "Sheet1", 1, "累计单位净值")
                except:
                    print('error!')
                    print('当前报错文件：', attname[i])
                    print('当前产品：', v, k)
                nv_col = FileProcess.get_xlscol(path + attname[i], "Sheet1", 4, "单位净值")
                ed_col = FileProcess.get_xlscol(path + attname[i], "Sheet1", 4, "日期")
                # 取三个值（这个会带上格子里的其他字符串，不是纯数字）
                net_val = FileProcess.get_xlsnetval(path + attname[i], "Sheet1", 4, nv_col + 1)
                eval_date = FileProcess.get_xlsnetval(path + attname[i], "Sheet1", 4, ed_col + 1)
                cum_netval = FileProcess.get_xlsnetval(path + attname[i], "Sheet1", cumnv_row + 1, 2)

                # 找到总资产净值 （asset net value）
                if v.find("安值福慧量化10号") != -1 or v.find("安值福慧量化5号") != -1 or v.find("安值中盛量化1号A期") != -1 or v.find(
                        "安值中盛量化1号私募") != -1:
                    anv_col = FileProcess.get_xlscol(path + attname[i], "Sheet1", 5, "市值")
                    anv_row = FileProcess.get_xlsrow(path + attname[i], "Sheet1", 1, "资产净值")
                    anv = FileProcess.get_xlsnetval(path + attname[i], "Sheet1", anv_row + 1, anv_col + 2)
                else:
                    anv_col = FileProcess.get_xlscol(path + attname[i], "Sheet1", 5, "市值-本币")
                    anv_row = FileProcess.get_xlsrow(path + attname[i], "Sheet1", 1, "资产净值")
                    anv = FileProcess.get_xlsnetval(path + attname[i], "Sheet1", anv_row + 1, anv_col + 1)

                # 处理
                # Haoyu: 这一段是改过的
                if v.find("安值福慧量化2号") != -1 or v.find("安值福慧量化3号") != -1:
                    net_val_prefix_removed = net_val.split(':', 1)[1]
                    net_val = net_val_prefix_removed.split('(', 1)[0].strip()
                else:
                    net_val = net_val[FileProcess.first_digpos(net_val):]
                # Haoyu:以上是改动过的一段。

                eval_date = eval_date[FileProcess.first_digpos(eval_date):]
                records[v] = (eval_date, str(net_val), str(cum_netval), str(anv))
                print(v, " ", eval_date, "  ", net_val, "  ", cum_netval, "  ", anv)
                txt_line = v + " " + eval_date + "  " + str(net_val) + "  " + str(cum_netval) + "  " + str(anv)
                result.write(txt_line + "\n")

            elif v in pdt_list_to_be_fixed and attname[i].find('txt')==-1:
                cumnv_row2 = FileProcess.get_xlsrow(path + attname[i], "Sheet1", 1, "累计单位净值")
                nv_col2 = FileProcess.get_xlscol(path + attname[i], "Sheet1", 4, "单位净值")
                ed_col2 = FileProcess.get_xlscol(path + attname[i], "Sheet1", 4, "日期")
                # 取三个值（这个会带上格子里的其他字符串，不是纯数字）
                net_val2 = FileProcess.get_xlsnetval(path + attname[i], "Sheet1", 4, nv_col2 + 1)
                eval_date2 = FileProcess.get_xlsnetval(path + attname[i], "Sheet1", 4, ed_col2 + 1)
                cum_netval2 = FileProcess.get_xlsnetval(path + attname[i], "Sheet1", cumnv_row2 + 1, 2)

                anv_col2 = FileProcess.get_xlscol(path + attname[i], "Sheet1", 5, "市值")
                anv_row2 = FileProcess.get_xlsrow(path + attname[i], "Sheet1", 1, "资产净值")
                anv2 = FileProcess.get_xlsnetval(path + attname[i], "Sheet1", anv_row2 + 1, anv_col2 + 1)
                net_val2 = net_val2[FileProcess.first_digpos(net_val2):]
                eval_date2 = eval_date2[FileProcess.first_digpos(eval_date2):]
                records[v] = (eval_date2, str(net_val2), str(cum_netval2), str(anv2))
                print(v, " ", eval_date2, "  ", net_val2, "  ", cum_netval2, "  ", anv2)
                txt_line = v + " " + eval_date2 + "  " + str(net_val2) + "  " + str(cum_netval2) + "  " + str(anv2)
                result.write(txt_line + "\n")




            elif v in pdt_list_2 and (attname[i].find("txt") == -1):
                # 找到累计单位净值、单位净值和估值日期的行列
                cumnv_row = FileProcess.get_xlsrow(path + attname[i], "Sheet1", 1, "累计单位净值")
                nv_col = FileProcess.get_xlscol(path + attname[i], "Sheet1", 3, "单位净值")
                ed_col = FileProcess.get_xlscol(path + attname[i], "Sheet1", 3, "日期")
                # 取三个值
                net_val = FileProcess.get_xlsnetval(path + attname[i], "Sheet1", 3, nv_col + 1)
                eval_date = FileProcess.get_xlsnetval(path + attname[i], "Sheet1", 3, ed_col + 1)
                cum_netval = FileProcess.get_xlsnetval(path + attname[i], "Sheet1", cumnv_row + 1, 2)
                # 找到总资产净值
                anv_col = FileProcess.get_xlscol(path + attname[i], "Sheet1", 4, "市值")
                anv_row = FileProcess.get_xlsrow(path + attname[i], "Sheet1", 1, "基金资产净值")
                anv = FileProcess.get_xlsnetval(path + attname[i], "Sheet1", anv_row + 1, anv_col + 1)
                # 处理
                net_val = net_val[FileProcess.first_digpos(net_val):]
                eval_date = eval_date[FileProcess.first_digpos(eval_date):]
                records[v] = (eval_date, str(net_val), str(cum_netval), str(anv))
                print(v, " ", eval_date, "  ", net_val, "  ", cum_netval, "  ", anv)
                txt_line = v + " " + eval_date + "  " + str(net_val) + "  " + str(cum_netval) + "  " + str(anv)
                result.write(txt_line + "\n")

            elif v in pdt_list_3 and (attname[i].find("txt") == -1):
                # 找到累计单位净值、单位净值和估值日期的行列
                cumnv_row = FileProcess.get_xlsrow(path + attname[i], "Sheet1", 2, "累计单位净值")
                nv_col = FileProcess.get_xlscol(path + attname[i], "Sheet1", 3, "单位净值")
                ed_col = FileProcess.get_xlscol(path + attname[i], "Sheet1", 3, "日期")
                # 取三个值
                net_val = FileProcess.get_xlsnetval(path + attname[i], "Sheet1", 3, nv_col + 1)
                eval_date = FileProcess.get_xlsnetval(path + attname[i], "Sheet1", 3, ed_col + 1)
                cum_netval = FileProcess.get_xlsnetval(path + attname[i], "Sheet1", cumnv_row + 1, 3)
                # 找到总资产净值
                anv_col = FileProcess.get_xlscol(path + attname[i], "Sheet1", 4, "市值")
                anv_row = FileProcess.get_xlsrow(path + attname[i], "Sheet1", 2, "基金资产净值")
                anv = FileProcess.get_xlsnetval(path + attname[i], "Sheet1", anv_row + 1, anv_col + 1)
                # 处理
                net_val = net_val[FileProcess.first_digpos(net_val):]
                eval_date = eval_date[FileProcess.first_digpos(eval_date):]
                records[v] = (eval_date, str(net_val), str(cum_netval), str(anv))
                print(v, " ", eval_date, "  ", net_val, "  ", cum_netval, "  ", anv)
                txt_line = v + " " + eval_date + "  " + str(net_val) + "  " + str(cum_netval) + "  " + str(anv)
                result.write(txt_line + "\n")
            # Haoyu:上面合并了海天3号和海天7号

            # Haoyu:给兴建1号加的
            elif (v.find("兴建量化1号") != -1) and (attname[i].find("txt") == -1):
                # 找到累计单位净值、单位净值和估值日期的行列
                cumnv_row3 = FileProcess.get_xlsrow(path + attname[i], "Sheet1", 1, "累计单位净值")
                nv_col3 = FileProcess.get_xlscol(path + attname[i], "Sheet1", 3, "单位净值")
                ed_col3 = FileProcess.get_xlscol(path + attname[i], "Sheet1", 3, "日期")
                # 取三个值
                net_val3 = FileProcess.get_xlsnetval(path + attname[i], "Sheet1", 3, nv_col3 + 1)
                eval_date3 = FileProcess.get_xlsnetval(path + attname[i], "Sheet1", 3, ed_col3 + 1)

                cum_netval3 = FileProcess.get_xlsnetval(path + attname[i], "Sheet1", cumnv_row3 + 1, 2)
                # 找到总资产净值
                anv_col3 = FileProcess.get_xlscol(path + attname[i], "Sheet1", 4, "市值")
                anv_row3 = FileProcess.get_xlsrow(path + attname[i], "Sheet1", 1, "基金资产净值")
                anv3 = FileProcess.get_xlsnetval(path + attname[i], "Sheet1", anv_row3 + 1, anv_col3 + 1)
                # 处理
                net_val3 = net_val3[FileProcess.first_digpos(net_val3):]
                eval_date_unformatted = eval_date3[FileProcess.first_digpos(eval_date3):]
                # 需要改一下日期的格式, 比其他多一个行以下代码
                eval_date3 = eval_date_unformatted[:4] + "-" + eval_date_unformatted[4:6] + "-" + eval_date_unformatted[
                                                                                                 6:]

                records[v] = (eval_date3, str(net_val3), str(cum_netval3), str(anv3))
                # 名字和安值产品净值.xlsm里面对应好
                print(v, " ", eval_date3, "  ", net_val3, "  ", cum_netval3, "  ", anv3)
                txt_line = v + " " + eval_date3 + "  " + str(net_val3) + "  " + str(cum_netval3) + "  " + str(anv3)
                result.write(txt_line + "\n")

''' #以下都是清算了的基金，不需要再统计
            elif v.find("启林") != -1 and (attname[i].find("txt") != -1):
                eval_date, net_val, cum_netval = FileProcess.get_txtnetval1(path + attname[i])
                eval_date = eval_date[:4] + "-" + eval_date[4:6] + "-" + eval_date[6:8]
                records[v] = (eval_date, str(net_val), str(cum_netval), "")
                print(v, " ", eval_date, "  ", net_val, "  ", cum_netval)
                txt_line = v + " " + eval_date + "  " + str(net_val) + "  " + str(cum_netval)
                result.write(txt_line + "\n")

            elif (v.find("九坤") != -1 and attname[i].find("基金净值表现估算") != -1) and attname[i].find("txt") != -1:
                eval_date, net_val, cum_netval = FileProcess.get_txtnetval(path + attname[i])
                records[v] = (eval_date, str(net_val), str(cum_netval), "")
                print(v, " ", eval_date, "  ", net_val, "  ", cum_netval)
                txt_line = v + " " + eval_date + "  " + str(net_val) + "  " + str(cum_netval)
                result.write(txt_line + "\n")

            elif v.find("茂源") != -1:
                cumnv_col = FileProcess.get_xlscol(path + attname[i], "Sheet1", 6, "累计净值")
                nv_col = FileProcess.get_xlscol(path + attname[i], "Sheet1", 6, "单位净值")
                ed_col = FileProcess.get_xlscol(path + attname[i], "Sheet1", 3, "日期")
                # 取三个值
                net_val = FileProcess.get_xlsnetval(path + attname[i], "Sheet1", 7, nv_col + 1)
                eval_date = FileProcess.get_xlsnetval(path + attname[i], "Sheet1", 3, ed_col + 1)
                cum_netval = FileProcess.get_xlsnetval(path + attname[i], "Sheet1", 7, cumnv_col + 1)
                # 找到总资产净值
                anv_col = FileProcess.get_xlscol(path + attname[i], "Sheet1", 6, "资产净值")
                anv = FileProcess.get_xlsnetval(path + attname[i], "Sheet1", 7, anv_col + 2)  # 存在合并的单元格
                # 处理
                net_val = net_val[FileProcess.first_digpos(net_val):]
                eval_date = eval_date[FileProcess.first_digpos(eval_date):]
                records[v] = (eval_date, str(net_val), str(cum_netval), str(anv))
                print(v, " ", eval_date, "  ", net_val, "  ", cum_netval, "  ", anv)
                txt_line = v + " " + eval_date + "  " + str(net_val) + "  " + str(cum_netval) + "  " + str(anv)
                result.write(txt_line + "\n")
'''
result.close()
print("净值提取已完成！汇总文件为：", result_path)
oprand = input("按回车退出程序")
