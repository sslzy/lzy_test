import os
import json
import black
import shutil
import xlrd
import xlsxwriter
video_dict = {"LC-HX-0004140719": 29,
              "LC-HX-0009048674": 223,
              "LC-HX-0033992219": 872,
              }
video_durantion = {
    "LC-HX-0004140719": "00:34:57",
    "LC-HX-0009048674": "00:22:22",
    "LC-HX-0033992219": "00:35:10",
}

# 将标签文件分类存放
# user_id = 37
# left_action_example = r'Y:\LC10000\labels\category_id=3\type_id=3\label_version=1.0\pvid=28\user_id=37\left_actions.json'
# right_action_example = r'Y:\LC10000\labels\category_id=7\type_id=3\label_version=1.0\pvid=28\user_id=37\right_actions.json'
#
# json_save = r'E:\biaozhu\special_envent\20210714_CZX\json_save'
# left_save_path = os.path.join(json_save, 'left')
# right_save_pth = os.path.join(json_save, 'right')
# if not os.path.exists(left_save_path):
#     os.makedirs(left_save_path)
#
# if not os.path.exists(right_save_pth):
#     os.makedirs(right_save_pth)
# for k, v in enumerate(video_dict):
#
#     left_ori = os.path.join(r'Y:\LC10000\labels\category_id=3\type_id=3\label_version=1.0','pvid='+str(video_dict[v])+'\\'+'user_id='+str(user_id)+'\\'+'left_actions.json')
#     right_ori = os.path.join(r'Y:\LC10000\labels\category_id=7\type_id=3\label_version=1.0','pvid=' + str(video_dict[v]) + '\\' + 'user_id=' +str(user_id)+'\\''right_actions.json')
#
#     shutil.copy(left_ori,os.path.join(left_save_path,v+'.json'))
#     shutil.copy(right_ori,os.path.join(right_save_pth,v+'.json'))

LABEL_CATEGORY = (
    (1, "器械"),
    (2, "器官"),
    (3, "左手动作"),
    (4, "阶段"),
    (5, "CVS"),
    (6, "事件"),
    (7, "右手动作"),
    (8, "左手器械"),
    (9, "右手器械"),
)

PHASE_CONTENT = {
    1: "抓取胆囊",
    2: "建立气腹",
    3: "分离粘连",
    4: "游离胆囊三角",
    5: "分离胆囊床",
    6: "清理术野"
}
Phase = {
    1: "EG",
    2: "EA",
    3: "AL",
    4: "MCT",
    5: "DGB",
    6: "COR"
}


special_dict = {
    2: 'r-hook-action',
}


actions_dict = {
    1: "钩",
    2: "电勾激发",
    3: "无效钩",
    4: "推",
    5: "刮",
    6: "抓",
    7: "凝固",
    8: "剪",
    9: "钩解剖",
    10: "夹",
    11: "钝性剥离",
    12: "擦",
    13: "吸",
    14: "压塞",
    15: "无效抓",
    16: "无效夹",
    17: "穿刺",
}

# 提取标签文件中的时间段
json_save = r'E:\biaozhu\special_envent\20210714_CZX\json_save'
left_save_path = os.path.join(json_save, 'left')
right_save_path = os.path.join(json_save, 'right')
phase_save_path = r'E:\biaozhu\special_envent\20210714_CZX\json_save\phases'
# 读取原始excel
path = r'E:\biaozhu\special_envent\20210714_CZX\动作元原始数据.xlsx'
wb = xlrd.open_workbook(path)
sheet = wb.sheet_by_index(0)
row_list = sheet.row_values(0)
# print(row_list)

# 读取json文件的内容
for k,v in enumerate(video_dict):
    # 读取左手动作
    left_fin = open(os.path.join(left_save_path,v+'.json'), encoding='UTF-8')
    left_dict = json.load(left_fin)

    # 读取右手动作
    right_fin = open(os.path.join(right_save_path,v+'.json'), encoding='UTF-8')
    right_dict = json.load(right_fin)

    # 读取不同阶段的标注时长
    phase_fin = open(os.path.join(phase_save_path,v+'.json'),encoding='UTF-8')
    phase_dict = json.load(phase_fin)

    # print(left_dict)
    # print(right_dict)

    left_data = {}
    right_data = {}
    for l_data in left_dict:
        for phases in phase_dict:
            phase_start = phases['start']
            phase_end = phases['end']
            if phase_start <= l_data['end'] <= phase_end:
                l_data['phase_name'] = Phase[phases['id']]

    for r_data in right_dict:
        for phases in phase_dict:
            phase_start = phases['start']
            phase_end = phases['end']
            if phase_start <= r_data['end'] <= phase_end:
                r_data['phase_name'] = Phase[phases['id']]
    print(right_data)
    print(left_data)

    for left_index,left_key in enumerate(left_dict):
        start = left_key['start']
        end = left_key['end']
        time = end - start

        if left_key['id'] not in left_data.keys():
            left_data[left_key['id']] = {}
        if 'time' and 'count' and left_dict['phase_name'] not in left_data[left_key['id']].keys():
            left_data[left_key['id']]['time'] = time
            left_data[left_key['id']]['count'] = 1
        else:
            left_data[left_key['id']]['time'] += time
            left_data[left_key['id']]['count'] += 1

        # for phases in phase_dict:
        #     phase_start = phases['start']
        #     phase_end = phases['end']
        #     if phase_start <= end <= phase_end:
        #         left_data[left_key['id']]['phase'] = Phase[phases['id']]

    for right_index,right_key in enumerate(right_dict):
        start = right_key['start']
        end = right_key['end']
        time = end - start

        if right_key['id'] not in right_data.keys():
            right_data[right_key['id']] = {}
        if 'time' and 'count' and right_dict['phase_name'] not in right_data[right_key['id']].keys():
            right_data[right_key['id']]['time'] = time
            right_data[right_key['id']]['count'] = 1
        else:
            right_data[right_key['id']]['time'] += time
            right_data[right_key['id']]['count'] += 1

        # for phases in phase_dict:
        #     phase_start = phases['start']
        #     phase_end = phases['end']
        #     if phase_start <= end <= phase_end:
        #         right_data[right_key['id']]['phase'] = Phase[phases['id']]

    print(left_data)
    print(right_data)
    print(v)

    # 写入excel
    # 左手动作写入sheet表
    sheet_nameL = v+'_L'
    sheet_nameR = v+'_R'
    excel_save = r'E:\biaozhu\special_envent\20210714_CZX\excel_save'+'\\'+v+'.xlsx'
    print(excel_save)
    workbook = xlsxwriter.Workbook(excel_save)  # 创建一个名为‘workbook_path[i].xlsx’的excel文件
    worksheet = workbook.add_worksheet(sheet_nameL)  # 创建一个工作表对象
    L = 0
    # worksheet.set_column('A:A', 20)#设置第一列（A）的宽度为20px
    for i in left_data.keys():
        worksheet.write(0, 0 , v)
        worksheet.write(0, L + 1,'l-'+actions_dict[i]+'_action')
        worksheet.write(0, L + 2, 'l-'+actions_dict[i]+'_time')
        worksheet.write(1, 0 , v)
        worksheet.write(1, L + 1,str(left_data[i]['count']))
        worksheet.write(1, L + 2, str(left_data[i]['time']))
        L += 2


    worksheetR = workbook.add_worksheet(sheet_nameR)
    R = 0
    for y in right_data.keys():
        worksheetR.write(0, 0 , v)
        worksheetR.write(0, R + 1,'r-'+actions_dict[y]+'_action')
        worksheetR.write(0, R + 2, 'r-'+actions_dict[y]+'_time')
        worksheetR.write(1, 0 , v)
        worksheetR.write(1, R + 1,str(right_data[y]['count']))
        worksheetR.write(1, R + 2, str(right_data[y]['time']))
        R += 2
    workbook.close()
