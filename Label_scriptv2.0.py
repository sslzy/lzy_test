import xlrd
import json
import datetime
import pytz
import copy

def get_time():
    local_time = datetime.datetime.now(pytz.timezone('Asia/Shanghai')).strftime("%Y-%m-%d %H:%M:%S")
    return local_time

def ActionTrans(path):
    global actions_dict
    global special_events_dict
    global tools_dict
    global phases_dict
    ActionLabel = []
    wb = xlrd.open_workbook(path)
    sheet = wb.sheet_by_index(0)
    TimeStart = sheet.col_values(0)
    TimeEnd = sheet.col_values(1)
    Id = sheet.col_values(2)
    Label = sheet.col_values(3)
    for i in range(1,len(TimeStart)):
        if Id[i] == '':
            pass
        else:
            Label_dict = {}
            Label_dict['id'] = int(Id[i])
            Label_dict['start'] = round(TimeStart[i]*24*3600)
            Label_dict['end'] = round(TimeEnd[i]*24*3600)
            Label_dict['label'] = Label[i]
            Label_dict['index'] = i
            Label_dict['updated_at'] = get_time()
            ActionLabel.append(Label_dict)
    return ActionLabel

def Review(path):
    global actions_dict
    global special_events_dict
    global tools_dict
    global phases_dict
    Path = path
    error_dict = {}
    Suffix = Path.split('.')[-1]  # 获取文件后缀
    Prefix = path.split('.')[0].split('\\')[-1]  # 获取文件前缀
    review = {}
    # 判断文件后缀是否正确，错误则返回
    if 'Suffix' not in error_dict.keys():
        error_dict['Suffix'] = []
    if 'Prefix' not in error_dict.keys():
        error_dict['Prefix'] = []
    if Suffix not in ['xlsx','xls']:
        error_dict['suffix'].append('文件后缀错误')

    print(Prefix)
    # 判断文件前缀是否正确。错误则返回
    if Prefix in ['left_actions','right_actions','right_tools','left_tools','phases','special_events']:
        error_dict['Prefix'] = []
        # 通过文件前缀返回不同的标签字典
        if Prefix in ['left_actions','right_actions']:
            review = actions_dict
        if Prefix in  ['right_tools','left_tools']:
            review = tools_dict
        if Prefix == 'phases':
            review = phases_dict
        if Prefix == 'special_events':
            review = special_events_dict
    else:
        error_dict['Prefix'].append('文件前缀错误')

    # 判断内容格式是否正确，前置条件：文件前缀和文件格式正确
    if error_dict['Prefix'] == [] and error_dict['Suffix'] == []:

        wb = xlrd.open_workbook(path)
        sheet = wb.sheet_by_name('填写')
        TimeStart = sheet.col_values(0)
        TimeEnd = sheet.col_values(1)
        Id = sheet.col_values(2)
        Label = sheet.col_values(3)

        for i in range(1,len(TimeStart)):
            # 检查时间格式是否正确
            if i+1 < len(TimeStart):
                # 检查TimeStart的每一行的是否为空
                # if TimeStart[i] == '' and TimeStart[i+1] and TimeStart[i-1] != '':
                if TimeStart[i] == '':
                    if 'time_start_empty' not in error_dict.keys():
                        error_dict['time_start_empty'] = []
                    else:
                        error_dict['time_start_empty'].append(f'start第{i+1}行中有空字符')

                # 检查TimeStart的时间格式是否正确
                if 'time_start_type' not in error_dict.keys():
                    error_dict['time_start_type'] = []
                if type(TimeStart[i]) != float or TimeStart[i] >= 0.125:
                    # print(i)
                    # print(TimeStart[i])
                    error_dict['time_start_type'].append(f'start第{i+1}行开始时间格式存在错误')

                # 检查TimeEnd的每一行的是否为空
                if TimeEnd[i] == '':
                    if 'time_end_empty' not in error_dict.keys():
                        error_dict['time_end_empty'] = []
                    error_dict['time_end_empty'].append(f'end第{i+1}行中有空字符')
                # if TimeEnd[i] == '' and TimeEnd[i+1] or TimeEnd[i-1] != '':

                # 检查TimeStart的时间格式是否正确
                if 'time_end_type' not in error_dict.keys():
                    error_dict['time_end_type'] = []
                if type(TimeEnd[i]) != float or TimeEnd[i] >= 0.125:
                    error_dict['time_end_type'].append(f'end第{i+1}行的时间格式存在错误')

                # 检查id列是否存在空字符

                if Id[i] == '' and Id[i + 1] == '' or Id[i - 1] == '':
                    if 'Id_empty' not in error_dict.keys():
                        error_dict['Id_empty'] = []
                    else:
                        error_dict['Id_empty'].append(f'id第{i+1}行中有空字符')

                # 检查id列的格式是否正确
                if Id[i] != '' :
                    if 'id_values' not in error_dict.keys():
                        error_dict['id_values'] = []
                    if int(Id[i]) not in review.keys():
                        error_dict['id_values'].append(f'id第{i+1}行的序号错误')
                if type(Id[i]) != float:
                    if 'Id_type' not in error_dict.keys():
                        error_dict['Id_type'] = []
                    else:
                        error_dict['Id_type'].append(f'id第{i+1}行格式错误')

                # 检查label列格式与内容是否正确
                if Label[i] == '' and Label[i + 1] or Label[i - 1] == '':
                    if 'Label_empty' not in error_dict.keys():
                        error_dict['Label_empty'] = []
                    error_dict['Label_empty'].append(f'id第{i+1}行中有空字符')

                if Label[i] not in review.values():
                    if 'Label_values' not in error_dict.keys():
                        error_dict['Label_values'] = []
                    error_dict['Label_values'].append(f'label第{i+1}行的内容错误')
    return error_dict


if __name__ == '__main__':
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
    special_events_dict = {
        1: "夹胆囊管",
        2: "夹胆囊动脉",
        3: "离断胆囊管",
        4: "离断胆囊动脉",
    }
    phases_dict = {
        1: "抓取胆囊",
        2: "建立气腹",
        3: "分离粘连",
        4: "游离胆囊三角",
        5: "分离胆囊床",
        6: "清理术野"
    }
    tools_dict = {
        1: "戳卡",
        2: "无损伤抓钳",
        3: "电勾",
        4: "施夹器",
        5: "可吸收夹",
        6: "金属夹",
        7: "马里兰钳",
        8: "纱布",
        9: "直分离钳",
        10: "剪刀",
        11: "大抓钳",
        12: "分离钳",
        13: "标本袋",
        14: "穿刺针",
        15: "吸引器",
        16: "电凝",
        17: "引流管",
        18: "波浪钳",
    }

    excel_path = r'E:\withai_document\视频管理\LC10000\手术标签脚本\right_actions.xlsx'
    # excel_path = r'E:\withai_document\视频管理\LC10000\手术标签脚本\test_for_label\right_actions.xlsx'
    check_message = Review(path=excel_path)
    print(check_message)
    count = len(check_message.values())
    # 根据check_message的内容，判断是否进入函数ActionTrans，
    for check in check_message.values():
        if len(check) > 0:
            count += 1
        else:
            pass
    if count == len(check_message.values()):  # 判断文件格式及内容是否错误，没有错误则开始转换文件格式
        json_list = ActionTrans(excel_path)
        print(json_list)



