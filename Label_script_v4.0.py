import xlrd
import datetime
import pytz


def get_time():
    local_time = datetime.datetime.now(pytz.timezone('Asia/Shanghai')).strftime("%Y-%m-%d %H:%M:%S")
    return local_time

def action_trans(path):
    action_labels = []
    wb = xlrd.open_workbook(path)
    sheet = wb.sheet_by_name('填写')
    start_times = sheet.col_values(0)
    for index,start_time in enumerate(start_times):
        if index != 0:
            value_list = sheet.row_values(index)
            Label_dict = {}
            Label_dict['id'] = int(value_list[2])
            Label_dict['start'] = round(value_list[0]*24*3600)
            Label_dict['end'] = round(value_list[1]*24*3600)
            Label_dict['label'] = value_list[3]
            Label_dict['index'] = index
            Label_dict['updated_at'] = get_time()
            action_labels.append(Label_dict)
    return action_labels

def review(path):
    error_dict = {}
    error_list = []
    suffix = path.split('.')[-1]  # 获取文件后缀
    prefix = path.split('.')[0].split('\\')[-1]  # 获取文件前缀,windows环境下是'\\',linux环境下是'/'
    # print(prefix)
    review = {}

    # 判断文件后缀是否正确，错误则返回
    if 'suffix' not in error_dict.keys():
        error_dict['suffix'] = []
    if 'prefix' not in error_dict.keys():
        error_dict['prefix'] = []
    if suffix not in ['xlsx','xls']:
        error_dict['suffix'].append('文件后缀错误')


    # 判断文件前缀是否正确。错误则返回
    if prefix in ['left_actions','right_actions','right_tools','left_tools','phases','special_events']:
        error_dict['prefix'] = []

        # 通过文件前缀返回不同的标签字典
        if prefix in ['left_actions','right_actions']:
            review = ACTION_DICT
        if prefix in  ['right_tools','left_tools']:
            review = TOOL_DICT
        if prefix == 'phases':
            review = PHASE_DICT
        if prefix == 'special_events':
            review = SPECIAL_EVENT_DICT
    else:
        error_dict['prefix'].append('文件前缀错误')

    # print('error_dict',error_dict)
    # 判断内容格式是否正确，前置条件：文件前缀和文件格式正确
    if not error_dict['prefix'] and not error_dict['suffix']:
    # if error_dict['prefix'] == [] and error_dict['suffix'] == []:
        wb = xlrd.open_workbook(path)
        sheet = wb.sheet_by_name('填写')
        start_times = sheet.col_values(0)
        max_time = 0.125
        for index,start in enumerate(start_times):
            if index != 0:
                data = sheet.row_values(index)
                print(data)
                # 检查是否有空字符串
                if '' in data:
                    error_list.append(f'第{index+1}行有空字符，请检查')
                    return error_list

                # 检查时间格式是否正确
                if type(data[0]) != float or type(data[1]) != float:
                    error_list.append(f'第{index+1}行时间格式错误，请检查')
                    return error_list

                if data[0] > max_time or data[1] > max_time:
                    error_list.append(f'第{index+1}行时间范围错误，请检查')
                    return error_list

                if data[1] < data[0]:
                    error_list.append(f'第{index+1}行时间范围错误，请检查')
                    return error_list

                if data[2] not in review.keys():
                    error_list.append(f'第{index+1}行id错误，请检查')
                    return error_list
                elif data[3] != review[data[2]]:
                    error_list.append(f'第{index+1}行label列错误，请检查')
                    return error_list
    else:
        error_list.append(error_dict)


    return error_list

if __name__ == '__main__':
    ACTION_DICT = {
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
    SPECIAL_EVENT_DICT = {
        1: "夹胆囊管",
        2: "夹胆囊动脉",
        3: "离断胆囊管",
        4: "离断胆囊动脉",
    }
    PHASE_DICT = {
        1: "抓取胆囊",
        2: "建立气腹",
        3: "分离粘连",
        4: "游离胆囊三角",
        5: "分离胆囊床",
        6: "清理术野"
    }
    TOOL_DICT = {
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

    # 以下均为windows环境下excel路径，若需要在linux环境下测试，请修改review()函数中
    # prefix = path.split('.')[0].split('\\')[-1] 修改为 prefix = path.split('.')[0].split('/')[-1]

    # excel_path = r'E:\withai_document\视频管理\LC10000\手术标签脚本\right_actions.xlsx'
    # excel_path = r'E:\withai_document\视频管理\LC10000\手术标签脚本\excel_model-v2\phases.xlsx'
    # excel_path = r'E:\withai_document\视频管理\LC10000\手术标签脚本\excel_model-v2\right_tools.xlsx'
    excel_path = r'E:\withai_document\视频管理\LC10000\手术标签脚本\test_for_label\right_actions.xlsx'

    check_message = review(path=excel_path)
    print(check_message)
    # 判断返回的列表是否为空，为空则执行格式转换
    if not check_message:
        json_list = action_trans(excel_path)
        print(json_list)




