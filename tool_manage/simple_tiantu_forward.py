import requests
import re
import time
import openpyxl
from datetime import datetime
from keyboard import write, press_and_release
import uiautomation as auto
from pyperclip import copy
from xlrd import open_workbook
from PIL import Image
import pytesseract, os, pyautogui
import subprocess
import tkinter as tk

from Crypto.Cipher import AES
from Crypto.Util.Padding import pad
from base64 import b64encode

"""
天图美线企微群发

导入第三方库
pip install requests openpyxl keyboard uiautomation pyperclip xlrd==1.1.0 Image pytesseract pyautogui numpy opencv-python crypto pycryptodome
"""


# 天图密码加密方式
# tip: 如果出现找不到Crypto的情况，可以去文件找到把文件名crypto改成Crypto
def encrypt(word: str) -> str:
    DEFAULT_KEY = b'abcdefgabcdefg12'  # 确保这是 16, 24 或 32 字节长的字节串
    DEFAULT_IV = b'abcdefgabcdefg12'  # 确保这是 16 字节长的字节串
    # 创建 Cipher 对象并加密数据
    cipher = AES.new(DEFAULT_KEY, AES.MODE_CBC, DEFAULT_IV)
    encrypted_bytes = cipher.encrypt(pad(word.encode('utf-8'), AES.block_size))

    # 返回 Base64 编码的字符串
    return b64encode(encrypted_bytes).decode('utf-8')


# 登录天图
def get_token(base_url, user, pwd):
    url = base_url + "/api/sys/auth/login"
    password = encrypt(pwd)
    headers = {"content-type": "application/json; charset=utf-8",
               "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/135.0.0.0 Safari/537.36"}
    data = {"password": password,
            "username": user}
    response = requests.post(url, headers=headers, json=data)
    token = response.json()["data"]["token"]
    return token


# 运单数据导出
def get_data(base_url, token, delivery_no_list=""):
    # FIXME 增加实时时间，不然后续无法查询数据
    now_mouth = int(datetime.now().strftime("%m"))
    now_year = int(datetime.now().strftime("%Y"))
    if now_mouth > 2:
        begin_now_time = datetime.now().strftime(f"%Y-{now_mouth - 1}-%d %H:%M:%S")
    else:
        begin_now_time = datetime.now().strftime(f"{now_year - 1}-12-%d %H:%M:%S")
    last_now_time = datetime.now().strftime("%Y-%m-%d 23:59:59")
    url = base_url + "/api/fba-shipment/export_order"
    headers = {"content-type": "application/json", "token": token}
    data = {
        "code_type": "",
        "code_list": "",
        "customer_id_list": [],
        "finance_user_id_list": [],
        "customer_type": "",
        "service_id_list": [],
        "warehouse_code_list": [],
        "lade_number_list": "",
        "delivery_no_list": [delivery_no_list],
        "receiver_country_code_list": [],
        "declare_type_list": [],
        "customs_type_list": [],
        "tax_type_list": [],
        "freight_station_list": [],
        "attr_id_list": [],
        "assistant_user_id_list": [],
        "sale_user_id_list": [],
        "exception_list": [],
        "no_exception_list": [],
        "sale_user_name": "",
        "remark": "",
        "inner_remark": "",
        "marks": [],
        "no_marks": [],
        "vat_number": "",
        "po_number": "",
        "ltd_id_list": [],
        "route_ids": [],
        "create_time": [
            begin_now_time,
            last_now_time
        ],
        "picking_time": [],
        "sign_time": [],
        "return_time": [],
        "cancel_time": [],
        "work_order_type": "",
        "bloc_no_list": "",
        "picking_freight_station_list": [],
        "page": 1,
        "limit": 100,
        "create_start_time": "2025-02-07 00:00:00",
        "create_end_time": "2025-05-08 23:59:59"
    }
    response = requests.post(url, headers=headers, json=data)
    if response.status_code == 200:
        return
    else:
        raise Exception(response.status_code, response.text)


# 获取导出模板列表
def get_export_template_id(base_url, token):
    url = base_url + '/api/file/export/findFastExportFile'
    headers = {"content-type": "application/json", "token": token}
    response = requests.get(url, headers=headers)
    for i in range(len(response.json()['data'])):
        template_name = response.json()['data'][i]['templateName']
        if template_name == '运单导出':
            url = response.json()['data'][i]['fileUrl']
            return url


# 下载文件
def download_file(url, file_path, file_name):
    print(file_path + file_name)
    if not os.path.exists(file_path + file_name):
        response = requests.get(url)
        with open(file_path + file_name, 'wb') as f:
            f.write(response.content)


# 获取最大行数
def get_max_row(ws):
    i = ws.max_row
    row_count = 0
    while ws.max_row > 0:
        row_dict = {i.value for i in ws[i]}
        if row_dict == {None}:
            i = i - 1
        else:
            row_count = i
            break
    return row_count


# 读取excel 获取运单号
def get_shipment_num_list(excel_path, sheet):
    wb = openpyxl.load_workbook(excel_path)
    ws = wb[sheet]
    shipment_num_list = []
    row_count = get_max_row(ws)
    for i in range(2, row_count + 1):
        shipment_num_list.append(ws.cell(row=i, column=1).value)
    return shipment_num_list


# 过滤excel重复值
def skip_all_duplicates(arr):
    seen = []
    for value in arr:
        if value in seen:
            continue  # 如果当前值已经在集合中，则跳过
        seen.append(value)
    return seen


# 获取图像在屏幕中的位置
def get_photo_position(image_path):
    try:
        image = pyautogui.locateOnScreen(image_path, grayscale=True)
        center_x, center_y = pyautogui.center(image)
        return center_x, center_y
    except Exception as e:
        print(e, image_path)


# 复制文本信息
def copy_text_file_to_clipboard(file_path):
    with open(file_path, "r", encoding='utf8') as file:
        text = file.read()
        if text is not None:
            copy(text)


# 复制文件
def copy_file_to_clipboard(file_path):
    with open(file_path, "rb") as file:
        image = file.read()
        copy(image)


# 客户名称匹配群名称处理
def sales_match_group():
    deal_path = "D:\\group_send\\tiantu_deal\\运单导出.xlsx"
    read_wb = openpyxl.load_workbook("D:\\group_send\\tiantu_deal\\美线客服群发-基础信息维护.xlsx")
    read_ws = read_wb['群名称']
    # 获取基础信息表最大行号
    row_count = get_max_row(read_ws)
    read_wb.close()
    org_wb = openpyxl.load_workbook(deal_path)
    org_ws = org_wb['Sheet1']
    # 获取运单表最大行号
    org_row_count = get_max_row(org_ws)
    first_deal_ws = org_wb.create_sheet('Sheet2')
    head_row = ['客户名称', '运单号', '客户单号', '群聊名称']
    for i in range(len(head_row)):
        first_deal_ws.cell(row=1, column=i + 1, value=head_row[i])
    # 客户名称和群名称匹配
    for i in range(2, org_row_count + 1):
        customer_name = org_ws.cell(row=i, column=5).value
        for j in range(1, row_count + 1):
            if customer_name == str(read_ws.cell(row=j, column=1).value):
                first_deal_ws.cell(row=i, column=1, value=str(org_ws.cell(row=i, column=5).value))
                first_deal_ws.cell(row=i, column=2, value=str(org_ws.cell(row=i, column=1).value))
                first_deal_ws.cell(row=i, column=3, value=str(org_ws.cell(row=i, column=2).value))
                first_deal_ws.cell(row=i, column=4, value=str(read_ws.cell(row=j, column=2).value))
    org_wb.save(deal_path)
    # 获取唯一的群名称
    simple_sales_name = []
    for i in range(2, org_row_count + 1):
        deal_sales_name = first_deal_ws.cell(row=i, column=4).value
        simple_sales_name.append(deal_sales_name)
    select_data = skip_all_duplicates(simple_sales_name)
    second_deal_ws = org_wb.create_sheet('Sheet3')
    head_row = ['客户名称', '运单号', '客户单号', '群聊名称']
    for i in range(len(head_row)):
        second_deal_ws.cell(row=1, column=i + 1, value=head_row[i])
    k = 1
    # 根据唯一群名进行排序
    for i in range(len(select_data)):
        for j in range(2, org_row_count + 1):
            group_name = first_deal_ws.cell(row=j, column=4).value
            if select_data[i] == group_name and group_name is not None and group_name != '':
                k += 1
                second_deal_ws.cell(row=k, column=1, value=first_deal_ws.cell(row=j, column=1).value)
                second_deal_ws.cell(row=k, column=2, value=first_deal_ws.cell(row=j, column=2).value)
                second_deal_ws.cell(row=k, column=3, value=first_deal_ws.cell(row=j, column=3).value)
                second_deal_ws.cell(row=k, column=4, value=first_deal_ws.cell(row=j, column=4).value)
    # 整理发送群
    # FIXME 处理数据为空的情况
    select_data.remove(None)
    third_deal_ws = org_wb.create_sheet('Sheet4')
    head_row = ['群聊名称', '发送结果']
    for i in range(len(head_row)):
        third_deal_ws.cell(row=1, column=i + 1, value=head_row[i])
    for i in range(len(select_data)):
        if select_data[i]:
            third_deal_ws.cell(row=i + 2, column=1, value=select_data[i])
    org_wb.save(deal_path)
    org_wb.close()


# 天图下载运单文件处理
def system_download(base_url, user, pwd, delivery_no_list=""):
    token = get_token(base_url, user, pwd)
    time.sleep(0.5)
    get_data(base_url, token, delivery_no_list)
    time.sleep(3)
    url = get_export_template_id(base_url, token)
    time.sleep(1)
    if str(url).lower().find('None'):
        while True:
            time.sleep(1)
            url = get_export_template_id(base_url, token)
            if 'None' not in str(url):
                break
    download_file(url, 'D:\\group_send\\tiantu_deal\\', '运单导出.xlsx')


# 查询发送人信息
def search_sender(user_name, search_path):
    # 搜索输入框查询
    button_x, button_y = get_photo_position(search_path)
    pyautogui.moveTo(button_x - 60, button_y)
    pyautogui.click()
    write(user_name)
    time.sleep(1)
    pyautogui.moveTo(button_x + 10, button_y + 70)
    pyautogui.click()


# 点击逐条转发按钮
def forward_item(file_path):
    forward_item_x, forward_item_y = get_photo_position(file_path)
    pyautogui.moveTo(forward_item_x, forward_item_y)
    pyautogui.click()


# 根据图片移动鼠标位置
def position_and_move(file_path):
    send_button_x, send_button_y = get_photo_position(file_path)
    pyautogui.moveTo(send_button_x, send_button_y)
    pyautogui.click()


# 发送按钮
def send_button(file_path):
    send_button_x, send_button_y = get_photo_position(file_path)
    # pyautogui.moveTo(send_button_x, send_button_y)
    pyautogui.moveTo(send_button_x - 150, send_button_y)
    pyautogui.click()


# 发送文本信息 new
def new_send_text(file_path, group_name):
    copy_text_file_to_clipboard(file_path)
    time.sleep(0.7)
    press_and_release('ctrl+v')
    time.sleep(0.7)
    text = '\n' + get_customer_and_shipment(group_name)
    copy(text)
    time.sleep(0.7)
    press_and_release('ctrl+v')
    press_and_release('enter')


# 循环文件夹内文件 消息内容.txt 除外
def get_file_name(folder_path):
    files = os.listdir(folder_path)
    file_paths = []
    for file in files:
        if re.search('消息内容', file):
            continue
        file_path = os.path.join(folder_path, file)
        if os.path.isfile(file_path):
            file_paths.append(folder_path + "\\" + file)
    return file_paths


# 发送文件信息
def send_file(file_path):
    if file_path is not None:
        # copy_file_to_clipboard(file_path)
        args = ['powershell', 'Get-Item {} | Set-Clipboard'.format(file_path)]
        subprocess.Popen(args=args, encoding='utf-8')
        time.sleep(0.8)
        press_and_release('ctrl+v')
        time.sleep(0.8)


# 输入信息
def send_point(text):
    write(text)
    press_and_release('enter')


# 处理图片
def deal_picture(image_path, text_data):
    image = Image.open(image_path)
    pytesseract.load_freq_dawg = False
    pytesseract.load_system_dawg = False
    data = pytesseract.image_to_data(image, output_type=pytesseract.Output.DICT)
    for i in range(len(data['text'])):
        if int(float(data['conf'][i])) > 0:
            # if math.ceil(float(data['conf'][i])) > 50:
            if data['text'][i] == text_data:
                x = data['left'][i]
                y = data['top'][i]
                w = data['width'][i]
                h = data['height'][i]
                return text_data, x, y, w, h


# 获取窗口尺寸
def get_window_size():
    size = pyautogui.size()
    return size


# 删除图片文件
def delete_file(image_path):
    os.remove(image_path)
    print(f'Deleted {image_path}')


# 唤起企业微信
def wake_up_work_chat():
    ww_window = auto.WindowControl(Name='企业微信', ClassName='WeWorkWindow')
    ww_window.SetFocus()
    ww_window.Maximize(True)


# 循环文件夹内文件夹
def get_folder(main_path):
    files = os.listdir(main_path)
    folder_paths = []
    for file in files:
        file_path = os.path.join(main_path, file)
        if os.path.isdir(file_path):
            folder_paths.append(file_path)
    # print(folder_paths)
    return folder_paths


# xlrd要调用xlsx，只可以 pip install xlrd==1.1.0
# 获取企业微信用户名 new
def new_get_excel_user(excel_path, sheet):
    workbook = open_workbook(excel_path)
    worksheet = workbook.sheet_by_name(sheet)
    row_count = worksheet.nrows
    user_names = []
    for row_num in range(1, row_count):
        user_name = worksheet.cell_value(row_num, 0)
        if user_name is not None:
            user_names.append(user_name)
    return user_names


# 计算耗时
def count_time(time_str_1, time_str_2):
    # 定义时间格式
    time_format = "%a %b %d %H:%M:%S %Y"

    # 将时间字符串解析为 datetime 对象
    time_obj_1 = datetime.strptime(time_str_1, time_format)
    time_obj_2 = datetime.strptime(time_str_2, time_format)

    # 计算时间差
    time_difference = time_obj_2 - time_obj_1

    # 输出时间差（以秒为单位）
    # print(f"耗时: {time_difference.total_seconds()} 秒")

    # 如果需要更详细的输出（天、小时、分钟、秒）
    days = time_difference.days
    seconds = time_difference.seconds
    hours, remainder = divmod(seconds, 3600)
    minutes, seconds = divmod(remainder, 60)
    return days, hours, minutes, seconds


# 发送结果给群/用户
def send_result_for_me(user_name, forward_count, success_count, start_time, end_time):
    days, hours, minutes, seconds = count_time(start_time, end_time)
    # 搜索输入框查询
    search_x, search_y = get_photo_position('D:\\group_send\\picture_manage\\search_right_button.png')
    pyautogui.moveTo(search_x - 50, search_y)
    pyautogui.click()
    write(user_name)
    press_and_release('enter')
    time.sleep(1)
    pyautogui.moveTo(search_x + 70, search_y + 70)
    pyautogui.click()
    time.sleep(1)
    if forward_count - success_count > 0:
        text = (f'需转发群数量：{forward_count}\n发送成功：{success_count}\n发送失败：{forward_count - success_count}\n'
                f'失败原因：当前企微无法查到该用户,请维护正确的群名称\n耗时详细：{hours}:{minutes}:{seconds}')
    else:
        text = (f'需转发群数量：{forward_count}\n发送成功：{success_count}\n发送失败：{forward_count - success_count}\n'
                f'耗时详细：{hours}:{minutes}:{seconds}')
    copy(text)
    time.sleep(0.5)
    press_and_release('ctrl+v')
    press_and_release('enter')


# 获取客户名称和运单号
def get_customer_and_shipment(group_name):
    deal_path = 'D:\\group_send\\tiantu_deal\\运单导出.xlsx'
    read_wb = openpyxl.load_workbook(deal_path)
    read_ws = read_wb['Sheet3']
    read_row_count = get_max_row(read_ws)
    str_txt = ''
    for i in range(2, read_row_count + 1):
        deal_group_name = read_ws.cell(row=i, column=4).value
        if deal_group_name == group_name:
            shipment_num = read_ws.cell(row=i, column=2).value
            customer_num = read_ws.cell(row=i, column=3).value
            if shipment_num is None:
                shipment_num = ''
            if customer_num is None:
                customer_num = ''
            str_txt = str_txt + str(shipment_num) + ' ' + str(customer_num) + '\n'
    return str_txt


# FIXME 处理了发送的异常情况
# FIXME 1、查询不到数据
# FIXME 2、处理None值的问题

# FIXME 记录 查询不到数据
# 主流程 def main_use_flow(event):
def new_main_use_flow(result_send):
    try:
        start_time = time.asctime(time.localtime(time.time()))
        wake_up_work_chat()
        deal_excel = 'D:\\group_send\\tiantu_deal\\运单导出.xlsx'
        user_names = new_get_excel_user(deal_excel, 'Sheet4')
        wb = openpyxl.load_workbook(deal_excel)
        ws = wb['Sheet4']
        for i in range(0, len(user_names)):
            user_name = user_names[i]
            if user_name == '' or user_name is None:
                continue
            search_sender(user_name, 'D:\\group_send\\picture_manage\\search_right_button.png')
            time.sleep(1.2)
            if get_photo_position("D:\\group_send\\picture_manage\\all_select_box.png") is not None:
                if ws.cell(i + 2, 2).value == '' or ws.cell(i + 2, 2).value is None:
                    ws.cell(row=i + 2, column=2, value='未找到用户或群聊')
                wb.save(deal_excel)
                continue
            # 发送文本信息
            new_send_text('D:\\group_send\\upload_file\\消息内容.txt', user_name)
            files = get_file_name('D:\\group_send\\upload_file')
            # 发送文件信息
            if files != [] or len(files) > 1:
                for file in files:
                    send_file(file)
                time.sleep(0.7)
                press_and_release('enter')
            time.sleep(0.7)
        end_time = time.asctime(time.localtime(time.time()))
        # 发送结果
        forward_count = get_max_row(ws)
        success_count = 0
        for row in range(2, forward_count + 1):
            if ws.cell(row, 2).value == '' or ws.cell(row, 2).value is None:
                success_count += 1
                ws.cell(row=row, column=2, value='已发送')
        wb.save(deal_excel)
        time.sleep(0.5)
        send_result_for_me(result_send, forward_count - 1, success_count, start_time, end_time)
    except Exception as e:
        print(e)


def cs_forward_main(base_url, user, pwd, delivery_no_list, result_send):
    # 设置 Tesseract 的路径(仅在必要时，视你的安装情况而定)
    pytesseract.pytesseract.tesseract_cmd = r'D:\group_send\Tesseract-OCR\tesseract.exe'
    # pyautogui禁用故障保护
    pyautogui.FAILSAFE = False
    # pyautogui增加暂停时间
    pyautogui.PAUSE = 0.2
    # 调用键盘
    subprocess.Popen("DrMain", shell=True)
    time.sleep(1)
    files = get_file_name('D:\\group_send\\tiantu_deal')
    if files != [] or len(files) > 1:
        for file in files:
            if '运单导出' in file.split('\\')[-1]:
                os.remove(file)
    now_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    out_time = "2025-08-09 00:00:00"
    if now_time < out_time:
        if len(delivery_no_list.split(' ')) > 1:
            delivery_no_list = delivery_no_list.split(' ')
        system_download(base_url, user, pwd, delivery_no_list)
        sales_match_group()
        time.sleep(1)
        new_main_use_flow(result_send)


# 客户端 工具页
def tool_tiantu_simple_cs_window(tool):
    tool_use_page = tk.Frame(tool, height=884, width=1000, background='white', highlightcolor='black', relief='ridge')
    # 右边功能窗口
    r1 = tk.Label(tool_use_page, text='美线客服 根据出货箱号企微发送', background='white', justify='left')
    r1.place(x=0, y=0)

    d1 = tk.Label(tool_use_page, text='请求地址', background='white', justify='left')
    d1.place(x=0, y=25)

    # 设置默认安速系统环境
    e1 = tk.Entry(tool_use_page, width=30, relief='groove')
    e1.insert(0, 'https://fba.tttxex.com/')

    e1.place(x=100, y=28)

    d2 = tk.Label(tool_use_page, text='用户名', background='white')
    d2.place(x=0, y=45)

    e2 = tk.Entry(tool_use_page, width=30, relief='groove')
    e2.place(x=100, y=48)

    d3 = tk.Label(tool_use_page, text='密码', background='white')
    d3.place(x=0, y=65)

    e3 = tk.Entry(tool_use_page, width=30, relief='groove')
    e3.place(x=100, y=68)

    d5 = tk.Label(tool_use_page, text='出货箱号', background='white')
    d5.place(x=0, y=85)

    e5 = tk.Entry(tool_use_page, width=30, relief='groove')
    e5.place(x=100, y=88)

    d4 = tk.Label(tool_use_page, text='接收结果人', background='white')
    d4.place(x=0, y=105)

    e4 = tk.Entry(tool_use_page, width=30, relief='groove')
    e4.place(x=100, y=105)

    # e1.insert(0, 'https://fba.tttxex.com/')
    # e2.insert(0, 'admin')
    # e3.insert(0, '123456')
    # e4.insert(0, '天光')

    def clear():
        e1.delete(0, 'end')
        e2.delete(0, 'end')
        e3.delete(0, 'end')
        e4.delete(0, 'end')
        e5.delete(0, 'end')

    implement_button = tk.Button(tool_use_page,
                                 text='执行',
                                 background='#AFEEEE',
                                 command=lambda: cs_forward_main(e1.get(), e2.get(), e3.get(), e5.get(), e4.get())
                                 )
    implement_button.place(x=2, y=126)

    clear_button = tk.Button(tool_use_page,
                             text='清空',
                             background='#AFEEEE',
                             command=clear)
    clear_button.place(x=40, y=126)

    # 模块位置
    tool_use_page.place(x=220, y=20)
    return tool_use_page


if __name__ == '__main__':
    base_url = "https://fba.tttxex.com/"
    user = "admin"
    pwd = "rCaVObz/sHF3P60+STfZgg=="
    delivery_no_list = "CH000011744"
    result_send = "天光"
    cs_forward_main(base_url, user, pwd, delivery_no_list, result_send)
