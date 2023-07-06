import ftplib
import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
import time
import os
import shutil
from datetime import date
import openpyxl
import xlrd
import sys

# 全局变量
gl_files_names = []


def get_script_directory():
    # 获取当前脚本所在的路径
    script_dir = os.path.dirname(os.path.abspath(sys.argv[0]))
    return script_dir


def download_files_from_ftp():
    # FTP服务器的连接信息
    ftp_host = "114.116.255.153"
    ftp_user = "FLIGHTONTIME"
    ftp_password = "FLIGHTONTIME@123"

    # 获取当前日期和前一天的日期
    current_date = datetime.date.today()
    previous_date = current_date - datetime.timedelta(days=1)

    # 格式化日期为yyyymmdd格式
    current_date_str = current_date.strftime("%Y%m%d")
    previous_date_str = previous_date.strftime("%Y%m%d")

    # 构建文件名列表
    file_names = [
        f"{previous_date_str}-lvke.txt",
        f"{previous_date_str}-huoyou.txt",
        f"select-preplan_preplan_valid-{current_date_str}.csv",
        f"select-tidb-flight_data_fme-{previous_date_str}.txt"
    ]

    # 连接FTP服务器并下载文件
    with ftplib.FTP(ftp_host, ftp_user, ftp_password) as ftp:
        ftp.cwd("/filesync")  # 切换到文件路径
        for file_name in file_names:
            with open(file_name, "wb") as file:
                ftp.retrbinary(f"RETR {file_name}", file.write)
            print(f"下载文件 {file_name} 完成")

    print("所有文件下载完成")


def upload_files_to_website():
    # 获取当前脚本所在的路径
    script_dir = get_script_directory()

    # 网站地址
    website_url = "http://192.168.101.32:8188/#/upload"
    # 使浏览器不自动关闭
    option = webdriver.ChromeOptions()
    option.add_experimental_option("detach", True)
    try:
        # 创建Chrome浏览器实例
        driver = webdriver.Chrome(options=option)

    except Exception as e:
        print('你没有正确安装ChromeDriver驱动（如果已安装请检查环境变量是否正确设置），无法运行本脚本。')
        pass

    print(website_url)
    # 打开网站
    driver.get(website_url)

    # 选择日期
    date_input = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "//div[@class='el-input__wrapper']//input"))
    )

    # 获取今天的日期并格式化为yyyy-mm-dd
    today = datetime.date.today()
    today_str = today.strftime("%Y-%m-%d")
    today_str2 = today.strftime("%Y%m%d")

    # 在日期选择框中输入今天的日期
    date_input.clear()
    date_input.send_keys(today_str)
    previous_date = today - datetime.timedelta(days=1)
    previous_date_str = previous_date.strftime("%Y%m%d")

    #  模拟点击空白处 定位到空白区域元素
    # element = driver.find_element_by_xpath("//body")
    element = driver.find_element(By.XPATH, "//body")

    # 模拟鼠标点击操作
    actions = ActionChains(driver)
    actions.move_to_element(element).click().perform()
    time.sleep(1)
    # 等待选择文件按钮加载完成
    upload_button = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "//button[@class='el-button el-button--primary']"))
    )
    time.sleep(1)

    # 等待文件选择对话框加载完成
    file_input = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "//input[@type='file']"))
    )
    time.sleep(1)



    A = gl_files_names[0]
    B = gl_files_names[1]
    C = gl_files_names[2]

    # 上传文件
    file_names = [
        previous_date_str + "-lvke.txt",
        previous_date_str + "-huoyou.txt",
        "select-preplan_preplan_valid-" + today_str2 + ".csv",
        "select-tidb-flight_data_fme-" + previous_date_str + ".txt",
        A, B, C
    ]
    for file_name in file_names:
        file_path = os.path.join(script_dir, file_name)
        print(file_name)
        file_input.send_keys(file_path)
    upload_button = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "//button[@class='el-button el-button--success ml-3']"))
    )

    actions.move_to_element(upload_button).click().perform()


    print("文件上传完成")
    print("文件上传完成")
    print("文件上传完成")


def copy_matching_excel_files():
    # 读取配置文件
    config_file = 'test.cfg'
    with open(config_file, 'r') as f:
        config = f.read()

    destination_dir = os.path.dirname(sys.executable)
    wx_path_m = config
    todayw = datetime.date.today()
    todayw_str = todayw.strftime("%Y-%m")
    source_dir = wx_path_m + todayw_str
    print('从这个路径获取机场运输业务量统计表')
    print(source_dir)
    expected_files = 3
    obtained_files = []

    while len(obtained_files) < expected_files:
        today_files = []

        # 遍历源目录中的文件
        for root, dirs, files in os.walk(source_dir):
            for file in files:
                file_path = os.path.join(root, file)

                # 检查文件是否为Excel文件
                if file.endswith('.xlsx'):
                    file_mtime = datetime.date.fromtimestamp(os.path.getmtime(file_path))
                    # 判断.xlsx文件是不是今天的
                    if file_mtime == todayw:
                        # 判断A1单元格是不是符合要求
                        wb = openpyxl.load_workbook(file_path)
                        sheet = wb.active
                        if sheet['A1'].value == "机场运输业务量统计表":
                            # 复制文件到目标目录
                            # file_name = os.path.basename(file_path)
                            # destination_path = os.path.join(destination_dir, file_name)
                            # shutil.copy(file_path, destination_path)
                            today_files.append(file_path)


                elif file.endswith('.xls'):
                    # 检查文件的修改日期是否为今天
                    file_mtime = datetime.date.fromtimestamp(os.path.getmtime(file_path))
                    if file_mtime == todayw:
                        # 检查A1单元格的内容是否为"机场运输业务量统计表"
                        wb = xlrd.open_workbook(file_path)
                        sheet = wb.sheet_by_index(0)
                        if sheet.cell_value(0, 0) == "机场运输业务量统计表":
                            today_files.append(file_path)

        # 复制符合条件的文件到目标目录
        for file_path in today_files:
            try:
                # 复制文件到目标目录
                file_name = os.path.basename(file_path)
                gl_files_names.append(file_name)
                destination_path = os.path.join(destination_dir, file_name)
                shutil.copy(file_path, destination_path)
                print(f'已复制文件：{file_name}')
                obtained_files.append(file_path)

            except Exception as e:
                print(f'复制文件时出错：{file_path}')
                print(f'错误信息：{str(e)}')

        if len(obtained_files) < expected_files:
            print(f'已获取到 {len(obtained_files)} 个文件，等待 30 秒后重新获取...')
            time.sleep(30)
        print(gl_files_names)


def copy_wechat_excel():
    # 读取配置文件
    config_file = 'test.cfg'
    with open(config_file, 'r') as f:
        config = f.read()

    destination_dir = os.path.dirname(sys.executable)
    wx_path_m = config
    todayw = datetime.date.today()
    todayw_str = todayw.strftime("%Y-%m")
    source_dir = wx_path_m + todayw_str
    # 生成年份和月份的路径
    today = datetime.date.today()
    year_month_path = os.path.join(wx_path_m, today.strftime('%Y-%m'))

    # 查找符合条件的Excel文件
    search_date = today - datetime.timedelta(days=1)
    search_date_str = search_date.strftime('%#m月%#d日')
    today_date_str = today.strftime('%#m月%#d日')
    excel_file_pattern = f'{source_dir}\{search_date_str}运行情况及{today_date_str}计划情况.xlsx'
    print(excel_file_pattern)
    try:
        shutil.copy(excel_file_pattern, destination_dir)
    except:
        print('运行情况及计划情况.xlsx不存在,30秒后重新获取')
        time.sleep(30)
        copy_wechat_excel()
    print('文件复制完成。')


if __name__ == '__main__':
    print('日简报生成V1.1,Powered by cjy 20230706，内部使用')
    try:
        download_files_from_ftp()
    except:
        print('ftp下载文件失败，15s后重试')
        time.sleep(15)
        try:
            download_files_from_ftp()
        except:
            print('还是不行，但是跳过（可能是已经有了）')
            pass
    print('ftp下载完成')
    print('ftp下载完成')
    print('ftp下载完成')
    print('ftp下载完成')
    print('ftp下载完成')
    time.sleep(3)
    copy_wechat_excel()
    time.sleep(3)
    print('运行情况和计划情况excel获取完成')
    print('运行情况和计划情况excel获取完成')
    print('运行情况和计划情况excel获取完成')
    print('运行情况和计划情况excel获取完成')
    print('运行情况和计划情况excel获取完成')
    print('接下来获取剩余的excel文件（机场运输业务量统计表）')
    print('接下来获取剩余的excel文件（机场运输业务量统计表）')
    copy_matching_excel_files()
    time.sleep(3)
    print('准备开始上传，请勿乱动弹出窗口')
    print('准备开始上传，请勿乱动弹出窗口')
    print('准备开始上传，请勿乱动弹出窗口')
    print('准备开始上传，请勿乱动弹出窗口')
    print('准备开始上传，请勿乱动弹出窗口')
    time.sleep(1)
    upload_files_to_website()
