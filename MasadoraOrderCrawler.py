import requests
import json
import yaml
from datetime import datetime
from openpyxl import Workbook
import os
import tkinter as tk
from tkinter import messagebox

try:
    # 读取配置文件
    with open('config.yaml', 'r', encoding='utf-8') as config_file:
        config = yaml.safe_load(config_file)

    headers = config['headers']
    cookies = config['cookies']
    path = config['path']
except FileNotFoundError:
    root = tk.Tk()
    root.withdraw()
    messagebox.showerror("Error", "Configuration file 'config.yaml' not found.")
    exit(1)
except yaml.YAMLError as e:
    root = tk.Tk()
    root.withdraw()
    messagebox.showerror("Error", f"Error reading configuration file: {e}")
    exit(1)

# 创建一个新的工作簿
wb = Workbook()
ws = wb.active
ws.append(
    ["订单号", "支付状态", "物流状态", "下单时间", "产品名称", "原价（日元）", "含手续费的订单金额（日元）", "来源网站",
     "产品链接"])
# 如果路径不存在，则创建目录
os.makedirs(os.path.dirname(path), exist_ok=True)

try:
    # 发送请求并遍历订单页
    page = 0  # 起始页
    while True:
        url = f'https://www.masadora.jp/api/order?page={page}&size=100&timeSel=5&type=all'
        response = requests.get(url, headers=headers, cookies=cookies)
        response.encoding = 'utf-8'
        parsed_data = response.json()
        orders = parsed_data['data']['content']

        # 如果没有订单数据，则跳出循环
        if not orders:
            break

        # 提取订单信息并写入Excel表格
        for order_info in orders:
            order_time = datetime.fromtimestamp(order_info['createTime'] / 1000)
            for product_info in order_info['products']:
                ws.append([
                    order_info['domesticOrderNo'],
                    order_info['payStatus']['text'],
                    order_info['logisticsStatus']['text'],
                    order_time.strftime('%Y-%m-%d %H:%M:%S'),
                    product_info['name'],
                    product_info['price'],
                    order_info['domesticPriceModel']['totalPrice'],
                    order_info['sourceSite']['siteName'],
                    product_info['url']
                ])

        # 如果已经到达最后一页，则跳出循环
        if page >= parsed_data['data']['totalPage']:
            break

        # 前进到下一页
        page += 1

    # 保存工作簿
    wb.save(path)
    root = tk.Tk()
    root.withdraw()
    messagebox.showinfo("Success", "Excel 文件生成成功！")
except requests.RequestException as e:
    root = tk.Tk()
    root.withdraw()
    messagebox.showerror("Error", f"Error making HTTP request: {e}")
except KeyError:
    root = tk.Tk()
    root.withdraw()
    messagebox.showerror("Error", "Unexpected response format.")
except Exception as e:
    root = tk.Tk()
    root.withdraw()
    messagebox.showerror("Error", f"An unexpected error occurred: {e}")
