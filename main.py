# 导入 webdriver
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time
import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook

driver = webdriver.Chrome()
time.sleep(2)
data = []


url1 = "http://sj.tjj.ordos.gov.cn/datashow/quick/QuickShowAct.htm?cn=B0107&quickCode=HGND&treeCode=5fe98d958fe042de9035cb08cb8de697&defaultTime="
title1 = ["地区", "户籍总人口", "男", "女", "市镇人口", "乡村人口", "常住总人口", "城镇人口", "乡村人口"]

for i in range(2022, 1999, -1):
    driver.get(url1 + str(i))

    elements = driver.find_element(By.ID, "chartTable")
    html_str = elements.get_attribute("innerHTML")
    soup = BeautifulSoup(html_str, "html.parser")
    # 找到所有的<tr>元素
    trs = soup.find_all("tr")

    lines = []

    # 遍历每个<tr>元素
    for tr in trs:
        if tr.get("style") == "height:61PT;":
            # 找到<tr>元素中的所有<td>元素
            tds = tr.find_all("td")
            # 遍历每个<td>元素
            for td in tds:
                # 输出<td>元素的文本内容或"null"
                if td.text.strip():
                    # print(td.text)
                    lines.append(td.text)
                else:
                    # print("null")
                    lines.append("null")

    start_collecting = False
    flag = True
    for line in lines:
        if line == "达拉特旗":
            start_collecting = True
            data.append(i)

        elif line == "准格尔旗":
            start_collecting = False
        elif line == "鄂托克旗":
            start_collecting = True
            flag = False
        elif line == "乌审旗":
            start_collecting = False
        if start_collecting:
            if flag:
                data.append(line)
            else:
                data.append(line)
# 关闭浏览器
driver.quit()

# 创建一个新的Excel工作簿
workbook = Workbook()
# 选择默认的活动工作表
sheet = workbook.active
sheet.append(["年份"] + title1 * 3)
# 循环20次，将每个3*9的数据写入一行
for i in range(0, len(data), 28):
    # 将27个数据写入Excel的一行
    sheet.append(data[i : i + 28])
# 保存Excel文件
name = "年末总人口数及构成"
workbook.save(name + ".xlsx")
print("爬取完毕!")
