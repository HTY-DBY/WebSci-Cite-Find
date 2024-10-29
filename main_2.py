import concurrent
import os
import re
import time

import pandas as pd
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

from Global_Using import Global_Using


# 滚动页面
def scroll(driver):
	SCROLL_PAUSE_TIME = 0.05  # 每次滚动的时长
	driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
	time.sleep(SCROLL_PAUSE_TIME)


# 处理每篇文章的函数
def process_paper(iii, paper_index):
	driver = Global_Using().init_driver()
	driver.get('https://webofscience.clarivate.cn/wos/alldb/basic-search')
	Global_Using().accept_cookies(driver)

	cleaned_string = paper_index.strip('[]').replace("'", "")
	url_index_list = [url.strip() for url in cleaned_string.split(',')]

	for jjj, url_index_in_a_paper in enumerate(url_index_list):
		print(f'进度：{iii + 1}/{len(url_all)}--{jjj + 1}/{len(url_index_list)}')
		driver.get(url_index_in_a_paper)
		scroll(driver)

		try:
			more_button = WebDriverWait(driver, 2).until(
				EC.element_to_be_clickable((By.XPATH, '//*[@id="FRACTa-authorAddressView"]'))
			)
			more_button.click()
		except Exception:
			pass

		try:
			scroll(driver)
			address_element = WebDriverWait(driver, 10).until(
				EC.element_to_be_clickable((By.XPATH, '//*[@id="snMainArticle"]/div[9]/span'))
			)
			address_element_text = address_element.text
		except Exception as e:
			print(e)
			print(f"\033[91m文章 '{Ref_index_main.iloc[iii, 0]}' 出现错误\033[0m")
			print(f"\033[91m在 {url_index_in_a_paper}\033[0m")
			address_element_text = ""

		address_list = []
		lines = address_element_text.splitlines()
		for i, line in enumerate(lines):
			if re.match(r'^\d+$', line.strip()) and i + 1 < len(lines):
				address = lines[i + 1].strip()
				if address:
					address_list.append(address)

		index_data = {
			'标题': Ref_index_main.iloc[iii, 0],
			'假引用次数': Ref_index_main.iloc[iii, 1],
			'总查询链接': Ref_index_main.iloc[iii, 2],
			'真引用次数': Ref_index_main.iloc[iii, 3],
			'该次查询地址': url_index_in_a_paper,
			'文章index': iii,
			'引用index': jjj,
			'单位': address_list
		}
		one_paper_data.append(index_data)

	driver.quit()


# 定义路径和文件名
my_path = Global_Using().my_path
file_name = 'Data.xlsx'
file_path = os.path.join(my_path, file_name)

# 存储列表
Ref_index_main = pd.read_excel(file_path)
url_all = Ref_index_main.iloc[:, 4].tolist()
one_paper_data = []

# 使用多线程进行并发处理
with concurrent.futures.ThreadPoolExecutor(max_workers=os.cpu_count()) as executor:
	for iii, paper_index in enumerate(url_all):
		executor.submit(process_paper, iii, paper_index)

# 存储为 Excel
for entry in one_paper_data:
	units = entry.pop('单位')
	numbered_units = {f'单位{i + 1}': unit for i, unit in enumerate(units)}
	entry.update(numbered_units)

df = pd.DataFrame(one_paper_data)
df = df.sort_values(by=['文章index', '引用index'], ascending=True)
df.to_excel('Ref_index_search.xlsx', index=False)
