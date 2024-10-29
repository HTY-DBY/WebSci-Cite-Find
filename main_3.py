import os
import re
import time
import pandas as pd
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from Global_Using import Global_Using

# 定义路径和文件名
my_path = Global_Using().my_path
file_name = 'Ref_index_search.xlsx'
file_path = os.path.join(my_path, file_name)

# 存储列表
Ref_index_main = pd.read_excel(file_path)
address_all = Ref_index_main['单位1'].tolist()
# 找到 NaN 值的位置
nan_positions = [index for index, value in enumerate(address_all) if pd.isna(value)]

# 存储列表
url_all = Ref_index_main.iloc[nan_positions, 4].tolist()
one_paper_data = []
SCROLL_PAUSE_TIME = 0.05  # 每次滚动的时长
scroll_step = 500  # 每次滚动的步长

driver = Global_Using().init_driver()
driver.get('https://webofscience.clarivate.cn/wos/alldb/basic-search')

Global_Using().accept_cookies(driver)


def scroll():
	last_height = driver.execute_script("return document.body.scrollHeight")
	new_height = last_height
	# 滚动到页面底部
	while True:
		driver.execute_script(f"window.scrollBy(0, {scroll_step});")
		time.sleep(SCROLL_PAUSE_TIME)

		new_height = driver.execute_script("return document.body.scrollHeight")

		if new_height == last_height:
			driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
			time.sleep(SCROLL_PAUSE_TIME)

			new_height = driver.execute_script("return document.body.scrollHeight")

			if new_height == last_height:
				# print("已经滚动到页面底部")
				break

		last_height = new_height


for iii, paper_index in enumerate(url_all):
	cleaned_string = paper_index.strip('[]').replace("'", "")
	url_index_list = [url.strip() for url in cleaned_string.split(',')]

	for jjj, url_index_in_a_paper in enumerate(url_index_list):
		print(f'进度：{iii + 1}/{len(url_all)}--{jjj + 1}/{len(url_index_list)}')

		sup_1 = 0
		driver.get(url_index_in_a_paper)
		scroll()

		try:
			more_button = WebDriverWait(driver, 2).until(
				EC.element_to_be_clickable((By.XPATH, '//*[@id="FRACTa-authorAddressView"]'))
			)
			more_button.click()
		except Exception:
			pass  # 点击更多按钮失败，不做处理
		try:
			address_element = WebDriverWait(driver, 10).until(
				EC.element_to_be_clickable((By.XPATH, '//*[@id="address_1"]/span[2]'))
			)
			address_element_text = address_element.text
			sup_1 = 1
		except Exception as e:
			try:
				scroll()  ##
				address_element = WebDriverWait(driver, 10).until(
					EC.element_to_be_clickable((By.XPATH, '//*[@id="snMainArticle"]/div[9]/span'))
				)
			except Exception as e:
				print(e)
				print(f"\033[91m文章 '{Ref_index_main.iloc[iii, 0]}' 出现错误\033[0m")
				print(f"\033[91m在 {url_index_in_a_paper}\033[0m")
				address_element_text = ""

		# 按行分割文本并提取每行数字后面的地址
		address_list = []
		if sup_1 == 0:
			lines = address_element_text.splitlines()
			for i, line in enumerate(lines):
				if re.match(r'^\d+$', line.strip()) and i + 1 < len(lines):
					address = lines[i + 1].strip()
					if address:
						address_list.append(address)
		else:
			address_list.append(address_element_text)
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

for entry in one_paper_data:
	units = entry.pop('单位')
	numbered_units = {f'单位{i + 1}': unit for i, unit in enumerate(units)}
	entry.update(numbered_units)

for index, nan_positions_addresss in enumerate(nan_positions):
	need_data = one_paper_data[1]
	keys = list(need_data.keys())
	start_index = keys.index('单位1')
	keys_after_unit1_count = len(keys) - (start_index + 1)
	for i in range(keys_after_unit1_count + 1):
		Ref_index_main.at[nan_positions_addresss, '单位' + str(i + 1)] = need_data['单位' + str(i + 1)]

df = pd.DataFrame(Ref_index_main)
# df = df.sort_values(by=['文章index', '引用index'], ascending=True)
df.to_excel('Ref_index_search_fin.xlsx', index=False)
