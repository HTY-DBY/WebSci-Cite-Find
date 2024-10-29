import time

import pandas as pd
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.options import Options
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.microsoft import EdgeChromiumDriverManager

import Get_title

NeedFind_PaperTitle = Get_title.NeedFind_PaperTitle


# 初始化浏览器的设置
def init_driver():
	options = Options()
	options.add_argument("--headless")  # 无头模式
	# options.add_argument("--disable-gpu")  # 禁用 GPU 加速（有时会提高稳定性）
	options.add_argument("--no-sandbox")  # 防止沙盒问题（在某些系统上可能需要）

	# 创建 Edge 服务
	edge_service = EdgeService(EdgeChromiumDriverManager().install())
	driver = webdriver.Edge(service=edge_service, options=options)
	return driver


driver = init_driver()
driver.get('https://webofscience.clarivate.cn/wos/alldb/basic-search')


# 点击接受cookie按钮
def accept_cookies():
	print("等待点击 cookie 按钮中")
	while True:
		try:
			accept_button = WebDriverWait(driver, 10).until(
				EC.element_to_be_clickable((By.XPATH, '//*[@id="onetrust-accept-btn-handler"]'))
			)
			accept_button.click()
			print("已点击 cookie 按钮")
			break
		except Exception as e:
			print(f"点击 cookie 按钮失败，继续尝试")
			time.sleep(1)


accept_cookies()

SCROLL_PAUSE_TIME = 0.05  # 每次滚动的时长
scroll_step = 500  # 每次滚动的步长


def scroll():
	"""滚动到页面底部"""
	last_height = driver.execute_script("return document.body.scrollHeight")
	while True:
		driver.execute_script(f"window.scrollBy(0, {scroll_step});")
		time.sleep(SCROLL_PAUSE_TIME)
		new_height = driver.execute_script("return document.body.scrollHeight")
		if new_height == last_height:
			break
		last_height = new_height


def search_paper_titles(titles):
	# 存储列表
	paper_data = []
	"""搜索每个标题并获取引用数据"""
	for index, title in enumerate(titles):
		try:
			input_box = WebDriverWait(driver, 10).until(
				EC.presence_of_element_located((By.ID, 'search-option'))
			)
			input_box.clear()
			input_box.send_keys(title)
			print(f'进度：{index + 1}/{len(titles)}，搜索文章: {title}')

			search_button = driver.find_element(By.XPATH, '//*[@id="snSearchType"]/div[3]/button[2]')
			search_button.click()

			element = WebDriverWait(driver, 10).until(
				EC.presence_of_element_located((By.XPATH, '/html/body/app-wos/main/div/div/div[2]/div/div/div[2]/app-input-route/app-base-summary-component/div/div[2]/app-records-list/app-record/div/div/div[4]/div/div[1]/div[1]/a'))
			)

			ref_num = element.text  # 获取引用次数
			ref_main_href = element.get_attribute('href')

			print('获取到总引用链接')
			driver.get(ref_main_href)

			href_value_index = []
			i_page = 1
			while True:
				try:
					WebDriverWait(driver, 2).until(
						EC.element_to_be_clickable((By.XPATH, '/html/body/app-wos/main/div/div/div[2]/div/div/div[2]/app-input-route/app-base-summary-component/div/div[2]/app-records-list/app-record[1]/div'))
					)

					scroll()  # 滚动页面

					# 获取页面源码
					page_source = driver.page_source
					soup = BeautifulSoup(page_source, 'html.parser')
					for link in soup.find_all('a', href=True):
						href = link['href']
						if href.startswith('/wos/alldb/full-record'):
							full_href = 'https://webofscience.clarivate.cn' + href
							href_value_index.append(full_href)

					i_page += 1
					driver.get(driver.current_url[:-1] + str(i_page))

				except Exception:
					ref_num_real = len(href_value_index)
					pd_name_list = ['标题', '假引用次数', '总查询链接', '真引用次数', '各引用的查询链接']
					index_data = {
						pd_name_list[0]: title,
						pd_name_list[1]: ref_num,
						pd_name_list[2]: ref_main_href,
						pd_name_list[3]: ref_num_real,
						pd_name_list[4]: href_value_index
					}
					paper_data.append(index_data)

					print(f"真引用次数：{ref_num_real}")
					break


		except Exception as e:
			print(f"\033[91m文章 '{title}' 出现错误: {e}\033[0m")
		finally:
			# 返回搜索页面
			driver.get('https://webofscience.clarivate.cn/wos/alldb/basic-search')
	return pd.DataFrame(paper_data)


paper_data = search_paper_titles(NeedFind_PaperTitle)
# 存储为 Excel
paper_data = pd.DataFrame(paper_data)
paper_data.to_excel('Data.xlsx', index=False)
