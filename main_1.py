import time
import Get_title
from Global_Using import Global_Using
import pandas as pd
from bs4 import BeautifulSoup
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait


def scroll_to_bottom(driver, n=50, scroll_step=500, pause_time=0.03):
	"""向下滚动n次，每次滚动固定步长。"""
	for _ in range(n):
		driver.execute_script(f"window.scrollBy(0, {scroll_step});")
		time.sleep(pause_time)  # 暂停以模拟人为滚动，确保页面稳定加载


def search_paper_titles(titles):
	# 存储列表
	paper_data = []
	next_is = 0
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
					if next_is == 0:
						Global_Using().accept_nexts(driver)
						next_is = 1

					scroll_to_bottom(driver)  # 滚动页面

					# 获取页面源码
					page_source = driver.page_source
					soup = BeautifulSoup(page_source, 'html.parser')
					for link in soup.find_all('a', href=True):
						href = link['href']
						if href.startswith('/wos/alldb/full-record'):
							full_href = 'https://webofscience.clarivate.cn' + href
							href_value_index.append(full_href)
					i_page += 1
					driver.get(f"{driver.current_url[:-1]}{i_page}")
				except Exception:
					break  # 翻页结束
			ref_num_real = len(href_value_index)
			paper_data.append({
				'标题': title,
				'假引用次数': ref_num,
				'总查询链接': ref_main_href,
				'真引用次数': ref_num_real,
				'各引用的查询链接': href_value_index
			})
			print(f"真引用次数：{ref_num_real}")
		except Exception as e:
			print(f"\033[91m文章 '{title}' 出现错误: {e}\033[0m")
		finally:
			# 返回搜索页面
			driver.get('https://webofscience.clarivate.cn/wos/alldb/basic-search')
	return pd.DataFrame(paper_data)


NeedFind_PaperTitle = Get_title.NeedFind_PaperTitle
driver = Global_Using().init_driver()

driver.get('https://webofscience.clarivate.cn/wos/alldb/basic-search')
Global_Using().accept_cookies(driver)

paper_data = search_paper_titles(NeedFind_PaperTitle)

# 存储为 Excel
paper_data = pd.DataFrame(paper_data)
paper_data.to_excel('Data.xlsx', index=False)
