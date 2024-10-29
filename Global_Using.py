import time

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.options import Options
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.microsoft import EdgeChromiumDriverManager


class Global_Using:
	def __init__(self):
		self.my_path = r'D:\hty\creat\code\python\web sci search'

	# 初始化浏览器的设置
	def init_driver(self):
		options = Options()
		options.add_argument("--headless")  # 无头模式
		# options.add_argument("--disable-gpu")  # 禁用 GPU 加速（有时会提高稳定性）
		options.add_argument("--no-sandbox")  # 防止沙盒问题（在某些系统上可能需要）

		# 创建 Edge 服务
		edge_service = EdgeService(EdgeChromiumDriverManager().install())
		driver = webdriver.Edge(service=edge_service, options=options)
		return driver

	# 点击接受cookie按钮
	def accept_cookies(self, driver):
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
