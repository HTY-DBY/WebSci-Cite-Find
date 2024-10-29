import os

from docx import Document

from Global_Using import Global_Using

# 用于从word中获取各个文章的标题，并存储到txt中做观察和备份

# 定义路径和文件名
my_path = Global_Using().my_path
file_name = 'Ref.docx'
file_path = os.path.join(my_path, file_name)

# 读取 Word 文件
doc = Document(file_path)
content = [para.text for para in doc.paragraphs]

# 提取每段的第一个 '.' 和第二个 '.' 之间的内容
NeedFind_PaperTitle = []
for text in content:
	parts = text.split('.')
	if len(parts) > 2:  # 确保有足够的部分
		title = parts[1].strip()
		NeedFind_PaperTitle.append(title)

# 将提取的内容写入到文件中，每个条目一行
with open(os.path.join(my_path, 'Ref_Title.txt'), 'w', encoding='utf-8') as f:
	f.write('\n'.join(NeedFind_PaperTitle))
print('get title ok')
