from src.utils.md_parser import MDParser

# 测试markdown解析
parser = MDParser('# 说白色.md')
content = parser.parse()

print('一级标题:', content['h0'])
print('二级标题:', content['h1'])
print('三级标题数量:', len(content['h2']))
print('第一个三级标题内容:', content['h2'][0] if content['h2'] else '无')