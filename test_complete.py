from src.utils.md_parser import MDParser
from src.utils.ppt_generator import PPTGenerator
import time

# 测试完整的PPT生成流程
print('开始测试完整的PPT生成流程...')

# 1. 解析Markdown文件
print('1. 解析Markdown文件...')
parser = MDParser('# 说白色.md')
md_content = parser.parse()
print(f'解析完成，获得 {len(md_content["h2"])} 个三级标题')

# 2. 生成PPT
print('2. 生成PPT...')
generator = PPTGenerator('模板.pptx', md_content)
# 使用时间戳作为输出文件名，避免文件占用问题
output_path = f'output_{int(time.time())}.pptx'
generator.presentation.save(output_path)
print(f'PPT生成完成，输出路径: {output_path}')

print('测试完成！')