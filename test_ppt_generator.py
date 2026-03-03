from src.utils.md_parser import MDParser
from src.utils.ppt_generator import PPTGenerator

# 解析markdown文件
parser = MDParser('# 说白色（shuō bái sè）.MD')
content = parser.parse()

# 测试PPT生成
generator = PPTGenerator('模板.pptx', content)
try:
    output_path = generator.generate()
    print(f'PPT生成成功，输出路径: {output_path}')
except Exception as e:
    print(f'PPT生成失败: {str(e)}')