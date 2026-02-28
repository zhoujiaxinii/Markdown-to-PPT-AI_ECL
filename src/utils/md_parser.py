import re
import markdown

class MDParser:
    def __init__(self, md_path):
        self.md_path = md_path
        self.content = self._read_file()
    
    def _read_file(self):
        with open(self.md_path, 'r', encoding='utf-8') as f:
            return f.read()
    
    def parse(self):
        """解析markdown文件，提取结构化内容"""
        lines = self.content.split('\n')
        result = {
            'h0': [],  # 一级标题
            'h1': [],  # 二级标题
            'h2': [],  # 三级标题
            'h3': [],  # 四级标题
            'content': []  # 内容
        }
        
        current_h1 = None
        current_h2 = None
        current_content = []
        
        for line in lines:
            line = line.strip()
            
            # 匹配一级标题
            if line.startswith('# '):
                result['h0'].append(line[2:].strip())
            
            # 匹配二级标题
            elif line.startswith('## '):
                if current_h2:
                    result['h2'].append({'title': current_h2, 'content': current_content})
                    current_content = []
                current_h1 = line[3:].strip()
                result['h1'].append(current_h1)
            
            # 匹配三级标题
            elif line.startswith('### '):
                if current_h2:
                    result['h2'].append({'title': current_h2, 'content': current_content})
                    current_content = []
                current_h2 = line[4:].strip()
            
            # 匹配四级标题
            elif line.startswith('#### '):
                pass  # 四级标题作为内容分区标识
            
            # 匹配图片链接
            elif '![图片' in line:
                match = re.search(r'!\[图片[^\]]*\]\(([^)]+)\)', line)
                if match:
                    current_content.append({'type': 'image', 'url': match.group(1)})
            
            # 匹配其他链接
            elif 'http' in line:
                current_content.append({'type': 'link', 'url': line})
            
            # 普通文本
            elif line:
                current_content.append({'type': 'text', 'content': line})
        
        # 处理最后一个三级标题的内容
        if current_h2:
            result['h2'].append({'title': current_h2, 'content': current_content})
        
        return result