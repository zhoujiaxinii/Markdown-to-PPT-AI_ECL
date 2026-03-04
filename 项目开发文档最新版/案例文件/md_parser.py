# -*- coding: utf-8 -*-
"""
Markdown解析器
支持：3级标题、多文本框分隔、音视频识别、字数统计
"""

import re

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
            'h0': [],    # 一级标题 - 课程标题
            'h1': [],    # 二级标题 - 章节名
            'h2': [],    # 三级标题 - 小节名
            'content': []  # 内容
        }
        
        current_h1 = None
        current_h2 = None
        current_content = []
        
        for line in lines:
            line = line.strip()
            
            # 匹配一级标题 - 课程标题 (# 标题)
            if line.startswith('# ') and not line.startswith('##'):
                if current_h2:
                    parsed_content = self._parse_content(current_content)
                    result['h2'].append(self._create_h2_item(current_h2, parsed_content))
                    current_content = []
                if current_h1:
                    result['h1'].append(current_h1)
                result['h0'].append(line[2:].strip())
                current_h1 = None
                current_h2 = None
            
            # 匹配二级标题 - 章节名 (## 标题)
            elif line.startswith('## ') and not line.startswith('###'):
                if current_h2:
                    parsed_content = self._parse_content(current_content)
                    result['h2'].append(self._create_h2_item(current_h2, parsed_content))
                    current_content = []
                # 保存当前的章节名
                if current_h1:
                    result['h1'].append(current_h1)
                current_h1 = line[3:].strip()
                current_h2 = None
            
            # 匹配三级标题 - 小节名 (### 标题)
            elif line.startswith('### ') and not line.startswith('####'):
                if current_h2:
                    parsed_content = self._parse_content(current_content)
                    result['h2'].append(self._create_h2_item(current_h2, parsed_content))
                    current_content = []
                current_h2 = line[4:].strip()
            
            # 匹配四级标题 - 内容标题
            elif line.startswith('#### '):
                content_text = line[5:].strip()
                if content_text:
                    current_content.append({'type': 'text', 'content': content_text})
            
            # 匹配图片 - ![图片](url) 或 ![图片1](url)
            elif re.search(r'!\[图片[^\]]*\]', line):
                match = re.search(r'!\[图片[^\]]*\]\(([^)]+)\)', line)
                if match:
                    current_content.append({'type': 'image', 'url': match.group(1).strip()})
            
            # 匹配视频 - ![视频](url)
            elif '![' in line and '视频' in line:
                match = re.search(r'!\[视频\]\(([^)]+)\)', line)
                if match:
                    current_content.append({'type': 'video', 'url': match.group(1).strip()})
            
            # 匹配音频 - ![音频](url)
            elif '![' in line and '音频' in line:
                match = re.search(r'!\[音频\]\(([^)]+)\)', line)
                if match:
                    current_content.append({'type': 'audio', 'url': match.group(1).strip()})
            
            # 匹配其他链接（非音视频）
            elif 'http' in line and '.mp3' not in line.lower() and '.mp4' not in line.lower():
                current_content.append({'type': 'link', 'url': line})
            
            # 普通文本（非空行）
            elif line and not line.startswith('#'):
                current_content.append({'type': 'text', 'content': line})
        
        # 处理最后一个小节
        if current_h2:
            parsed_content = self._parse_content(current_content)
            result['h2'].append(self._create_h2_item(current_h2, parsed_content))
        elif current_h1:
            result['h1'].append(current_h1)
        
        return result
    
    def _create_h2_item(self, title, parsed_content):
        """创建小节项，包含字数限制提示"""
        char_count = parsed_content['char_count']
        return {
            'title': title,
            'content': parsed_content['texts'],
            'images': parsed_content['images'],
            'video': parsed_content['video'],
            'audio': parsed_content['audio'],
            'char_count': char_count,
            'word_limit_tip': self.get_word_limit_tip(char_count, len(parsed_content['texts']))
        }
    
    def _parse_content(self, content_list):
        """解析内容列表，提取文本、图片、视频、音频"""
        texts = []
        images = []
        video = None
        audio = None
        total_char_count = 0
        
        for item in content_list:
            if item['type'] == 'text':
                # 检查是否包含多文本框分隔符 ||
                if '||' in item['content']:
                    text_parts = item['content'].split('||')
                    for part in text_parts:
                        part = part.strip()
                        if part:
                            texts.append(part)
                            total_char_count += len(part)
                else:
                    texts.append(item['content'])
                    total_char_count += len(item['content'])
            elif item['type'] == 'image':
                images.append(item['url'])
            elif item['type'] == 'video':
                video = item['url']
            elif item['type'] == 'audio':
                audio = item['url']
        
        return {
            'texts': texts,
            'images': images,
            'video': video,
            'audio': audio,
            'char_count': total_char_count
        }
    
    def get_word_limit_tip(self, char_count, text_box_count):
        """根据字数和文本框数量返回限制范围提示"""
        # 推荐文本框数量
        if text_box_count == 0:
            box_recommend = "建议0-1个文本框"
        elif text_box_count == 1:
            box_recommend = "1个文本框（合适）"
        elif text_box_count == 2:
            box_recommend = "2个文本框（合适）"
        elif text_box_count == 3:
            box_recommend = "3个文本框（合适）"
        else:
            box_recommend = "4个文本框（已满）"
        
        # 字数范围
        if char_count == 0:
            char_range = "0字（空白页）"
        elif char_count <= 50:
            char_range = "1-50字（简短）"
        elif char_count <= 100:
            char_range = "51-100字（中等）"
        elif char_count <= 200:
            char_range = "101-200字（较长）"
        elif char_count <= 300:
            char_range = "201-300字（较长）"
        elif char_count <= 400:
            char_range = "301-400字（内容较多）"
        else:
            char_range = f"{char_count}字（内容过多，建议分页）"
        
        return f"{char_range} | {box_recommend}"


# 测试用
if __name__ == '__main__':
    import sys
    if len(sys.argv) > 1:
        parser = MDParser(sys.argv[1])
        result = parser.parse()
        print("=== 解析结果 ===")
        print(f"课程标题: {result['h0']}")
        print(f"章节: {result['h1']}")
        print(f"小节数量: {len(result['h2'])}")
        for i, h2 in enumerate(result['h2'][:3]):
            print(f"  小节{i+1}: {h2['title']}")
            print(f"    文本框数: {len(h2['content'])}")
            print(f"    图片数: {len(h2['images'])}")
            print(f"    字数: {h2['char_count']}")
            print(f"    提示: {h2['word_limit_tip']}")
            if h2['video']:
                print(f"    视频: {h2['video']}")
            if h2['audio']:
                print(f"    音频: {h2['audio']}")
    else:
        print("用法: python md_parser.py <md文件路径>")