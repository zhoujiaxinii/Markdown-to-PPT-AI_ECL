from pptx import Presentation
from pptx.util import Inches
import re
from pypinyin import pinyin, Style

class PPTGenerator:
    def __init__(self, template_path, md_content):
        self.template_path = template_path
        self.md_content = md_content
        self.presentation = Presentation(template_path)
    
    def _add_pinyin(self, text):
        """为汉字添加拼音，拼音显示在汉字上方"""
        # 对于PowerPoint，我们需要创建带有拼音的文本
        # 这里使用特殊格式，后续在替换占位符时处理
        result = []
        for char in text:
            if '\u4e00' <= char <= '\u9fff':  # 判断是否为汉字
                py = pinyin(char, style=Style.NORMAL)[0][0]
                # 存储为带有拼音信息的结构
                result.append({'char': char, 'pinyin': py})
            else:
                result.append({'char': char, 'pinyin': None})
        return result
    
    def _find_placeholder(self, slide, placeholder_type):
        """查找指定类型的占位符"""
        for shape in slide.shapes:
            if shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        # 查找包含占位符的文本
                        if placeholder_type in run.text:
                            return run
        return None
    
    def _replace_placeholder(self, slide, placeholder_type, content):
        """替换占位符内容"""
        placeholder = self._find_placeholder(slide, placeholder_type)
        if placeholder:
            # 检查content是否为带有拼音信息的列表
            if isinstance(content, list) and all(isinstance(item, dict) for item in content):
                # 清空现有文本
                placeholder.text = ''
                # 为每个字符添加拼音
                result = []
                for item in content:
                    char = item['char']
                    pinyin = item['pinyin']
                    if pinyin:
                        result.append(f'{char}（{pinyin}）')
                    else:
                        result.append(char)
                placeholder.text = ''.join(result)
            else:
                # 普通文本
                placeholder.text = content
    
    def _match_slides(self):
        """匹配模板页，根据内容自动匹配对应的模板"""
        matched_slides = []
        
        # 匹配封面页
        cover_slide = None
        for slide in self.presentation.slides:
            if self._find_placeholder(slide, 'h0_0'):
                cover_slide = slide
                break
        if cover_slide:
            matched_slides.append(cover_slide)
        
        # 匹配目录页
        toc_slide = None
        for slide in self.presentation.slides:
            if self._find_placeholder(slide, 'h1_0'):
                toc_slide = slide
                break
        if toc_slide:
            matched_slides.append(toc_slide)
        
        # 分类模板页
        section_slides = []
        content_slides = []  # 所有正文模板
        
        for slide in self.presentation.slides:
            if self._find_placeholder(slide, 'h1_0') and slide != toc_slide:
                section_slides.append(slide)
            elif self._find_placeholder(slide, 'h2_0'):
                content_slides.append(slide)
        
        # 匹配结束页
        end_slide = None
        for slide in self.presentation.slides:
            for shape in slide.shapes:
                if shape.has_text_frame:
                    for paragraph in shape.text_frame.paragraphs:
                        if 'xiè xie' in paragraph.text:
                            end_slide = slide
                            break
            if end_slide:
                break
        
        # 添加章节页
        for i, h1 in enumerate(self.md_content['h1']):
            if i < len(section_slides):
                matched_slides.append(section_slides[i])
            else:
                # 如果章节页不够，复制一个
                if section_slides:
                    new_slide = self.presentation.slides.add_slide(section_slides[0].slide_layout)
                    matched_slides.append(new_slide)
        
        # 添加正文页
        for i, h2 in enumerate(self.md_content['h2']):
            if i < len(content_slides):
                matched_slides.append(content_slides[i])
            else:
                # 如果正文页不够，复制一个
                if content_slides:
                    new_slide = self.presentation.slides.add_slide(content_slides[0].slide_layout)
                    matched_slides.append(new_slide)
        
        # 添加结束页
        if end_slide:
            matched_slides.append(end_slide)
        
        # 删除未匹配的幻灯片
        slides_to_delete = [slide for slide in self.presentation.slides if slide not in matched_slides]
        for slide in slides_to_delete:
            idx = self.presentation.slides.index(slide)
            self.presentation.slides._sldIdLst.remove(self.presentation.slides._sldIdLst[idx])
    
    def _generate_content(self):
        """生成PPT内容"""
        # 处理封面页
        if self.md_content['h0'] and len(self.presentation.slides) > 0:
            cover_slide = self.presentation.slides[0]
            title = self._add_pinyin(self.md_content['h0'][0])
            self._replace_placeholder(cover_slide, 'h0_0', title)
        
        # 处理目录页
        if len(self.presentation.slides) > 1:
            toc_slide = self.presentation.slides[1]
            for i, h1 in enumerate(self.md_content['h1']):
                placeholder = f'h1_{i}'
                content = self._add_pinyin(h1)
                self._replace_placeholder(toc_slide, placeholder, content)
        
        # 处理章节页和正文页
        slide_idx = 2
        h2_index = 0
        
        # 处理所有章节页
        for i, h1 in enumerate(self.md_content['h1']):
            if slide_idx < len(self.presentation.slides):
                section_slide = self.presentation.slides[slide_idx]
                self._replace_placeholder(section_slide, 'h1_0', self._add_pinyin(h1))
                slide_idx += 1
        
        # 处理所有正文页
        while h2_index < len(self.md_content['h2']) and slide_idx < len(self.presentation.slides):
            h2 = self.md_content['h2'][h2_index]
            content_slide = self.presentation.slides[slide_idx]
            
            # 替换三级标题
            self._replace_placeholder(content_slide, 'h2_0', self._add_pinyin(h2['title']))
            
            # 替换内容
            for j, content_item in enumerate(h2['content']):
                if content_item['type'] == 'text':
                    placeholder = f'h3_{j}'
                    self._replace_placeholder(content_slide, placeholder, content_item['content'])
                elif content_item['type'] == 'image':
                    placeholder = f'h4_{j}'
                    self._replace_placeholder(content_slide, placeholder, content_item['url'])
                elif content_item['type'] == 'audio':
                    placeholder = f'h5_{j}'
                    self._replace_placeholder(content_slide, placeholder, content_item['url'])
                elif content_item['type'] == 'video':
                    placeholder = f'h6_{j}'
                    self._replace_placeholder(content_slide, placeholder, content_item['url'])
                elif content_item['type'] == 'link':
                    placeholder = f'h4_{j}'
                    self._replace_placeholder(content_slide, placeholder, content_item['url'])
            
            slide_idx += 1
            h2_index += 1
    
    def generate(self):
        """生成PPT"""
        self._match_slides()
        self._generate_content()
        output_path = 'output.pptx'
        self.presentation.save(output_path)
        return output_path