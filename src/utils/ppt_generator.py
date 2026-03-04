# -*- coding: utf-8 -*-
"""
PPT生成器 - 混合格式版本
保留文本显示，同时添加ruby拼音
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
from pptx.oxml.xmlchemy import OxmlElement
from pptx.oxml.ns import qn
import re
from pypinyin import pinyin, Style

class PPTGenerator:
    def __init__(self, template_path, md_content):
        self.template_path = template_path
        self.md_content = md_content
        self.presentation = Presentation(template_path)
    
    def _create_text_with_ruby(self, text_frame, text, font_size=24):
        """创建带Ruby拼音的文本，同时保留可见文本"""
        text_frame.clear()
        
        from pypinyin import pinyin as py_func, Style
        import re
        
        # 创建段落
        if not text_frame.paragraphs:
            text_frame.add_paragraph()
        para = text_frame.paragraphs[0]
        para.font.size = Pt(font_size)
        
        # 遍历每个字符
        for char in text:
            if '\u4e00' <= char <= '\u9fff':  # 汉字
                # 获取拼音
                py = py_func(char, style=Style.TONE)[0][0]
                py = re.sub(r'(\d)$', '', py)  # 去掉声调
                
                # 创建run - 先显示拼音+汉字的普通文本
                run = para.add_run()
                run.text = f'{py} {char}'  # 显示 "拼音 汉字" 格式
                run.font.size = Pt(font_size)
                
                # 同时添加ruby元素（用于精确显示）
                try:
                    ruby = OxmlElement('a:ruby')
                    
                    rubyPr = OxmlElement('a:rubyPr')
                    rubyPr.set(qn('a:rubyFont'), 'Arial')
                    rubyPr.set('sz', str(int(font_size * 100)))
                    
                    bt = OxmlElement('a:bt')
                    t_bt = OxmlElement('a:t')
                    t_bt.text = char
                    bt.append(t_bt)
                    
                    rt = OxmlElement('a:rt')
                    t_rt = OxmlElement('a:t')
                    t_rt.text = py
                    rt.append(t_rt)
                    
                    ruby.append(rubyPr)
                    ruby.append(bt)
                    ruby.append(rt)
                    
                    run._r.append(ruby)
                except:
                    pass
                
            else:  # 非汉字
                run = para.add_run()
                run.text = char
                run.font.size = Pt(font_size)
    
    def _replace_placeholder_with_pinyin(self, slide, placeholder_type, content, font_size=24):
        """替换占位符并添加拼音"""
        for shape in slide.shapes:
            if shape.has_text_frame:
                text_frame = shape.text_frame
                
                for para in text_frame.paragraphs:
                    for run in para.runs:
                        if placeholder_type in run.text:
                            self._create_text_with_ruby(text_frame, content, font_size)
                            return True
        return False
    
    def _replace_placeholder_simple(self, slide, placeholder_type, content, font_size=24):
        """简单替换占位符"""
        for shape in slide.shapes:
            if shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    for run in para.runs:
                        if placeholder_type in run.text:
                            run.text = content
                            return True
        return False
    
    def _find_placeholder_shape(self, slide, placeholder_type):
        """查找包含指定占位符的shape"""
        for shape in slide.shapes:
            if shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    for run in para.runs:
                        if placeholder_type in run.text:
                            return shape
        return None
    
    def _find_all_placeholders(self, slide):
        """查找页面所有占位符"""
        placeholders = {}
        for shape in slide.shapes:
            if shape.has_text_frame:
                text = shape.text_frame.text
                matches = re.findall(r'\{\{(\w+)\}\}', text)
                for match in matches:
                    if match not in placeholders:
                        placeholders[match] = shape
        return placeholders
    
    def _match_slides(self):
        """匹配模板页"""
        matched_slides = []
        
        # 1. 封面页
        cover_slide = None
        for slide in self.presentation.slides:
            if self._find_placeholder_shape(slide, 'h0_0'):
                cover_slide = slide
                break
        if cover_slide:
            matched_slides.append(('cover', cover_slide))
        
        # 2. 分类模板页
        section_slides = []
        content_slides = []
        
        for slide in self.presentation.slides:
            if slide == cover_slide:
                continue
            placeholders = self._find_all_placeholders(slide)
            if 'h1_0' in placeholders and 'h2_0' not in placeholders:
                section_slides.append(slide)
            elif 'h2_0' in placeholders:
                content_slides.append(slide)
        
        # 3. 目录页
        if content_slides:
            toc_slide = self.presentation.slides.add_slide(content_slides[0].slide_layout)
            matched_slides.append(('toc', toc_slide))
        
        # 4. 章节页
        for i, h1 in enumerate(self.md_content['h1']):
            if i < len(section_slides):
                matched_slides.append(('section', section_slides[i]))
            else:
                if section_slides:
                    new_slide = self.presentation.slides.add_slide(section_slides[0].slide_layout)
                    matched_slides.append(('section', new_slide))
                elif content_slides:
                    new_slide = self.presentation.slides.add_slide(content_slides[0].slide_layout)
                    matched_slides.append(('section', new_slide))
        
        # 5. 正文页
        for h2 in self.md_content['h2']:
            if h2.get('video'):
                if content_slides:
                    new_slide = self.presentation.slides.add_slide(content_slides[0].slide_layout)
                    matched_slides.append(('video', new_slide))
            elif h2.get('audio'):
                if content_slides:
                    new_slide = self.presentation.slides.add_slide(content_slides[0].slide_layout)
                    matched_slides.append(('audio', new_slide))
            else:
                if content_slides:
                    matched_slides.append(('content', content_slides[0]))
                else:
                    if cover_slide:
                        new_slide = self.presentation.slides.add_slide(cover_slide.slide_layout)
                        matched_slides.append(('content', new_slide))
        
        # 6. 结束页
        end_slide = None
        for slide in self.presentation.slides:
            for shape in slide.shapes:
                if shape.has_text_frame:
                    if 'xiè xie' in shape.text_frame.text or '谢谢' in shape.text_frame.text:
                        end_slide = slide
                        break
            if end_slide:
                break
        if end_slide:
            matched_slides.append(('end', end_slide))
        
        return matched_slides
    
    def _generate_content(self):
        """生成PPT内容"""
        # 1. 封面页
        if self.md_content['h0'] and len(self.presentation.slides) > 0:
            cover_slide = self.presentation.slides[0]
            title = self.md_content['h0'][0]
            self._replace_placeholder_with_pinyin(cover_slide, 'h0_0', title, font_size=36)
        
        # 2. 目录页
        if len(self.presentation.slides) > 1 and self.md_content['h1']:
            toc_slide = self.presentation.slides[1]
            for i, h1 in enumerate(self.md_content['h1']):
                if i == 0:
                    self._replace_placeholder_with_pinyin(toc_slide, 'h2_0', h1, font_size=24)
                else:
                    placeholder = f'h3_{i-1}' if i <= 4 else None
                    if placeholder:
                        self._replace_placeholder_with_pinyin(toc_slide, placeholder, h1, font_size=18)
        
        # 3. 章节页和正文页
        slide_idx = 2
        section_count = len(self.md_content['h1'])
        
        # 章节页
        for i in range(section_count):
            if slide_idx < len(self.presentation.slides):
                section_slide = self.presentation.slides[slide_idx]
                self._replace_placeholder_with_pinyin(section_slide, 'h1_0', self.md_content['h1'][i], font_size=32)
                slide_idx += 1
        
        # 正文页
        h2_index = 0
        while h2_index < len(self.md_content['h2']) and slide_idx < len(self.presentation.slides):
            h2 = self.md_content['h2'][h2_index]
            content_slide = self.presentation.slides[slide_idx]
            
            # 小节名
            self._replace_placeholder_with_pinyin(content_slide, 'h2_0', h2['title'], font_size=28)
            
            # 文本框
            for j, text in enumerate(h2.get('content', [])):
                if j <= 3:
                    placeholder = f'h3_{j}'
                    self._replace_placeholder_with_pinyin(content_slide, placeholder, text, font_size=20)
            
            # 图片
            for j in range(len(h2.get('images', []))):
                if j <= 4:
                    placeholder = f'img_{j}'
                    self._replace_placeholder_simple(content_slide, placeholder, f'[图片{j+1}]')
            
            # 音频
            if h2.get('audio'):
                self._replace_placeholder_simple(content_slide, 'audio_0', '[音频播放]')
            
            # 视频
            if h2.get('video'):
                self._replace_placeholder_simple(content_slide, 'video_0', '[视频播放]')
            
            slide_idx += 1
            h2_index += 1
    
    def generate(self):
        """生成PPT"""
        self._match_slides()
        self._generate_content()
        output_path = 'output.pptx'
        self.presentation.save(output_path)
        return output_path


if __name__ == '__main__':
    import sys
    from md_parser import MDParser
    
    if len(sys.argv) > 2:
        md_file = sys.argv[1]
        template_file = sys.argv[2]
        
        print(f"解析MD: {md_file}")
        parser = MDParser(md_file)
        md_content = parser.parse()
        
        print(f"生成PPT: {template_file} -> output.pptx")
        generator = PPTGenerator(template_file, md_content)
        output = generator.generate()
        print(f"完成: {output}")
    else:
        print("用法: python ppt_generator.py <md文件> <模板文件>")