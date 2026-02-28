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
        """为汉字添加拼音"""
        result = []
        for char in text:
            if '\u4e00' <= char <= '\u9fff':  # 判断是否为汉字
                py = pinyin(char, style=Style.NORMAL)[0][0]
                result.append(f'{char}（{py}）')
            else:
                result.append(char)
        return ''.join(result)
    
    def _find_placeholder(self, slide, placeholder_type):
        """查找指定类型的占位符"""
        for shape in slide.shapes:
            if shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        if f'{{{{{placeholder_type}' in run.text:
                            return run
        return None
    
    def _replace_placeholder(self, slide, placeholder_type, content):
        """替换占位符内容"""
        placeholder = self._find_placeholder(slide, placeholder_type)
        if placeholder:
            placeholder.text = content
    
    def _match_slides(self):
        """匹配模板页"""
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
        
        # 匹配章节页
        section_slides = []
        for slide in self.presentation.slides:
            if self._find_placeholder(slide, 'h1_0') and slide != toc_slide:
                section_slides.append(slide)
        
        # 匹配正文页
        content_slides = []
        for slide in self.presentation.slides:
            if self._find_placeholder(slide, 'h2_0'):
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
        
        # 计算需要的页面数量
        needed_section_slides = len(self.md_content['h1'])
        needed_content_slides = len(self.md_content['h2'])
        
        # 添加章节页
        for i in range(min(needed_section_slides, len(section_slides))):
            matched_slides.append(section_slides[i])
        
        # 添加正文页
        for i in range(min(needed_content_slides, len(content_slides))):
            matched_slides.append(content_slides[i])
        
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
        if self.md_content['h0']:
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
        for i, h1 in enumerate(self.md_content['h1']):
            # 处理章节页
            if slide_idx < len(self.presentation.slides):
                section_slide = self.presentation.slides[slide_idx]
                self._replace_placeholder(section_slide, 'h1_0', self._add_pinyin(h1))
                slide_idx += 1
            
            # 处理当前章节的正文页
            while h2_index < len(self.md_content['h2']):
                if slide_idx >= len(self.presentation.slides):
                    break
                
                h2 = self.md_content['h2'][h2_index]
                content_slide = self.presentation.slides[slide_idx]
                
                # 替换三级标题
                self._replace_placeholder(content_slide, 'h2_0', self._add_pinyin(h2['title']))
                
                # 替换内容
                for j, content in enumerate(h2['content']):
                    if content['type'] == 'text':
                        placeholder = f'h3_{j}'
                        self._replace_placeholder(content_slide, placeholder, content['content'])
                    elif content['type'] == 'image':
                        placeholder = f'h4_{j}'
                        self._replace_placeholder(content_slide, placeholder, content['url'])
                    elif content['type'] == 'link':
                        placeholder = f'h4_{j}'
                        self._replace_placeholder(content_slide, placeholder, content['url'])
                
                slide_idx += 1
                h2_index += 1
                
                # 检查是否需要进入下一个章节
                if i < len(self.md_content['h1']) - 1:
                    # 简单判断：如果下一个三级标题内容较多，可能属于下一个章节
                    if h2_index < len(self.md_content['h2']) and len(self.md_content['h2'][h2_index]['content']) > 5:
                        break
    
    def generate(self):
        """生成PPT"""
        self._match_slides()
        self._generate_content()
        output_path = 'output.pptx'
        self.presentation.save(output_path)
        return output_path