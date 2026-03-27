# Markdown转PPT项目 - 项目开发文档

> 更新日期：2026-03-27
> 基于代码版本：V16 (ppt_generator.py)
> 模板版本：3.11 模板.pptx

## 项目概述

本项目实现将 Markdown 文件自动转换为 PPT 演示文稿，主要面向 Echineselearning 对外汉语教学场景。核心能力包括：
- 汉字自动拼音标注（精确对齐汉字上方）
- 三维模板匹配（文本框数 × 图片数 × 媒体类型）
- 音视频嵌入（MP3/MP4，含视频封面提取）
- 网页链接嵌入（点击用浏览器打开）
- 图片自动替换（在线URL下载 + 本地路径）
- 混合字体处理（中文 SimHei / 英文 Arial）

## 核心功能

| 功能 | 描述 | 状态 |
|------|------|------|
| Markdown 解析 | 解析3级标题，提取文本/图片/音视频/链接 | ✅ |
| 拼音标注 | 2×N透明表格，拼音精确对齐汉字正上方 | ✅ |
| 三维模板匹配 | 按 (文本框数, 图片数, 媒体类型) 索引模板 | ✅ |
| XML深拷贝克隆 | 保留模板背景、装饰，先克隆后填充 | ✅ |
| 图片嵌入 | 在线URL下载 + 本地路径，按位置替换模板图片 | ✅ |
| 音频嵌入 | MP3嵌入PPT，add_movie方式，图标显示可播放 | ✅ |
| 视频嵌入 | MP4嵌入PPT，ffmpeg提取第一帧作为封面 | ✅ |
| 链接嵌入 | {{link}}占位符，click_action.hyperlink跳转 | ✅ |
| 混合字体 | 中文SimHei + 英文Arial，按run分别设置 | ✅ |
| 自动字号调整 | 根据文本框数量和内容量自动选择字号 | ✅ |
| 多行拼音 | 支持换行符，每行独立创建拼音表格 | ✅ |

## 文档目录

| 文档 | 说明 |
|------|------|
| [开发需求.md](./开发需求.md) | 15项需求详细说明 |
| [开发状态.md](./开发状态.md) | 版本演进 + 完成情况 |
| [项目结构.md](./项目结构.md) | 目录结构 + 模块说明 |
| [代码说明.md](./代码说明.md) | 核心代码详解 + 关键算法 |
| [MD撰写规范.md](./MD撰写规范.md) | MD文件撰写语法规范 |
| [PPT模板编写规则.md](./PPT模板编写规则.md) | 模板占位符命名 + 页面类型 |
| [匹配规则说明.md](./匹配规则说明.md) | 内容→模板的匹配逻辑 |

## 快速开始

### 安装依赖
```bash
pip install python-pptx pypinyin lxml flask
# 系统依赖（视频封面提取）
# Ubuntu/Debian: apt install ffmpeg
# macOS: brew install ffmpeg
```

### 命令行调用
```bash
python src/utils/ppt_generator.py 案例/3.11测试—说白色.md 案例/3.11模板.pptx
```

### 代码调用
```python
from src.utils.md_parser import MDParser
from src.utils.ppt_generator import PPTGenerator

parser = MDParser('案例/3.11测试—说白色.md')
md_content = parser.parse()

generator = PPTGenerator('案例/3.11模板.pptx', md_content)
output = generator.generate()
print(f'生成完成: {output}')
```

### Flask API
```bash
python app.py
# POST /api/convert
# Form data: md_file, ppt_template
```

## 技术栈

| 依赖 | 用途 |
|------|------|
| Python 3.7+ | 运行环境 |
| python-pptx | PPT 创建和操作 |
| pypinyin | 汉字转拼音 |
| lxml | XML 底层操作 |
| flask | Web API（可选） |
| ffmpeg | 视频封面提取（系统依赖，可选） |

## 仓库地址

https://github.com/zhoujiaxinii/Markdown-to-PPT-AI_ECL
