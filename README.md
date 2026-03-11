# Markdown转PPT项目

## 项目功能
- 将Markdown文件转换为PPT文件
- 支持从模板库中匹配合适的模板页（二维匹配：文本框数×图片数）
- 自动为汉字添加拼音标注（两行表格：拼音+汉字）
- 支持处理图片、音频、视频和网页链接
- 英文文本保留空格，纯英文行合并单元格
- 支持多种字号规则：
  - 1个文本框：40pt
  - 2个文本框：36pt
  - 3个文本框：30pt
  - 4个文本框：24pt
  - 1-2个文本框+无图片+中文>10字：28pt

## 项目结构
```
markdown-to-ppt/
├── src/
│   └── utils/
│       ├── md_parser.py      # Markdown解析器
│       └── ppt_generator.py  # PPT生成器
├── 案例/                      # 测试案例
├── test/                     # 测试文件
├── app.py                    # 主入口文件
├── requirements.txt          # 依赖文件
└── README.md                 # 项目说明
```

## 安装依赖
```bash
pip install -r requirements.txt
```

## 使用方法

### 1. 启动API服务
```bash
python app.py
```

### 2. 调用API接口

**API地址**: `http://localhost:5000/api/convert`

**请求方法**: POST

**请求参数**:
- `md_file`: Markdown文件（.md格式）
- `ppt_template`: PPT模板文件（.pptx格式）

**返回结果**:
```json
{
  "success": true,
  "output_path": "output.pptx"
}
```

### 3. 直接运行测试
```bash
python test_ppt_generator.py
```

## Markdown语法

### 标题
```markdown
# 封面标题
## 章节标题
### 正文标题
```

### 文本框（多文本框用 || 分隔）
```markdown
文本框1内容||文本框2内容
```

### 图片
```markdown
![图片描述](图片URL)
```

### 音频
```markdown
https://example.com/audio.mp3
```

### 视频
```markdown
https://example.com/video.mp4
```

### 网页链接
```markdown
https://example.com/link
```

## 模板要求

### 占位符
- 封面页：`{{h0_0}}`
- 章节页：`{{h1_0}}`
- 正文标题：`{{h2_0}}`
- 正文内容：`{{h3_0}}`, `{{h3_1}}`, ...（多个文本框）
- 图片：`{{pic_0}}`, `{{pic_1}}`, ...
- 音频：`{{audio}}`
- 视频：`{{video}}`
- 链接：`{{link}}`

### 模板匹配规则
- 按 (文本框数, 图片数, 媒体类型) 三维索引匹配
- 媒体类型优先级：video > audio > link > None
- 如果没有匹配的媒体类型模板，使用无媒体模板

## 字号规则

| 文本框数 | 有图片 | 中文字数 | 字号 |
|---------|--------|---------|------|
| 1 | 否 | ≤10 | 40pt |
| 1 | 否 | >10 | 28pt |
| 2 | 否 | ≤10 | 36pt |
| 2 | 否 | >10 | 28pt |
| 2 | 是 | - | 36pt |
| 3 | - | - | 30pt |
| 4 | - | - | 24pt |

## 注意事项
1. 仅支持.md/.docx格式的Markdown文件
2. 仅支持.pptx格式的PPT模板
3. 模板中必须包含指定的占位符
4. 生成的PPT会保持模板的原有样式
5. 汉字会自动添加拼音标注（放入表格）
6. 英文文本保留空格，纯英文行合并单元格
7. 音频嵌入后显示图标，点击可播放

## 示例

### 输入
- Markdown文件：包含标题、文本和多媒体链接
- PPT模板：包含占位符的模板文件

### 输出
- 生成的PPT文件：包含匹配的模板页和填充的内容

## 技术栈
- Python 3.7+
- Flask：Web API框架
- python-pptx：PPT处理库
- markdown：Markdown解析库
- pypinyin：拼音标注库

## 更新日志

### 2026-03-10
- 支持 {{link}} 占位符，点击后用浏览器打开
- 英文文本放入表格，纯英文行合并单元格
- 1-2个文本框+无图片+中文>10字时字号设为28pt
- 音频嵌入显示图标，点击可播放
- 视频封面提取
- 拼音表格支持占位符对齐方式