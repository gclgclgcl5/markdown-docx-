# Markdown与DOC格式互转工具

这是一个基于Web的工具，可以将Markdown格式文本和DOC/DOCX格式文档进行互相转换。

## 最新更新

- 添加了真正的DOCX格式输出功能！现在可以直接生成Word文档(.docx)文件，而不只是HTML。

## 功能特点

- Markdown转换为DOC格式（真正的.docx文件）
- DOC/DOCX文档转换为Markdown格式
- 复制转换结果到剪贴板
- 下载转换后的文件
- 简洁美观的用户界面
- 完全在浏览器中运行，无需服务器

## 使用方法

1. 打开`index.html`文件
2. 在左侧输入区域粘贴您的Markdown或DOC文本
3. 选择对应的输入格式
4. 点击"转换"按钮生成转换后的格式
5. 使用"复制结果"或"下载文件"获取转换后的内容
6. 当转换Markdown到DOC时，点击"下载文件"按钮获取真正的Word文档(.docx)文件

## 技术说明

本工具使用以下技术：

- HTML5 + CSS3：构建用户界面
- JavaScript：实现转换逻辑
- Showdown.js：将Markdown转换为HTML
- docx.js：将Markdown转换为Word文档(.docx)格式
- Mammoth.js：将DOCX文档转换为HTML
- FileSaver.js：保存生成的文件

## 支持的Markdown格式

- 标题（# 到 ######）
- 段落
- 粗体（** ** 或 __ __）
- 列表（有序和无序）
- 引用（>）
- 代码块（```）
- 水平线（---）

## 限制说明

- DOC转Markdown功能需要上传.docx文件才能正常工作
- 对于复杂的文档格式（如嵌套表格、复杂公式等），转换结果可能不完美
- 一些Word高级功能（如注释、修订、宏等）在转换过程中会丢失
- 目前的行内格式支持有限，主要支持粗体

## 文件结构

- `index.html`：主页面HTML
- `styles.css`：样式表文件
- `script.js`：JavaScript功能实现

## 开始使用

直接在浏览器中打开`index.html`文件即可使用本工具。 