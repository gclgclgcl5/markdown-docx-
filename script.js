document.addEventListener('DOMContentLoaded', () => {
    // 获取DOM元素
    const inputText = document.getElementById('input-text');
    const outputText = document.getElementById('output-text');
    const detectedFormatBadge = document.getElementById('detected-format');
    const outputFormatBadge = document.getElementById('output-format');
    const convertBtn = document.getElementById('convert-btn');
    const copyBtn = document.getElementById('copy-btn');
    const downloadBtn = document.getElementById('download-btn');
    const fileImport = document.getElementById('file-import');

    // 格式标识
    const FORMAT_MARKDOWN = 'markdown';
    const FORMAT_DOC = 'doc';

    // 当前检测到的输入格式
    let detectedInputFormat = null;
    // 当前的输出格式
    let currentOutputFormat = null;
    // 用于存储生成的docx文档对象
    let generatedDocx = null;

    // 检查docx库是否正确加载
    if (typeof docx === 'undefined') {
        console.error('错误：docx.js库未正确加载！');
    } else {
        console.log('docx.js库已成功加载，版本:', docx.Document ? 'v7.x' : '其他版本');
    }

    // 初始化Showdown转换器（用于Markdown<->HTML转换）
    const showdownConverter = new showdown.Converter({
        tables: true,
        tasklists: true,
        strikethrough: true,
        emoji: true
    });
    
    // 监听输入框内容变化，动态检测格式
    inputText.addEventListener('input', debounce(detectInputFormat, 300));
    
    // 文件导入处理
    fileImport.addEventListener('change', handleFileImport);
    
    // 当点击转换按钮时执行智能转换
    convertBtn.addEventListener('click', smartConvert);
    
    // 复制结果按钮
    copyBtn.addEventListener('click', () => {
        if (!outputText.value) {
            alert('没有可复制的内容');
            return;
        }
        
        outputText.select();
        document.execCommand('copy');
        
        // 显示复制成功提示
        const originalText = copyBtn.textContent;
        copyBtn.textContent = '复制成功!';
        setTimeout(() => {
            copyBtn.textContent = originalText;
        }, 2000);
    });
    
    // 下载文件按钮
    downloadBtn.addEventListener('click', () => {
        if (!outputText.value && !generatedDocx) {
            alert('没有可下载的内容');
            return;
        }
        
        if (currentOutputFormat === FORMAT_MARKDOWN) {
            // 下载为Markdown文件
            const blob = new Blob([outputText.value], { type: 'text/markdown' });
            saveAs(blob, 'converted.md');
        } else if (currentOutputFormat === FORMAT_DOC && generatedDocx) {
            try {
                // 下载为DOCX文件 - 使用docx.Packer打包
                saveDocxFile();
            } catch (error) {
                console.error('生成DOCX文件时出错:', error);
                alert('无法生成DOCX文件，改用HTML格式下载');
                const blob = new Blob([outputText.value], { type: 'text/html' });
                saveAs(blob, 'converted.html');
            }
        } else {
            // 备用方案：下载为HTML
            const blob = new Blob([outputText.value], { type: 'text/html' });
            saveAs(blob, 'converted.html');
        }
    });

    // 处理文件导入
    function handleFileImport(event) {
        const file = event.target.files[0];
        if (!file) return;
        
        const fileName = file.name.toLowerCase();
        
        // 根据文件扩展名自动检测格式
        if (fileName.endsWith('.md')) {
            // 导入Markdown文件
            readMarkdownFile(file);
        } else if (fileName.endsWith('.doc') || fileName.endsWith('.docx')) {
            // 导入DOC/DOCX文件
            readDocxFile(file);
        } else {
            alert('不支持的文件格式。请上传 .md、.doc 或 .docx 文件。');
        }
        
        // 重置文件输入框，允许导入同一文件
        fileImport.value = '';
    }
    
    // 读取Markdown文件
    function readMarkdownFile(file) {
        const reader = new FileReader();
        reader.onload = function(e) {
            // 读取文件内容并填充到输入区域
            inputText.value = e.target.result;
            // 检测并显示格式
            detectInputFormat();
        };
        reader.onerror = function() {
            alert('读取文件失败');
        };
        reader.readAsText(file);
    }
    
    // 读取DOC/DOCX文件
    function readDocxFile(file) {
        // 显示加载提示
        inputText.value = '正在读取文件，请稍候...';
        
        const reader = new FileReader();
        reader.onload = function(e) {
            const arrayBuffer = e.target.result;
            
            // 使用mammoth.js将DOCX转换为HTML
            mammoth.convertToHtml({arrayBuffer})
                .then(result => {
                    // 保存完整的HTML内容（用于转换）
                    const html = result.value;
                    
                    // 在输入区域存储文件名和完整HTML
                    inputText.value = `<!-- 已导入文档：${file.name} -->\n${html}`;
                    
                    // 检测并显示格式
                    detectInputFormat();
                    
                    // 如果有警告，显示在控制台
                    if (result.messages.length > 0) {
                        console.warn('文档转换警告:', result.messages);
                    }
                })
                .catch(error => {
                    console.error('文档读取失败:', error);
                    inputText.value = '无法读取文档: ' + error.message;
                });
        };
        reader.onerror = function() {
            inputText.value = '读取文件失败';
        };
        reader.readAsArrayBuffer(file);
    }
    
    // 智能转换函数 - 自动检测输入格式并转换为相应的输出格式
    async function smartConvert() {
        const input = inputText.value.trim();
        
        if (!input) {
            alert('请先输入内容或导入文件');
            return;
        }
        
        try {
            // 重新检测输入格式
            detectInputFormat();
            
            if (!detectedInputFormat) {
                alert('无法确定输入内容的格式，请检查内容或尝试重新导入文件。');
                return;
            }
            
            // 基于检测到的输入格式决定输出格式
            if (detectedInputFormat === FORMAT_MARKDOWN) {
                // Markdown -> DOC
                currentOutputFormat = FORMAT_DOC;
                outputFormatBadge.textContent = 'DOC格式';
                outputFormatBadge.className = 'format-badge format-doc';
                await markdownToDoc(input);
            } else if (detectedInputFormat === FORMAT_DOC) {
                // DOC -> Markdown
                currentOutputFormat = FORMAT_MARKDOWN;
                outputFormatBadge.textContent = 'Markdown格式';
                outputFormatBadge.className = 'format-badge format-markdown';
                await docToMarkdown(input);
            }
        } catch (error) {
            console.error('转换出错:', error);
            alert('转换过程中出现错误：' + error.message);
        }
    }
    
    // 检测输入格式
    function detectInputFormat() {
        const input = inputText.value.trim();
        
        if (!input) {
            detectedInputFormat = null;
            detectedFormatBadge.textContent = '';
            return;
        }
        
        // 检查是否为HTML (DOC导入后的格式)
        if (input.startsWith('<') && (input.includes('</') || input.includes('/>'))) {
            detectedInputFormat = FORMAT_DOC;
            detectedFormatBadge.textContent = '已检测为DOC格式';
            detectedFormatBadge.className = 'format-badge format-doc';
            return;
        }
        
        // 检查是否为Markdown
        // Markdown常见特征: 标题(#), 列表(- 或 1.)，代码块(```)等
        const markdownFeatures = [
            /^#+\s+.+$/m,  // 标题
            /^[*\-+]\s+.+$/m,  // 无序列表
            /^\d+\.\s+.+$/m,  // 有序列表
            /^>\s+.+$/m,  // 引用
            /^`{3}.+`{3}/ms,  // 代码块
            /\[.+\]\(.+\)/m,  // 链接
            /\*\*.+\*\*/m,  // 粗体
            /\*.+\*/m,  // 斜体
            /^(?:\|[^|]+)+\|$/m,  // 表格
        ];
        
        // 如果匹配多个Markdown特征，则视为Markdown
        const markdownScore = markdownFeatures.reduce((count, pattern) => {
            return pattern.test(input) ? count + 1 : count;
        }, 0);
        
        if (markdownScore >= 1) {
            detectedInputFormat = FORMAT_MARKDOWN;
            detectedFormatBadge.textContent = '已检测为Markdown格式';
            detectedFormatBadge.className = 'format-badge format-markdown';
        } else {
            // 默认作为DOC格式
            detectedInputFormat = FORMAT_DOC;
            detectedFormatBadge.textContent = '已检测为DOC格式';
            detectedFormatBadge.className = 'format-badge format-doc';
        }
    }
    
    // 保存DOCX文件
    function saveDocxFile() {
        if (!generatedDocx) {
            throw new Error('没有可用的DOCX文档');
        }
        
        // 使用docx.js的Packer将文档导出为Blob
        docx.Packer.toBlob(generatedDocx).then(blob => {
            // 使用FileSaver库保存文件
            saveAs(blob, 'converted.docx');
        }).catch(error => {
            console.error('导出DOCX文件时出错:', error);
            throw error;
        });
    }

    // Markdown 转 DOC
    async function markdownToDoc(markdown) {
        try {
            // 使用Showdown将Markdown转换为HTML (用于预览)
            const html = showdownConverter.makeHtml(markdown);
            
            // 设置预览内容
            outputText.value = '正在生成DOCX文件...\n\n以下是预览内容 (HTML格式):\n\n' + html;
            
            // 使用docx.js创建Word文档
            generatedDocx = await createSimpleDocxFromMarkdown(markdown);
            
            if (generatedDocx) {
                // 更新预览内容，通知用户DOCX文件已准备好
                outputText.value = '✅ DOCX文件已生成，点击"下载文件"按钮获取文件。\n\n以下是预览内容 (HTML格式):\n\n' + html;
            } else {
                throw new Error('无法创建DOCX文档');
            }
        } catch (error) {
            console.error('Markdown转Doc时出错:', error);
            throw new Error('Markdown转换到DOC格式失败');
        }
    }
    
    // DOC 转 Markdown
    async function docToMarkdown(docContent) {
        try {
            // 移除可能存在的注释信息（如导入文件名）
            let content = docContent.replace(/<!--.*?-->/s, '').trim();
            
            // 判断输入是否为HTML格式
            const isHtml = content.startsWith('<') && content.includes('</');
            
            let htmlContent;
            
            if (isHtml) {
                // 如果已经是HTML，直接使用
                htmlContent = content;
            } else {
                // 提示用户上传.docx文件
                outputText.value = '请注意：直接转换DOC文本内容不被支持。请使用左侧输入区域上方的"导入文件"按钮上传.doc或.docx文件';
                return;
            }
            
            // 从HTML提取body内容
            const bodyContent = extractBodyContent(htmlContent);
            
            // 使用自定义函数将HTML转换为Markdown
            const markdown = htmlToMarkdown(bodyContent);
            
            outputText.value = markdown;
            generatedDocx = null;
        } catch (error) {
            console.error('Doc转Markdown时出错:', error);
            throw new Error('DOC转换到Markdown格式失败');
        }
    }
    
    // 从HTML中提取body内容
    function extractBodyContent(html) {
        // 如果有body标签，提取其中的内容
        const bodyMatch = /<body[^>]*>([\s\S]*?)<\/body>/i.exec(html);
        if (bodyMatch) {
            return bodyMatch[1].trim();
        }
        
        // 如果没有body标签但有HTML文档结构，尝试提取内容部分
        const htmlMatch = /<html[^>]*>([\s\S]*?)<\/html>/i.exec(html);
        if (htmlMatch) {
            // 排除head部分
            const content = htmlMatch[1].replace(/<head[\s\S]*?<\/head>/i, '');
            return content.trim();
        }
        
        // 如果没有HTML和BODY标签，可能是部分HTML片段，直接返回
        return html;
    }
    
    // 简单的HTML转Markdown实现
    function htmlToMarkdown(html) {
        let markdown = html;
        
        // 替换标题
        markdown = markdown.replace(/<h1[^>]*>([\s\S]*?)<\/h1>/gi, '# $1\n\n');
        markdown = markdown.replace(/<h2[^>]*>([\s\S]*?)<\/h2>/gi, '## $1\n\n');
        markdown = markdown.replace(/<h3[^>]*>([\s\S]*?)<\/h3>/gi, '### $1\n\n');
        markdown = markdown.replace(/<h4[^>]*>([\s\S]*?)<\/h4>/gi, '#### $1\n\n');
        markdown = markdown.replace(/<h5[^>]*>([\s\S]*?)<\/h5>/gi, '##### $1\n\n');
        markdown = markdown.replace(/<h6[^>]*>([\s\S]*?)<\/h6>/gi, '###### $1\n\n');
        
        // 替换段落，确保非贪婪匹配
        markdown = markdown.replace(/<p[^>]*>([\s\S]*?)<\/p>/gi, '$1\n\n');
        
        // 替换粗体和斜体
        markdown = markdown.replace(/<strong[^>]*>([\s\S]*?)<\/strong>/gi, '**$1**');
        markdown = markdown.replace(/<b[^>]*>([\s\S]*?)<\/b>/gi, '**$1**');
        markdown = markdown.replace(/<em[^>]*>([\s\S]*?)<\/em>/gi, '*$1*');
        markdown = markdown.replace(/<i[^>]*>([\s\S]*?)<\/i>/gi, '*$1*');
        
        // 替换链接
        markdown = markdown.replace(/<a href="(.*?)"[^>]*>([\s\S]*?)<\/a>/gi, '[$2]($1)');
        
        // 替换图片
        markdown = markdown.replace(/<img src="(.*?)"[^>]*?(?:alt="(.*?)")?[^>]*>/gi, '![$2]($1)');
        
        // 替换列表 - 优化处理嵌套列表
        markdown = markdown.replace(/<ul[^>]*>([\s\S]*?)<\/ul>/gi, (match, content) => {
            return content.replace(/<li[^>]*>([\s\S]*?)<\/li>/gi, '- $1\n');
        });
        
        markdown = markdown.replace(/<ol[^>]*>([\s\S]*?)<\/ol>/gi, (match, content) => {
            let index = 1;
            return content.replace(/<li[^>]*>([\s\S]*?)<\/li>/gi, (match, item) => {
                return `${index++}. ${item}\n`;
            });
        });
        
        // 替换代码块
        markdown = markdown.replace(/<pre[^>]*><code[^>]*>([\s\S]*?)<\/code><\/pre>/gi, '```\n$1\n```\n\n');
        
        // 替换内联代码
        markdown = markdown.replace(/<code[^>]*>([\s\S]*?)<\/code>/gi, '`$1`');
        
        // 替换水平线
        markdown = markdown.replace(/<hr[^>]*>/gi, '---\n\n');
        
        // 替换引用块
        markdown = markdown.replace(/<blockquote[^>]*>([\s\S]*?)<\/blockquote>/gi, (match, content) => {
            // 为引用块内的每一行添加>前缀
            return content.split('\n').map(line => `> ${line}`).join('\n') + '\n\n';
        });
        
        // 简单替换表格（这是一个基本实现，复杂表格可能需要更多处理）
        markdown = markdown.replace(/<table[^>]*>([\s\S]*?)<\/table>/gi, (match, tableContent) => {
            let result = '';
            
            // 处理表头
            const headerMatch = /<thead[^>]*>([\s\S]*?)<\/thead>/i.exec(tableContent);
            if (headerMatch) {
                const headerContent = headerMatch[1];
                const headerCells = [];
                
                // 提取表头单元格
                let headerCellMatch;
                const thRegex = /<th[^>]*>([\s\S]*?)<\/th>/gi;
                while ((headerCellMatch = thRegex.exec(headerContent)) !== null) {
                    headerCells.push(headerCellMatch[1].trim());
                }
                
                if (headerCells.length > 0) {
                    // 创建表头行
                    result += `| ${headerCells.join(' | ')} |\n`;
                    
                    // 创建分隔行
                    result += `| ${headerCells.map(() => '---').join(' | ')} |\n`;
                }
            }
            
            // 处理表格主体
            const bodyMatch = /<tbody[^>]*>([\s\S]*?)<\/tbody>/i.exec(tableContent);
            if (bodyMatch) {
                const bodyContent = bodyMatch[1];
                const rows = [];
                
                // 提取表格行
                let rowMatch;
                const trRegex = /<tr[^>]*>([\s\S]*?)<\/tr>/gi;
                while ((rowMatch = trRegex.exec(bodyContent)) !== null) {
                    const rowContent = rowMatch[1];
                    const cells = [];
                    
                    // 提取单元格
                    let cellMatch;
                    const tdRegex = /<td[^>]*>([\s\S]*?)<\/td>/gi;
                    while ((cellMatch = tdRegex.exec(rowContent)) !== null) {
                        cells.push(cellMatch[1].trim());
                    }
                    
                    if (cells.length > 0) {
                        rows.push(`| ${cells.join(' | ')} |`);
                    }
                }
                
                if (rows.length > 0) {
                    result += rows.join('\n') + '\n\n';
                }
            }
            
            return result;
        });
        
        // 清理HTML标签前，处理一些特殊情况
        // 处理换行
        markdown = markdown.replace(/<br\s*\/?>/gi, '\n');
        markdown = markdown.replace(/&nbsp;/gi, ' ');
        
        // 处理div和span等容器标签
        markdown = markdown.replace(/<div[^>]*>([\s\S]*?)<\/div>/gi, '$1\n');
        markdown = markdown.replace(/<span[^>]*>([\s\S]*?)<\/span>/gi, '$1');
        
        // 清理剩余的HTML标签
        markdown = markdown.replace(/<[^>]+>/g, '');
        
        // 解码HTML实体
        markdown = markdown.replace(/&lt;/g, '<')
                          .replace(/&gt;/g, '>')
                          .replace(/&amp;/g, '&')
                          .replace(/&quot;/g, '"')
                          .replace(/&#39;/g, "'");
        
        // 修复多余的空行
        markdown = markdown.replace(/\n{3,}/g, '\n\n');
        
        return markdown;
    }
    
    // 使用docx.js创建一个简单的DOCX文档
    async function createSimpleDocxFromMarkdown(markdown) {
        try {
            console.log('开始创建DOCX文档');
            
            // 解析Markdown文本
            const lines = markdown.split('\n');
            const paragraphs = [];
            
            let i = 0;
            while (i < lines.length) {
                const line = lines[i].trim();
                
                // 跳过空行
                if (line === '') {
                    i++;
                    continue;
                }
                
                // 检查标题 (# 标题)
                const headingMatch = line.match(/^(#{1,6})\s+(.+)$/);
                if (headingMatch) {
                    const level = headingMatch[1].length;
                    const text = headingMatch[2].trim();
                    
                    paragraphs.push(
                        new docx.Paragraph({
                            text: text,
                            heading: `Heading${level}`,
                        })
                    );
                    i++;
                    continue;
                }
                
                // 检查水平线 (---, ***, ___)
                if (line.match(/^(\*{3,}|-{3,}|_{3,})$/)) {
                    paragraphs.push(
                        new docx.Paragraph({
                            border: {
                                bottom: {
                                    color: "auto",
                                    space: 1,
                                    style: "single",
                                    size: 6,
                                },
                            },
                        })
                    );
                    i++;
                    continue;
                }
                
                // 检查引用块 (> 引用文本)
                const blockquoteMatch = line.match(/^>\s+(.+)$/);
                if (blockquoteMatch) {
                    const text = blockquoteMatch[1];
                    paragraphs.push(
                        new docx.Paragraph({
                            text: text,
                            indent: {
                                left: 720, // 1/2 inch in twips
                            },
                            border: {
                                left: {
                                    color: "#CCCCCC",
                                    space: 10,
                                    style: "single",
                                    size: 10,
                                },
                            },
                        })
                    );
                    i++;
                    continue;
                }
                
                // 检查无序列表 (* 列表项, - 列表项, + 列表项)
                const ulMatch = line.match(/^[\*\-\+]\s+(.+)$/);
                if (ulMatch) {
                    const text = ulMatch[1];
                    paragraphs.push(
                        new docx.Paragraph({
                            text: text,
                            bullet: {
                                level: 0
                            }
                        })
                    );
                    i++;
                    continue;
                }
                
                // 检查有序列表 (1. 列表项, 2. 列表项...)
                const olMatch = line.match(/^\d+\.\s+(.+)$/);
                if (olMatch) {
                    const text = olMatch[1];
                    paragraphs.push(
                        new docx.Paragraph({
                            text: text,
                            numbering: {
                                reference: 1,
                                level: 0
                            }
                        })
                    );
                    i++;
                    continue;
                }
                
                // 检查代码块
                if (line.startsWith('```')) {
                    let codeContent = '';
                    i++; // 跳过开始的 ```
                    
                    // 收集代码块内容直到找到结束的 ```
                    while (i < lines.length && !lines[i].startsWith('```')) {
                        codeContent += lines[i] + '\n';
                        i++;
                    }
                    
                    if (i < lines.length) { // 找到结束的 ```
                        paragraphs.push(
                            new docx.Paragraph({
                                text: codeContent.trim(),
                                shading: {
                                    type: docx.ShadingType.SOLID,
                                    color: "F5F5F5",
                                },
                                font: {
                                    name: "Courier New"
                                }
                            })
                        );
                        i++; // 跳过结束的 ```
                    }
                    continue;
                }
                
                // 检查粗体和斜体
                if (line.includes('**') || line.includes('*') || 
                    line.includes('__') || line.includes('_')) {
                    
                    // 构建包含格式的段落
                    const textRuns = parseFormattedText(line);
                    paragraphs.push(
                        new docx.Paragraph({
                            children: textRuns
                        })
                    );
                } else {
                    // 普通段落
                    paragraphs.push(
                        new docx.Paragraph({
                            text: line
                        })
                    );
                }
                
                i++;
            }
            
            // 创建文档
            const doc = new docx.Document({
                numbering: {
                    config: [
                        {
                            reference: 1,
                            levels: [
                                {
                                    level: 0,
                                    format: "decimal",
                                    text: "%1.",
                                    alignment: "start",
                                    style: {
                                        paragraph: {
                                            indent: { left: 720, hanging: 260 }
                                        }
                                    }
                                }
                            ]
                        }
                    ]
                },
                sections: [
                    {
                        properties: {},
                        children: paragraphs
                    }
                ]
            });
            
            console.log('DOCX文档创建成功，包含' + paragraphs.length + '个段落');
            return doc;
        } catch (error) {
            console.error('创建DOCX文档时出错:', error);
            console.log('错误位置:', error.stack);
            console.log('docx.Document是否存在:', !!docx.Document);
            console.log('docx.Paragraph是否存在:', !!docx.Paragraph);
            console.log('docx.TextRun是否存在:', !!docx.TextRun);
            throw error;
        }
    }
    
    // 解析包含格式的文本（粗体、斜体等）
    function parseFormattedText(text) {
        const result = [];
        let currentIndex = 0;
        
        // 简单的粗体检测
        const boldRegex = /\*\*(.*?)\*\*|__(.*?)__/g;
        let boldMatch;
        
        // 重置正则表达式
        boldRegex.lastIndex = 0;
        
        // 找出所有粗体文本
        while ((boldMatch = boldRegex.exec(text)) !== null) {
            // 添加粗体之前的普通文本
            if (boldMatch.index > currentIndex) {
                result.push(
                    new docx.TextRun({
                        text: text.substring(currentIndex, boldMatch.index)
                    })
                );
            }
            
            // 添加粗体文本
            const boldText = boldMatch[1] || boldMatch[2];
            result.push(
                new docx.TextRun({
                    text: boldText,
                    bold: true
                })
            );
            
            currentIndex = boldMatch.index + boldMatch[0].length;
        }
        
        // 添加剩余的文本
        if (currentIndex < text.length) {
            result.push(
                new docx.TextRun({
                    text: text.substring(currentIndex)
                })
            );
        }
        
        // 如果没有找到任何格式，直接返回普通文本
        if (result.length === 0) {
            result.push(
                new docx.TextRun({
                    text: text
                })
            );
        }
        
        return result;
    }
    
    // 防抖函数，用于优化输入时的检测
    function debounce(func, wait) {
        let timeout;
        return function() {
            const context = this;
            const args = arguments;
            clearTimeout(timeout);
            timeout = setTimeout(() => func.apply(context, args), wait);
        };
    }
    
    // 初始化时检测一次格式
    detectInputFormat();
}); 