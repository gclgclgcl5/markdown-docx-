document.addEventListener('DOMContentLoaded', () => {
    // 获取DOM元素
    const inputText = document.getElementById('input-text');
    const outputText = document.getElementById('output-text');
    const inputFormat = document.getElementById('input-format');
    const outputFormat = document.getElementById('output-format');
    const convertBtn = document.getElementById('convert-btn');
    const swapBtn = document.getElementById('swap-btn');
    const copyBtn = document.getElementById('copy-btn');
    const downloadBtn = document.getElementById('download-btn');

    // 检查docx库是否正确加载
    if (typeof docx === 'undefined') {
        console.error('错误：docx.js库未正确加载！');
    } else {
        console.log('docx.js库已成功加载，版本:', docx.Document ? 'v7.x' : '其他版本');
        // 输出docx对象的所有顶级属性，帮助调试
        console.log('docx库包含的对象:', Object.keys(docx));
    }

    // 初始化Showdown转换器（用于Markdown<->HTML转换）
    const showdownConverter = new showdown.Converter({
        tables: true,
        tasklists: true,
        strikethrough: true,
        emoji: true
    });
    
    // 用于存储生成的docx文档对象
    let generatedDocx = null;
    
    // 当点击转换按钮时
    convertBtn.addEventListener('click', async () => {
        const input = inputText.value.trim();
        
        if (!input) {
            alert('请先输入内容');
            return;
        }
        
        try {
            // 基于选择的格式进行转换
            if (inputFormat.value === 'markdown' && outputFormat.value === 'doc') {
                // Markdown -> DOC (DOCX)
                await markdownToDoc(input);
            } else if (inputFormat.value === 'doc' && outputFormat.value === 'markdown') {
                // DOC -> Markdown
                await docToMarkdown(input);
            } else {
                // 相同格式，无需转换
                outputText.value = input;
                generatedDocx = null;
            }
        } catch (error) {
            console.error('转换出错:', error);
            alert('转换过程中出现错误：' + error.message);
        }
    });
    
    // 切换输入和输出格式
    swapBtn.addEventListener('click', () => {
        // 交换两个选择框的值
        const tempValue = inputFormat.value;
        inputFormat.value = outputFormat.value;
        outputFormat.value = tempValue;
        
        // 清空输出区域
        outputText.value = '';
        generatedDocx = null;
    });
    
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
        
        if (outputFormat.value === 'markdown') {
            // 下载为Markdown文件
            const blob = new Blob([outputText.value], { type: 'text/markdown' });
            saveAs(blob, 'converted.md');
        } else if (outputFormat.value === 'doc' && generatedDocx) {
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

    // Markdown 转 DOC (真正的DOCX文件)
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
    
    // DOC 转 Markdown
    async function docToMarkdown(docContent) {
        try {
            // 判断输入是否为HTML格式
            const isHtml = docContent.trim().startsWith('<') && docContent.includes('</');
            
            let htmlContent;
            
            if (isHtml) {
                // 如果已经是HTML，直接使用
                htmlContent = docContent;
            } else {
                // 提示用户上传.docx文件
                outputText.value = '请注意：直接转换DOC文本内容不被支持。请上传.docx文件';
                setupFileUploader();
                return;
            }
            
            // 从HTML提取body内容
            const bodyContent = extractBodyContent(htmlContent);
            
            // 使用TurndownService将HTML转换为Markdown
            // 由于不能直接使用turndown.js库，我们会实现一个简单的HTML到Markdown的转换
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
        const bodyMatch = /<body[^>]*>([\s\S]*)<\/body>/i.exec(html);
        return bodyMatch ? bodyMatch[1].trim() : html;
    }
    
    // 简单的HTML转Markdown实现
    function htmlToMarkdown(html) {
        let markdown = html;
        
        // 替换标题
        markdown = markdown.replace(/<h1[^>]*>(.*?)<\/h1>/gi, '# $1\n\n');
        markdown = markdown.replace(/<h2[^>]*>(.*?)<\/h2>/gi, '## $1\n\n');
        markdown = markdown.replace(/<h3[^>]*>(.*?)<\/h3>/gi, '### $1\n\n');
        markdown = markdown.replace(/<h4[^>]*>(.*?)<\/h4>/gi, '#### $1\n\n');
        markdown = markdown.replace(/<h5[^>]*>(.*?)<\/h5>/gi, '##### $1\n\n');
        markdown = markdown.replace(/<h6[^>]*>(.*?)<\/h6>/gi, '###### $1\n\n');
        
        // 替换段落
        markdown = markdown.replace(/<p[^>]*>(.*?)<\/p>/gi, '$1\n\n');
        
        // 替换粗体和斜体
        markdown = markdown.replace(/<strong[^>]*>(.*?)<\/strong>/gi, '**$1**');
        markdown = markdown.replace(/<b[^>]*>(.*?)<\/b>/gi, '**$1**');
        markdown = markdown.replace(/<em[^>]*>(.*?)<\/em>/gi, '*$1*');
        markdown = markdown.replace(/<i[^>]*>(.*?)<\/i>/gi, '*$1*');
        
        // 替换链接
        markdown = markdown.replace(/<a href="(.*?)"[^>]*>(.*?)<\/a>/gi, '[$2]($1)');
        
        // 替换图片
        markdown = markdown.replace(/<img src="(.*?)"[^>]*>/gi, '![]($1)');
        
        // 替换列表
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
        markdown = markdown.replace(/<code[^>]*>(.*?)<\/code>/gi, '`$1`');
        
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
        
        // 清理HTML标签
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
    
    // 设置文件上传处理器（用于DOCX文件）
    function setupFileUploader() {
        // 检查是否已经存在上传器
        let uploader = document.getElementById('docx-uploader');
        if (uploader) {
            return;
        }
        
        // 创建文件上传元素
        uploader = document.createElement('div');
        uploader.id = 'docx-uploader';
        uploader.style.marginTop = '10px';
        uploader.innerHTML = `
            <p>如需转换DOCX文件，请上传:</p>
            <input type="file" id="docx-file" accept=".docx">
        `;
        
        // 插入到输出区域上方
        outputText.parentNode.insertBefore(uploader, outputText);
        
        // 监听文件上传
        document.getElementById('docx-file').addEventListener('change', handleDocxUpload);
    }
    
    // 处理DOCX文件上传
    function handleDocxUpload(event) {
        const file = event.target.files[0];
        if (!file) return;
        
        const reader = new FileReader();
        reader.onload = function(loadEvent) {
            const arrayBuffer = loadEvent.target.result;
            
            // 使用mammoth.js将DOCX转换为HTML
            mammoth.convertToHtml({arrayBuffer})
                .then(result => {
                    const html = result.value;
                    const markdown = htmlToMarkdown(html);
                    outputText.value = markdown;
                    generatedDocx = null;
                    
                    // 移除上传器
                    const uploader = document.getElementById('docx-uploader');
                    if (uploader) {
                        uploader.parentNode.removeChild(uploader);
                    }
                })
                .catch(error => {
                    console.error('DOCX转换失败:', error);
                    outputText.value = '无法转换DOCX文件: ' + error.message;
                });
        };
        reader.readAsArrayBuffer(file);
    }
}); 