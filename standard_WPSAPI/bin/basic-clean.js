const fs = require('fs');
const path = require('path');

// 定义源目录和输出文件路径
const docsDir = path.join(__dirname, '..', 'docs');
const outputFile = path.join(__dirname, '..', 'wps_api_documentation.txt');

// 读取目录中的所有HTML文件
fs.readdir(docsDir, (err, files) => {
  if (err) {
    console.error('Error reading docs directory:', err);
    return;
  }

  // 过滤出HTML文件
  const htmlFiles = files.filter(file => path.extname(file) === '.html');
  
  // 创建输出流
  const writeStream = fs.createWriteStream(outputFile);
  
  writeStream.write('# WPS API Documentation\n\n');
  
  let processedCount = 0;
  
  // 处理每个HTML文件
  htmlFiles.forEach((file, index) => {
    const filePath = path.join(docsDir, file);
    
    // 读取文件内容
    fs.readFile(filePath, 'utf8', (err, data) => {
      if (err) {
        console.error(`Error reading file ${file}:`, err);
        return;
      }
      
      try {
        // 写入文件标题
        writeStream.write(`\n\n========================================\n`);
        writeStream.write(`FILE: ${file}\n`);
        writeStream.write(`========================================\n\n`);
        
        // 简单提取body中的内容
        let content = data;
        
        // 尝试提取<body>标签中的内容
        const bodyStart = content.indexOf('<body');
        const bodyEnd = content.lastIndexOf('</body>');
        
        if (bodyStart !== -1 && bodyEnd !== -1) {
          content = content.substring(bodyStart, bodyEnd + 7);
        }
        
        // 移除HTML标签
        content = content.replace(/<[^>]*>/g, '');
        
        // 解码常见的HTML实体
        content = content.replace(/&nbsp;/g, ' ');
        content = content.replace(/&amp;/g, '&');
        content = content.replace(/&lt;/g, '<');
        content = content.replace(/&gt;/g, '>');
        content = content.replace(/&quot;/g, '"');
        content = content.replace(/&#39;/g, "'");
        
        // 清理多余的空白字符
        let lines = content.split('\n');
        let cleanedLines = [];
        for (let i = 0; i < lines.length; i++) {
          let line = lines[i].trim();
          if (line !== '') {
            cleanedLines.push(line);
          } else if (cleanedLines[cleanedLines.length - 1] !== '') {
            // 只保留一个空行作为分隔
            cleanedLines.push('');
          }
        }
        
        // 删除开头和结尾的空行
        while (cleanedLines.length > 0 && cleanedLines[0] === '') {
          cleanedLines.shift();
        }
        while (cleanedLines.length > 0 && cleanedLines[cleanedLines.length - 1] === '') {
          cleanedLines.pop();
        }
        
        content = cleanedLines.join('\n');
        
        // 写入处理后的内容
        writeStream.write(content);
        
        processedCount++;
        console.log(`Processed ${processedCount}/${htmlFiles.length}: ${file}`);
        
        // 如果是最后一个文件，关闭写入流
        if (processedCount === htmlFiles.length) {
          writeStream.end(() => {
            console.log(`\nProcessing complete! Output written to: ${outputFile}`);
          });
        }
      } catch (parseError) {
        console.error(`Error parsing HTML in file ${file}:`, parseError);
      }
    });
  });
});