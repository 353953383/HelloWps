const fs = require('fs');
const path = require('path');
const { JSDOM } = require('jsdom');

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
        // 使用JSDOM解析HTML
        const dom = new JSDOM(data);
        const document = dom.window.document;
        
        // 提取标题
        const title = document.querySelector('title') ? document.querySelector('title').textContent : file;
        
        // 写入文件标题
        writeStream.write(`\n\n========================================\n`);
        writeStream.write(`FILE: ${file}\n`);
        writeStream.write(`TITLE: ${title}\n`);
        writeStream.write(`========================================\n\n`);
        
        // 移除脚本和样式标签
        const scripts = document.querySelectorAll('script');
        scripts.forEach(script => script.remove());
        
        const styles = document.querySelectorAll('style');
        styles.forEach(style => style.remove());
        
        // 移除导航栏和其他无关元素
        const navElements = document.querySelectorAll('nav, .navbar, .sidebar, .toc, .header, .footer');
        navElements.forEach(el => el.remove());
        
        // 获取主要内容
        let content = '';
        
        // 尝试找到主要内容区域
        const mainContent = document.querySelector('main, .content, #content, .main-content, article') || document.body;
        
        // 递归提取文本内容，同时保持一定的结构
        function extractTextWithStructure(element, depth = 0) {
          let text = '';
          
          // 跳过隐藏元素和空元素
          if (element.style && element.style.display === 'none' || 
              element.hidden || 
              (element.textContent && element.textContent.trim() === '')) {
            return text;
          }
          
          // 处理不同类型的元素
          if (element.nodeType === dom.window.Node.TEXT_NODE) {
            const trimmedText = element.textContent.trim();
            if (trimmedText) {
              text += trimmedText + ' ';
            }
          } else if (element.nodeType === dom.window.Node.ELEMENT_NODE) {
            // 添加换行以保持结构（根据元素类型）
            const blockElements = ['div', 'p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'li', 'tr', 'td', 'br'];
            if (blockElements.includes(element.tagName.toLowerCase())) {
              text += '\n';
            }
            
            // 处理标题元素
            if (element.tagName.match(/^H[1-6]$/)) {
              text += '\n' + '='.repeat(40) + '\n';
            }
            
            // 递归处理子节点
            for (let i = 0; i < element.childNodes.length; i++) {
              text += extractTextWithStructure(element.childNodes[i], depth + 1);
            }
            
            // 在块级元素后添加换行
            if (blockElements.includes(element.tagName.toLowerCase())) {
              text += '\n';
            }
            
            // 在标题元素后添加分隔线
            if (element.tagName.match(/^H[1-6]$/)) {
              text += '='.repeat(40) + '\n';
            }
          }
          
          return text;
        }
        
        content = extractTextWithStructure(mainContent);
        
        // 清理多余的空白字符
        content = content.replace(/
\s*
\s*
/g, '

'); // 将多个连续空行替换为单个空行
        content = content.replace(/ {2,}/g, ' '); // 将多个连续空格替换为单个空格
        
        // 写入处理后的内容
        writeStream.write(content.trim());
        
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