const express = require('express');
const fs = require('fs');
const path = require('path');

const app = express();
const PORT = 3000;

// 提供静态文件服务
app.use(express.static(path.join(__dirname, '../docs')));

// 主页路由
app.get('/', (req, res) => {
  res.send(`
    <!DOCTYPE html>
    <html>
    <head>
        <title>WPS API Documentation Viewer</title>
        <meta charset="utf-8">
        <style>
            body { font-family: Arial, sans-serif; margin: 0; padding: 20px; }
            .container { display: flex; height: 100vh; }
            .sidebar { width: 250px; border-right: 1px solid #ccc; padding: 10px; overflow-y: auto; }
            .content { flex: 1; padding: 10px; overflow-y: auto; }
            .sidebar a { display: block; padding: 5px 0; text-decoration: none; color: #333; }
            .sidebar a:hover { text-decoration: underline; }
            iframe { width: 100%; height: 95%; border: 1px solid #ddd; }
        </style>
    </head>
    <body>
        <h1>WPS API Documentation Viewer</h1>
        <div class="container">
            <div class="sidebar">
                <h3>Documentation Files</h3>
                ${generateFileList()}
            </div>
            <div class="content">
                <h2>Content</h2>
                <iframe id="docFrame" src=""></iframe>
            </div>
        </div>
        <script>
            function loadDoc(filename) {
                document.getElementById('docFrame').src = filename;
            }
            
            // 默认加载第一个文档
            window.onload = function() {
                const firstLink = document.querySelector('.sidebar a');
                if (firstLink) {
                    loadDoc(firstLink.getAttribute('onclick').match(/'([^']+)'/)[1]);
                }
            }
        </script>
    </body>
    </html>
  `);
});

// 生成文件列表
function generateFileList() {
  const docsDir = path.join(__dirname, '../docs');
  if (!fs.existsSync(docsDir)) {
    return '<p>No documentation files found. Please run the scraper first.</p>';
  }
  
  const files = fs.readdirSync(docsDir);
  let html = '';
  
  files.forEach(file => {
    if (file.endsWith('.html')) {
      html += `<a href="#" onclick="loadDoc('${file}'); return false;">${file}</a><br>`;
    }
  });
  
  return html || '<p>No HTML files found.</p>';
}

app.listen(PORT, () => {
  console.log(`Documentation viewer running at http://localhost:${PORT}`);
});