const fs = require('fs');

// 读取原始文件
const content = fs.readFileSync('process-docs.js', 'utf8');

// 按行分割
const lines = content.split('\n');

// 找到需要替换的行范围
let startIndex = -1;
let endIndex = -1;

for (let i = 0; i < lines.length; i++) {
  if (lines[i].includes('content = content.replace(/')) {
    startIndex = i;
  }
  if (lines[i].includes("'); // 将多个连续空行替换为单个空行")) {
    endIndex = i;
    break;
  }
}

console.log(`找到要替换的行: ${startIndex}-${endIndex}`);

// 构造修复后的内容
const fixedLines = [...lines];
fixedLines.splice(startIndex, endIndex - startIndex + 1, "        content = content.replace(/\\n\\s*\\n\\s*/g, '\\n\\n'); // 将多个连续空行替换为单个空行");

// 合并为完整内容
const fixedContent = fixedLines.join('\n');

// 写入修复后的文件
fs.writeFileSync('process-docs-fixed.js', fixedContent);

console.log('文件已修复并保存为 process-docs-fixed.js');