const fs = require('fs');
const path = require('path');
const puppeteer = require('puppeteer');

// 创建必要的目录
const docsDir = path.join(__dirname, '../docs');
if (!fs.existsSync(docsDir)) {
  fs.mkdirSync(docsDir, { recursive: true });
}

async function scrapePage(browser, url, filename) {
  try {
    console.log(`Scraping ${url}...`);
    const page = await browser.newPage();
    await page.goto(url, { waitUntil: 'networkidle2' });
    
    // 等待页面内容加载
    await new Promise(resolve => setTimeout(resolve, 3000));
    
    // 获取页面内容
    const html = await page.content();
    
    // 保存原始HTML
    const filePath = path.join(docsDir, filename);
    fs.writeFileSync(filePath, html);
    console.log(`Saved ${filename}`);
    
    await page.close();
    return html;
  } catch (error) {
    console.error(`Error scraping ${url}:`, error.message);
    return null;
  }
}

async function extractLinksFromPage(page) {
  try {
    // 等待一段时间确保JavaScript执行完毕
    await new Promise(resolve => setTimeout(resolve, 2000));
    
    // 提取页面上所有的链接
    const links = await page.evaluate(() => {
      // 尝试多种方式提取链接
      const allLinks = [];
      
      // 方法1: 传统的href链接
      const anchors = Array.from(document.querySelectorAll('a[href]'));
      anchors.forEach(anchor => {
        const href = anchor.href;
        const text = anchor.textContent.trim();
        if (href) {
          allLinks.push({ href, text });
        }
      });
      
      // 方法2: 查找可能是导航菜单的部分
      const navElements = document.querySelectorAll('.nav, .menu, .sidebar, .toc, nav, ul, ol');
      navElements.forEach(el => {
        const navAnchors = el.querySelectorAll('a[href]');
        navAnchors.forEach(anchor => {
          const href = anchor.href;
          const text = anchor.textContent.trim();
          if (href) {
            allLinks.push({ href, text });
          }
        });
      });
      
      // 方法3: 查找包含特定关键词的元素
      const clickableElements = document.querySelectorAll('[onclick], [data-url], [href], .clickable');
      clickableElements.forEach(el => {
        const href = el.href || el.getAttribute('data-url') || '';
        const text = el.textContent ? el.textContent.trim() : '';
        if (href && href.includes('docs')) {
          allLinks.push({ href, text });
        }
      });
      
      // 方法4: 检查所有具有特定类名或ID的元素
      const docElements = document.querySelectorAll('[id*="doc"], [class*="doc"], [id*="api"], [class*="api"]');
      docElements.forEach(el => {
        const linksInElement = el.querySelectorAll('a');
        linksInElement.forEach(anchor => {
          const href = anchor.href;
          const text = anchor.textContent.trim();
          if (href) {
            allLinks.push({ href, text });
          }
        });
      });

      // 方法5: 特别针对API参考文档的链接
      const apiReferenceSelectors = [
        'a[href*="addin-api-reference"]',
        'a[href*="macro-editor-api-reference"]',
        'a[href*="spreadsheet-api-reference"]',
        'a[href*="writer-api-reference"]',
        'a[href*="presentation-api-reference"]',
        'a[href*="common-api-reference"]',
        'a[href*="AboveAverage"]',
        'a[href*="AddIn"]',
        'a[href*="AddIns"]',
        'a[href*="AddIns2"]',
        'a[href*="Adjustments"]',
        'a[href*="AllowEditRange"]',
        'a[href*="AllowEditRanges"]',
        'a[href*="Application"]',
        'a[href*="Areas"]',
        'a[href*="AutoFilter"]',
        'a[href*="AutoRecover"]',
        'a[href*="Axes"]',
        'a[href*="Axis"]',
        'a[href*="AxisTitle"]',
        'a[href*="Border"]',
        'a[href*="Borders"]',
        'a[href*="CalculatedFields"]',
        'a[href*="CalculatedItems"]',
        'a[href*="CategoryCollection"]',
        'a[href*="CellFormat"]',
        'a[href*="Characters"]',
        'a[href*="Chart"]',
        'a[href*="ChartArea"]',
        'a[href*="ChartCategory"]',
        'a[href*="ChartFormat"]'
      ];
      
      apiReferenceSelectors.forEach(selector => {
        const elements = document.querySelectorAll(selector);
        elements.forEach(el => {
          const href = el.href || (el.getAttribute('href') ? new URL(el.getAttribute('href'), window.location.origin).href : '');
          const text = el.textContent ? el.textContent.trim() : '';
          if (href) {
            allLinks.push({ href, text });
          }
        });
      });

      return allLinks.filter(link => link.href && 
               (link.href.startsWith('https://open.wps.cn/') || 
                link.href.includes('wpsLoad') ||
                link.href.includes('docs') ||
                link.href.includes('previous') ||
                link.href.includes('AboveAverage') ||
                link.href.includes('AddIn') ||
                link.href.includes('AddIns') ||
                link.href.includes('AddIns2') ||
                link.href.includes('Adjustments') ||
                link.href.includes('AllowEditRange') ||
                link.href.includes('AllowEditRanges') ||
                link.href.includes('Application') ||
                link.href.includes('Areas') ||
                link.href.includes('AutoFilter') ||
                link.href.includes('AutoRecover') ||
                link.href.includes('Axes') ||
                link.href.includes('Axis') ||
                link.href.includes('AxisTitle') ||
                link.href.includes('Border') ||
                link.href.includes('Borders') ||
                link.href.includes('CalculatedFields') ||
                link.href.includes('CalculatedItems') ||
                link.href.includes('CategoryCollection') ||
                link.href.includes('CellFormat') ||
                link.href.includes('Characters') ||
                link.href.includes('Chart') ||
                link.href.includes('ChartArea') ||
                link.href.includes('ChartCategory') ||
                link.href.includes('ChartFormat')));
    });
    
    return links;
  } catch (error) {
    console.error(`Error extracting links:`, error.message);
    return [];
  }
}

async function scrapeSpecificPages(browser) {
  // 根据用户提供的信息，添加特定的URL
  const specificPages = [
    'https://open.wps.cn/previous/docs/client/wpsLoad',
    'https://open.wps.cn/previous/docs/client/wpsLoad/quick-start',
    'https://open.wps.cn/previous/docs/client/wpsLoad/client-development',
    'https://open.wps.cn/previous/docs/client/wpsLoad/server-development',
    'https://open.wps.cn/previous/docs/client/wpsLoad/faq',
    'https://open.wps.cn/previous/docs/client/wpsLoad/wps-addin',
    'https://open.wps.cn/previous/docs/client/wpsLoad/mobile-jsapi',
    'https://open.wps.cn/previous/docs/client/wpsLoad/payment',
    'https://open.wps.cn/previous/docs/client/wpsLoad/document-components',
    'https://open.wps.cn/previous/docs/client/wpsLoad/resources-download',
    'https://open.wps.cn/previous/docs/client/wpsLoad/app-logo',
    'https://open.wps.cn/previous/docs/client/addin-api-reference',
    'https://open.wps.cn/previous/docs/client/macro-editor-api-reference',
    'https://open.wps.cn/previous/docs/client/spreadsheet-api-reference',
    'https://open.wps.cn/previous/docs/client/writer-api-reference',
    'https://open.wps.cn/previous/docs/client/presentation-api-reference',
    'https://open.wps.cn/previous/docs/client/common-api-reference',
    'https://open.wps.cn/previous/docs/client/wps-extension-api',
    'https://open.wps.cn/previous/docs/client/idmso-list',
    // 添加更多特定API对象的页面
    'https://open.wps.cn/previous/docs/client/spreadsheet-api-reference/AboveAverage',
    'https://open.wps.cn/previous/docs/client/spreadsheet-api-reference/AddIn',
    'https://open.wps.cn/previous/docs/client/spreadsheet-api-reference/AddIns',
    'https://open.wps.cn/previous/docs/client/spreadsheet-api-reference/AddIns2',
    'https://open.wps.cn/previous/docs/client/spreadsheet-api-reference/Adjustments',
    'https://open.wps.cn/previous/docs/client/spreadsheet-api-reference/AllowEditRange',
    'https://open.wps.cn/previous/docs/client/spreadsheet-api-reference/AllowEditRanges',
    'https://open.wps.cn/previous/docs/client/spreadsheet-api-reference/Application',
    'https://open.wps.cn/previous/docs/client/spreadsheet-api-reference/Areas',
    'https://open.wps.cn/previous/docs/client/spreadsheet-api-reference/AutoFilter',
    'https://open.wps.cn/previous/docs/client/spreadsheet-api-reference/AutoRecover',
    'https://open.wps.cn/previous/docs/client/spreadsheet-api-reference/Axes',
    'https://open.wps.cn/previous/docs/client/spreadsheet-api-reference/Axis',
    'https://open.wps.cn/previous/docs/client/spreadsheet-api-reference/AxisTitle',
    'https://open.wps.cn/previous/docs/client/spreadsheet-api-reference/Border',
    'https://open.wps.cn/previous/docs/client/spreadsheet-api-reference/Borders',
    'https://open.wps.cn/previous/docs/client/spreadsheet-api-reference/CalculatedFields',
    'https://open.wps.cn/previous/docs/client/spreadsheet-api-reference/CalculatedItems',
    'https://open.wps.cn/previous/docs/client/spreadsheet-api-reference/CategoryCollection',
    'https://open.wps.cn/previous/docs/client/spreadsheet-api-reference/CellFormat',
    'https://open.wps.cn/previous/docs/client/spreadsheet-api-reference/Characters',
    'https://open.wps.cn/previous/docs/client/spreadsheet-api-reference/Chart',
    'https://open.wps.cn/previous/docs/client/spreadsheet-api-reference/ChartArea',
    'https://open.wps.cn/previous/docs/client/spreadsheet-api-reference/ChartCategory',
    'https://open.wps.cn/previous/docs/client/spreadsheet-api-reference/ChartCategory/object',
    'https://open.wps.cn/previous/docs/client/spreadsheet-api-reference/ChartCategory/object-members',
    'https://open.wps.cn/previous/docs/client/spreadsheet-api-reference/ChartCategory/object-properties',
    'https://open.wps.cn/previous/docs/client/spreadsheet-api-reference/ChartCategory/object-methods',
    'https://open.wps.cn/previous/docs/client/spreadsheet-api-reference/ChartFormat'
  ];

  // 抓取这些特定页面
  for (let i = 0; i < specificPages.length; i++) {
    const url = specificPages[i];
    const fileName = `specific_page_${i}_${url.split('/').pop() || 'index'}.html`;
    await scrapePage(browser, url, fileName);
  }
}

async function scrapeAllDocumentation() {
  console.log('Starting WPS API documentation scraper with Puppeteer...');
  
  // 启动浏览器
  const browser = await puppeteer.launch({
    headless: true,
    args: ['--no-sandbox', '--disable-setuid-sandbox']
  });
  
  try {
    // 先抓取特定页面
    await scrapeSpecificPages(browser);
    
    // 然后尝试从主页提取其他链接
    const mainUrl = 'https://open.wps.cn/previous/docs/client/wpsLoad';
    
    // 打开主页面以提取导航链接
    const page = await browser.newPage();
    await page.goto(mainUrl, { waitUntil: 'networkidle2' });
    
    // 等待页面完全加载
    await new Promise(resolve => setTimeout(resolve, 5000));
    
    // 尝试模拟用户交互以展开更多内容
    try {
      // 点击可能的展开按钮
      await page.evaluate(() => {
        const expandButtons = document.querySelectorAll('button, .expand, .toggle, .collapse');
        expandButtons.forEach(btn => {
          if (btn.textContent.includes('展开') || btn.textContent.includes('更多')) {
            btn.click();
          }
        });
        
        // 点击菜单项
        const menuItems = document.querySelectorAll('.menu-item, .nav-item, li');
        menuItems.forEach(item => {
          if (item.textContent.includes('API') || item.textContent.includes('文档')) {
            // 尝试触发点击事件
            const event = new MouseEvent('click', {
              view: window,
              bubbles: true,
              cancelable: true
            });
            item.dispatchEvent(event);
          }
        });
        
        // 特别查找和点击API参考相关的菜单
        const apiMenuItems = document.querySelectorAll('a[href*="api-reference"], a[href*="API"]');
        apiMenuItems.forEach(item => {
          const event = new MouseEvent('click', {
            view: window,
            bubbles: true,
            cancelable: true
          });
          item.dispatchEvent(event);
        });
      });
      
      // 再等待一会儿让内容加载
      await new Promise(resolve => setTimeout(resolve, 3000));
    } catch (interactionError) {
      console.log('Interaction with page elements failed, continuing with link extraction...');
    }
    
    // 提取所有链接
    const links = await extractLinksFromPage(page);
    console.log(`Found ${links.length} links on the page`);
    
    // 过滤掉重复链接并保持顺序
    const uniqueLinks = [];
    const seenUrls = new Set();
    
    for (const link of links) {
      if (!seenUrls.has(link.href)) {
        seenUrls.add(link.href);
        uniqueLinks.push(link);
      }
    }
    
    console.log(`Found ${uniqueLinks.length} unique links`);
    
    // 抓取每个唯一链接的内容（限制数量）
    for (let i = 0; i < Math.min(uniqueLinks.length, 100); i++) {
      const link = uniqueLinks[i];
      try {
        // 生成友好的文件名
        let fileName = `page_${i}_${link.text.replace(/[^a-zA-Z0-9\u4e00-\u9fa5]/g, '_') || 'unnamed'}.html`;
        fileName = fileName.substring(0, 100) + '.html'; // 限制文件名长度
        
        await scrapePage(browser, link.href, fileName);
      } catch (error) {
        console.error(`Error scraping link ${link.href}:`, error.message);
      }
    }
    
    await page.close();
    
    console.log('Documentation scraping completed!');
  } catch (error) {
    console.error('Error during scraping:', error.message);
  } finally {
    await browser.close();
  }
}

scrapeAllDocumentation();