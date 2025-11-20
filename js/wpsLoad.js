// WPS API 加载模块 - WPS JSA版本
// 专门为WPS表格JSA环境设计，直接使用WPS提供的API

(function() {
    // 检查是否已经定义了wpsLoad
    if (typeof window.wpsLoad === 'undefined') {
        /**
         * @returns {Promise} 返回Promise，解析为WPS API对象
         */
        window.wpsLoad = function() {
            return new Promise(function(resolve, reject) {
                try {
                    // console.log('WPS JSA环境：开始加载WPS API...');
                    
                    // 在WPS JSA环境中，直接使用wps全局对象
                    if (typeof window.wps !== 'undefined') {
                        // console.log('WPS JSA环境：使用全局wps对象');
                        resolve(window.wps);
                        return;
                    }
                    
                    // 如果wps对象不存在，但在WPS环境中，尝试直接创建一个API对象
                    // 这是为了确保在WPS表格环境中的兼容性
                    const wpsApi = {
                        // 使用window对象作为根，确保JSA规范兼容性
                        Application: window.Application || window.external?.Application
                    };
                    
                    if (wpsApi.Application) {
                        // console.log('WPS JSA环境：使用Application对象');
                        resolve(wpsApi);
                        return;
                    }
                    
                    // 开发环境下的模拟模式（仅用于测试）
                    if (window.location.hostname === 'localhost' || 
                        window.location.hostname === '127.0.0.1') {
                        // console.log('开发环境模式：创建Excel模拟API对象');
                        const mockWps = {
                            Application: {
                                Version: 'Mock WPS 11.1.0.13212',
                                Workbooks: {
                                    Count: 1,
                                    Item: function(index) { 
                                        return {
                                            Name: 'Mock Workbook.xlsx',
                                            Sheets: { Count: 3 }
                                        };
                                    },
                                    Add: function() { console.log('模拟添加工作簿'); }
                                },
                                ActiveWorkbook: {
                                    Name: 'Mock Workbook.xlsx',
                                    Worksheets: {
                                        Count: 3,
                                        Item: function(index) {
                                            return {
                                                Name: 'Sheet' + index,
                                                Range: function() {
                                                    return {
                                                        Value: '模拟单元格值'
                                                    };
                                                }
                                            };
                                        }
                                    },
                                    ActiveSheet: {
                                        Name: 'Sheet1',
                                        Range: function() {
                                            return {
                                                Value: '模拟单元格值'
                                            };
                                        }
                                    }
                                },
                                ActiveDocument: function() {
                                    return this.ActiveWorkbook;
                                }
                            }
                        };
                        resolve(mockWps);
                        return;
                    }
                    
                    // 如果在WPS环境中但无法获取API，返回适当的错误
                    reject(new Error('WPS JSA环境：无法获取WPS API对象'));
                } catch (error) {
                    console.error('WPS API加载过程中发生错误:', error);
                    reject(error);
                }
            });
        };
        
        // console.log('wpsLoad.js WPS JSA版本已加载');
    } else {
        // console.log('wpsLoad已经定义，跳过重新加载');
    }
})();