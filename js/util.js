//在后续的wps版本中，wps的所有枚举值都会通过wps.Enum对象来自动支持，现阶段先人工定义
var WPS_Enum = {
    msoCTPDockPositionLeft:0,
    msoCTPDockPositionRight:2
}

//从配置文件读取服务器配置 - 统一读取index.html配置
function loadServerConfig() {
    try {
        // 优先级1: 尝试从嵌入的配置中读取
        var configScript = document.getElementById('server-config');
        if (configScript && configScript.textContent) {
            var configText = configScript.textContent.trim();
            if (configText) {
                var config = {};
                
                // 解析嵌入的配置文件
                var lines = configText.split('\n');
                for (var j = 0; j < lines.length; j++) {
                    var line = lines[j].trim();
                    if (line && line.indexOf('=') > 0) {
                        var parts = line.split('=');
                        config[parts[0].trim()] = parts[1].trim();
                    }
                }
                
                // 验证主服务器配置
                if (config.PRODUCTION) {
                    return {
                        PRODUCTION: config.PRODUCTION
                    };
                }
            }
        }
        
        // 优先级2: 尝试从全局变量中读取（如果当前页面是index.html）
        if (typeof PRODUCTION !== 'undefined') {
            return {
                PRODUCTION: PRODUCTION || 'http://192.168.70.17:8888/V6/'
            };
        }
        
        // 优先级3: 尝试通过window.parent访问父页面配置
        try {
            if (window.parent && window.parent !== window && typeof window.parent.PRODUCTION !== 'undefined') {
                return {
                    PRODUCTION: window.parent.PRODUCTION || 'http://192.168.70.17:8888/V6/'
                };
            }
        } catch (e) {
            // 跨域访问被拒绝，继续下一级
        }
        
    } catch (e) {
        // 配置读取失败，记录错误
        if (typeof console !== 'undefined' && console.warn) {
            console.warn('⚠️ 配置加载失败:', e);
        }
    }
    
    // 最后的回退选项: 使用默认配置
        return {
            PRODUCTION: 'http://192.168.70.17:8888/V6/'
        };
}

// 统一配置管理
(function() {
    try {
        // 加载配置
        var CONFIG = loadServerConfig();
        
        // 服务器配置对象 - 简化配置
        var SERVER_CONFIG = {
            PRODUCTION: CONFIG.PRODUCTION
        };
        
        // 将配置暴露到全局作用域
        window.CONFIG = CONFIG;
        window.SERVER_CONFIG = SERVER_CONFIG;
        
        // 配置结果已简化
        
    } catch (e) {
        console.warn('配置初始化失败:', e);
    }
})();

/**
 * 获取主服务器地址
 * @returns {Promise<string>} 主服务器地址
 */
function selectServer() {
    return new Promise((resolve, reject) => {
        var selectedServer = window.CONFIG.PRODUCTION;
        resolve(selectedServer);
    });
}

function GetUrlPath() {
    let e = document.location.toString()
    return -1!=(e=decodeURI(e)).indexOf("/")&&(e=e.substring(0,e.lastIndexOf("/"))),e
}

/**
 * 通过wps提供的接口执行一段脚本
 * @param {*} param 需要执行的脚本
 */
function shellExecuteByOAAssist(param) {
    if (wps != null) {
        wps.OAAssist.ShellExecute(param)
    }
}