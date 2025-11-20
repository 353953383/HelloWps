#!/usr/bin/env node
/**
 * AI接口代理服务器 (Node.js版本)
 * 解决CORS问题，将前端请求代理到DashScope API
 */

const express = require('express');
const cors = require('cors');
const axios = require('axios');
const path = require('path');

const app = express();
const PORT = process.env.PORT || 3889;

// 配置CORS
app.use(cors());
app.use(express.json());

// API配置
const API_KEY = process.env.DASHSCOPE_API_KEY || 'sk-9bacbdffb7dd4b91b240c472d9c5e0c2';
const API_ENDPOINT = process.env.DASHSCOPE_ENDPOINT || 'https://dashscope.aliyuncs.com';
const DEFAULT_MODEL = process.env.DEFAULT_MODEL || 'qwen-plus';
const TIMEOUT = parseInt(process.env.TIMEOUT) || 60000;

// 统计信息
const STATS = {
    totalRequests: 0,
    successfulRequests: 0,
    failedRequests: 0,
    startTime: new Date()
};

/**
 * 验证API密钥
 */
function validateApiKey() {
    if (!API_KEY || API_KEY.length < 10) {
        throw new Error('API密钥未配置或格式不正确');
    }
    return true;
}

/**
 * 转换请求格式以适配DashScope API
 */
function convertRequestFormat(requestData) {
    try {
        // 检查是否是OpenAI兼容格式
        if (requestData.messages) {
            // 提取系统消息和用户消息
            const systemMessages = [];
            const userMessages = [];
            
            requestData.messages.forEach(msg => {
                if (msg.role === 'system') {
                    systemMessages.push(msg.content || '');
                } else if (msg.role === 'user') {
                    userMessages.push(msg.content || '');
                }
            });
            
            // 构建DashScope格式的输入
            const inputText = userMessages.join('\n') || '';
            const systemPrompt = systemMessages.join('\n') || '';
            
            const dashscopePayload = {
                model: requestData.model || DEFAULT_MODEL,
                input: {
                    prompt: inputText,
                    history: []
                },
                parameters: {
                    max_tokens: requestData.max_tokens || 4000,
                    temperature: requestData.temperature || 0.7,
                    top_p: requestData.top_p || 0.8
                }
            };
            
            if (systemPrompt) {
                dashscopePayload.input.system = systemPrompt;
            }
            
            return dashscopePayload;
        }
        
        // 如果已经是DashScope格式，直接返回
        return requestData;
        
    } catch (error) {
        console.error('请求格式转换失败:', error);
        throw new Error(`请求格式转换失败: ${error.message}`);
    }
}

/**
 * 转换响应格式为OpenAI兼容格式
 */
function convertResponseFormat(dashscopeResponse) {
    try {
        if (dashscopeResponse.output && dashscopeResponse.output.text) {
            return {
                choices: [
                    {
                        message: {
                            role: 'assistant',
                            content: dashscopeResponse.output.text
                        },
                        finish_reason: 'stop'
                    }
                ],
                usage: dashscopeResponse.usage || {},
                id: dashscopeResponse.request_id || '',
                model: dashscopeResponse.model || DEFAULT_MODEL,
                created: Math.floor(Date.now() / 1000)
            };
        }
        
        // 如果转换失败，返回原始响应
        return dashscopeResponse;
        
    } catch (error) {
        console.error('响应格式转换失败:', error);
        return {
            error: {
                message: `响应格式转换失败: ${error.message}`,
                type: 'conversion_error'
            }
        };
    }
}

/**
 * 发送请求到DashScope API
 */
async function makeRequest(payload) {
    const url = `${API_ENDPOINT}/api/v1/services/aigc/text-generation/generation`;
    
    const headers = {
        'Authorization': `Bearer ${API_KEY}`,
        'Content-Type': 'application/json',
        'User-Agent': 'AI-Proxy-Server/1.0'
    };
    
    const maxRetries = 3;
    
    for (let attempt = 1; attempt <= maxRetries; attempt++) {
        try {
            console.log(`尝试请求 ${url} (第${attempt}次)`);
            
            const response = await axios.post(url, payload, {
                headers: headers,
                timeout: TIMEOUT
            });
            
            if (response.status === 200) {
                return response.data;
            } else {
                console.warn(`请求失败，状态码: ${response.status}, 响应: ${response.data}`);
                
                if (attempt === maxRetries) {
                    throw new Error(`API请求失败: ${response.status}`);
                }
            }
            
        } catch (error) {
            console.error(`请求异常 (第${attempt}次):`, error.message);
            
            if (attempt === maxRetries) {
                throw error;
            }
            
            // 等待重试 (指数退避)
            await new Promise(resolve => setTimeout(resolve, Math.pow(2, attempt) * 1000));
        }
    }
    
    throw new Error('所有重试均失败');
}

// 路由

/**
 * 健康检查
 */
app.get('/api/health', (req, res) => {
    res.json({
        status: 'healthy',
        timestamp: new Date().toISOString(),
        uptime: `${Math.floor((Date.now() - STATS.startTime) / 1000)}秒`,
        stats: STATS
    });
});

/**
 * 测试API连接
 */
app.post('/api/test', async (req, res) => {
    try {
        validateApiKey();
        
        const testPayload = {
            model: DEFAULT_MODEL,
            input: {
                prompt: '测试连接，请回复"连接成功"'
            },
            parameters: {
                max_tokens: 100,
                temperature: 0.1
            }
        };
        
        const response = await makeRequest(testPayload);
        
        res.json({
            success: true,
            message: 'API连接测试成功',
            response: response
        });
        
    } catch (error) {
        console.error('API连接测试失败:', error);
        res.status(400).json({
            success: false,
            error: error.message,
            message: 'API连接测试失败'
        });
    }
});

/**
 * Chat Completions代理接口
 */
app.post('/api/chat/completions', async (req, res) => {
    STATS.totalRequests++;
    
    try {
        console.log('收到Chat Completions请求');
        
        // 验证API密钥
        validateApiKey();
        
        // 获取请求数据
        const requestData = req.body;
        if (!requestData) {
            return res.status(400).json({
                error: {
                    message: '请求体不能为空',
                    type: 'invalid_request'
                }
            });
        }
        
        console.log(`请求数据大小: ${JSON.stringify(requestData).length} 字符`);
        
        // 转换请求格式
        const dashscopePayload = convertRequestFormat(requestData);
        console.log('转换后的请求格式:', JSON.stringify(dashscopePayload, null, 2));
        
        // 发送请求
        const dashscopeResponse = await makeRequest(dashscopePayload);
        console.log('DashScope原始响应:', JSON.stringify(dashscopeResponse, null, 2));
        
        // 转换响应格式
        const openaiResponse = convertResponseFormat(dashscopeResponse);
        
        STATS.successfulRequests++;
        console.log('✅ AI请求成功处理');
        
        res.json(openaiResponse);
        
    } catch (error) {
        STATS.failedRequests++;
        console.error('❌ 请求处理失败:', error.message);
        
        res.status(error.response?.status || 500).json({
            error: {
                message: error.message || '内部服务器错误',
                type: error.code || 'internal_error',
                details: error.response?.data
            }
        });
    }
});

/**
 * 获取统计信息
 */
app.get('/api/stats', (req, res) => {
    res.json({
        stats: STATS,
        config: {
            endpoint: API_ENDPOINT,
            model: DEFAULT_MODEL,
            timeout: TIMEOUT,
            hasApiKey: !!API_KEY
        }
    });
});

/**
 * 错误处理中间件
 */
app.use((error, req, res, next) => {
    console.error('服务器错误:', error);
    res.status(500).json({
        error: {
            message: '内部服务器错误',
            type: 'internal_error'
        }
    });
});

/**
 * 404处理
 */
app.use('*', (req, res) => {
    res.status(404).json({
        error: {
            message: '接口不存在',
            type: 'not_found'
        }
    });
});

// 启动服务器
app.listen(PORT, '0.0.0.0', () => {
    console.log('🚀 AI接口代理服务器启动成功!');
    console.log(`📡 端口: ${PORT}`);
    console.log(`🔗 端点: ${API_ENDPOINT}`);
    console.log(`🤖 模型: ${DEFAULT_MODEL}`);
    console.log(`🔑 API密钥配置: ${API_KEY ? '是' : '否'}`);
    console.log('');
    console.log('🌐 可用接口:');
    console.log(`   健康检查: http://localhost:${PORT}/api/health`);
    console.log(`   连接测试: http://localhost:${PORT}/api/test`);
    console.log(`   Chat接口: http://localhost:${PORT}/api/chat/completions`);
    console.log(`   统计信息: http://localhost:${PORT}/api/stats`);
    console.log('');
    console.log('按 Ctrl+C 停止服务器');
});

module.exports = app;