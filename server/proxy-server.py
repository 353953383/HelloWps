#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AI接口代理服务器
解决CORS问题，将前端请求代理到DashScope API
"""

import os
import json
import time
import logging
import requests
from urllib.parse import urljoin
from flask import Flask, request, jsonify
from flask_cors import CORS
from datetime import datetime

# 配置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

app = Flask(__name__)
CORS(app)  # 解决CORS问题

# 读取环境变量并设置默认值
API_KEY = os.getenv('DASHSCOPE_API_KEY', 'sk-9bacbdffb7dd4b91b240c472d9c5e0c2')  # 默认API密钥
API_ENDPOINT = os.getenv('DASHSCOPE_ENDPOINT', 'https://dashscope.aliyuncs.com')
DEFAULT_MODEL = os.getenv('DEFAULT_MODEL', 'qwen-plus')
TIMEOUT = int(os.getenv('TIMEOUT', '60'))  # 默认60秒超时
MAX_RETRIES = int(os.getenv('MAX_RETRIES', '3'))
PORT = int(os.getenv('PORT', '3889'))  # 默认3889端口匹配前端配置

# 统计信息
STATS = {
    'total_requests': 0,
    'successful_requests': 0,
    'failed_requests': 0,
    'start_time': datetime.now()
}

class APIProxy:
    """API代理类"""
    
    @staticmethod
    def validate_api_key():
        """验证API密钥"""
        if not API_KEY or len(API_KEY) < 10:
            raise ValueError("API密钥未配置或格式不正确")
        return True
    
    @staticmethod
    def convert_request_format(request_data):
        """转换请求格式以适配DashScope API"""
        try:
            # 检查是否是OpenAI兼容格式
            if 'messages' in request_data:
                # 转换为DashScope格式
                messages = request_data.get('messages', [])
                
                # 提取系统消息和用户消息
                system_messages = []
                user_messages = []
                
                for msg in messages:
                    if msg.get('role') == 'system':
                        system_messages.append(msg.get('content', ''))
                    elif msg.get('role') == 'user':
                        user_messages.append(msg.get('content', ''))
                
                # 构建DashScope格式的输入
                input_text = '\n'.join(user_messages) if user_messages else ''
                system_prompt = '\n'.join(system_messages) if system_messages else ''
                
                dashscope_payload = {
                    "model": request_data.get('model', DEFAULT_MODEL),
                    "input": {
                        "prompt": input_text,
                        "history": []
                    },
                    "parameters": {
                        "max_tokens": request_data.get('max_tokens', 4000),
                        "temperature": request_data.get('temperature', 0.7),
                        "top_p": request_data.get('top_p', 0.8)
                    }
                }
                
                if system_prompt:
                    dashscope_payload["input"]["system"] = system_prompt
                
                return dashscope_payload
            
            # 如果已经是DashScope格式，直接返回
            return request_data
            
        except Exception as e:
            logger.error(f"请求格式转换失败: {e}")
            raise ValueError(f"请求格式转换失败: {e}")
    
    @staticmethod
    def convert_response_format(dashscope_response):
        """转换响应格式为OpenAI兼容格式"""
        try:
            if 'output' in dashscope_response:
                output = dashscope_response['output']
                if 'text' in output:
                    return {
                        "choices": [
                            {
                                "message": {
                                    "role": "assistant",
                                    "content": output['text']
                                },
                                "finish_reason": "stop"
                            }
                        ],
                        "usage": dashscope_response.get('usage', {}),
                        "id": dashscope_response.get('request_id', ''),
                        "model": dashscope_response.get('model', DEFAULT_MODEL),
                        "created": int(time.time())
                    }
            
            # 如果转换失败，返回原始响应
            return dashscope_response
            
        except Exception as e:
            logger.error(f"响应格式转换失败: {e}")
            return {
                "error": {
                    "message": f"响应格式转换失败: {e}",
                    "type": "conversion_error"
                }
            }
    
    @staticmethod
    def make_request(payload, endpoint_suffix="/api/v1/services/aigc/text-generation/generation"):
        """发送请求到DashScope API"""
        url = urljoin(API_ENDPOINT, endpoint_suffix)
        
        headers = {
            'Authorization': f'Bearer {API_KEY}',
            'Content-Type': 'application/json',
            'User-Agent': 'AI-Proxy-Server/1.0'
        }
        
        for attempt in range(MAX_RETRIES):
            try:
                logger.info(f"尝试请求 {url} (第{attempt + 1}次)")
                
                response = requests.post(
                    url,
                    headers=headers,
                    json=payload,
                    timeout=TIMEOUT
                )
                
                if response.status_code == 200:
                    return response.json()
                else:
                    logger.warning(f"请求失败，状态码: {response.status_code}, 响应: {response.text}")
                    
                    if attempt == MAX_RETRIES - 1:
                        raise requests.RequestException(f"API请求失败: {response.status_code}")
                    
            except requests.exceptions.Timeout:
                logger.warning(f"请求超时 (第{attempt + 1}次)")
                if attempt == MAX_RETRIES - 1:
                    raise
                    
            except requests.exceptions.RequestException as e:
                logger.error(f"请求异常 (第{attempt + 1}次): {e}")
                if attempt == MAX_RETRIES - 1:
                    raise
                
                # 等待重试
                time.sleep(2 ** attempt)
        
        raise Exception("所有重试均失败")

@app.route('/api/health', methods=['GET'])
def health_check():
    """健康检查接口"""
    return jsonify({
        'status': 'healthy',
        'timestamp': datetime.now().isoformat(),
        'uptime': str(datetime.now() - STATS['start_time']),
        'stats': STATS
    })

@app.route('/api/test', methods=['POST'])
def test_connection():
    """测试API连接"""
    try:
        APIProxy.validate_api_key()
        
        # 构建简单的测试请求
        test_payload = {
            "model": DEFAULT_MODEL,
            "input": {
                "prompt": "测试连接，请回复'连接成功'"
            },
            "parameters": {
                "max_tokens": 100,
                "temperature": 0.1
            }
        }
        
        response = APIProxy.make_request(test_payload)
        
        return jsonify({
            'success': True,
            'message': 'API连接测试成功',
            'response': response
        })
        
    except Exception as e:
        logger.error(f"API连接测试失败: {e}")
        return jsonify({
            'success': False,
            'error': str(e),
            'message': 'API连接测试失败'
        }), 400

@app.route('/api/chat/completions', methods=['POST'])
def chat_completions():
    """Chat Completions代理接口"""
    STATS['total_requests'] += 1
    
    try:
        # 验证API密钥
        APIProxy.validate_api_key()
        
        # 获取请求数据
        request_data = request.get_json()
        if not request_data:
            return jsonify({
                'error': {
                    'message': '请求体不能为空',
                    'type': 'invalid_request'
                }
            }), 400
        
        logger.info(f"收到Chat Completions请求，数据大小: {len(json.dumps(request_data))}")
        
        # 转换请求格式
        dashscope_payload = APIProxy.convert_request_format(request_data)
        
        # 发送请求
        dashscope_response = APIProxy.make_request(dashscope_payload)
        
        # 转换响应格式
        openai_response = APIProxy.convert_response_format(dashscope_response)
        
        STATS['successful_requests'] += 1
        
        return jsonify(openai_response)
        
    except ValueError as e:
        STATS['failed_requests'] += 1
        logger.error(f"请求验证失败: {e}")
        return jsonify({
            'error': {
                'message': str(e),
                'type': 'validation_error'
            }
        }), 400
        
    except Exception as e:
        STATS['failed_requests'] += 1
        logger.error(f"请求处理失败: {e}")
        return jsonify({
            'error': {
                'message': '内部服务器错误',
                'type': 'internal_error',
                'details': str(e)
            }
        }), 500

@app.route('/api/stats', methods=['GET'])
def get_stats():
    """获取统计信息"""
    return jsonify({
        'stats': STATS,
        'config': {
            'endpoint': API_ENDPOINT,
            'model': DEFAULT_MODEL,
            'timeout': TIMEOUT,
            'max_retries': MAX_RETRIES
        }
    })

@app.errorhandler(404)
def not_found(error):
    return jsonify({
        'error': {
            'message': '接口不存在',
            'type': 'not_found'
        }
    }), 404

@app.errorhandler(500)
def internal_error(error):
    return jsonify({
        'error': {
            'message': '内部服务器错误',
            'type': 'internal_error'
        }
    }), 500

if __name__ == '__main__':
    logger.info("启动AI接口代理服务器...")
    logger.info(f"DashScope端点: {API_ENDPOINT}")
    logger.info(f"默认模型: {DEFAULT_MODEL}")
    logger.info(f"API密钥配置: {'是' if API_KEY else '否'}")
    logger.info(f"服务器端口: {PORT}")
    
    # 检查必需的环境变量
    if not API_KEY:
        logger.warning("警告: 未设置DASHSCOPE_API_KEY环境变量")
        logger.info("请设置环境变量: export DASHSCOPE_API_KEY='your-api-key'")
    
    # 启动服务器
    app.run(
        host='0.0.0.0',
        port=PORT,
        debug=True,
        threaded=True
    )