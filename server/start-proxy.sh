#!/bin/bash
# AI接口代理服务器启动脚本

echo "🚀 启动AI接口代理服务器..."

# 检查Python环境
if ! command -v python3 &> /dev/null; then
    echo "❌ 错误: 未找到python3，请先安装Python 3.6+"
    exit 1
fi

# 检查Flask和requests是否安装
python3 -c "import flask, requests" 2>/dev/null
if [ $? -ne 0 ]; then
    echo "📦 安装必需的Python包..."
    pip3 install flask flask-cors requests
fi

# 设置环境变量
if [ -f "config.env" ]; then
    echo "📋 加载配置文件..."
    export $(cat config.env | grep -v '#' | xargs)
fi

# 检查API密钥
if [ -z "$DASHSCOPE_API_KEY" ]; then
    echo "⚠️  警告: 未设置DASHSCOPE_API_KEY环境变量"
    echo "请设置您的DashScope API密钥:"
    echo "export DASHSCOPE_API_KEY='your-api-key'"
    echo ""
    echo "或者编辑config.env文件并设置您的API密钥"
    echo ""
    read -p "是否继续启动? (y/N): " -n 1 -r
    echo
    if [[ ! $REPLY =~ ^[Yy]$ ]]; then
        exit 1
    fi
fi

echo "🔧 服务器配置:"
echo "   端点: ${DASHSCOPE_ENDPOINT:-https://dashscope.aliyuncs.com}"
echo "   模型: ${DEFAULT_MODEL:-qwen-plus}"
echo "   超时: ${TIMEOUT:-30}秒"
echo "   重试: ${MAX_RETRIES:-3}次"
echo ""

echo "🌐 启动HTTP服务器在端口8080..."
echo "   健康检查: http://localhost:8080/api/health"
echo "   测试连接: http://localhost:8080/api/test"
echo "   Chat接口: http://localhost:8080/api/chat/completions"
echo ""
echo "按 Ctrl+C 停止服务器"
echo "================================================"

# 启动服务器
python3 proxy-server.py