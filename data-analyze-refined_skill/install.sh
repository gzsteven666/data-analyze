#!/bin/bash

# 数据分析信息图Skill安装脚本

echo "🚀 开始安装数据分析信息图Skill..."

# 检查Python版本
echo "📋 检查Python环境..."
python3 --version || { echo "❌ Python 3未安装，请先安装Python 3.7+"; exit 1; }

# 检查Node.js版本
echo "📋 检查Node.js环境..."
node --version || { echo "❌ Node.js未安装，请先安装Node.js 14+"; exit 1; }

# 安装Python依赖
echo "📦 安装Python依赖..."
pip3 install -r requirements.txt || { echo "❌ Python依赖安装失败"; exit 1; }

# 安装Node.js依赖
echo "📦 安装Node.js依赖..."
npm install || { echo "❌ Node.js依赖安装失败"; exit 1; }

# 安装Playwright浏览器
echo "🌐 安装Playwright浏览器..."
npx playwright install chromium || { echo "❌ Playwright安装失败"; exit 1; }

# 创建必要的目录
echo "📁 创建输出目录..."
mkdir -p data outputs/csv outputs/excel outputs/figures/{png,svg} outputs/reports "生成结果信息图"

# 设置权限
echo "🔧 设置文件权限..."
chmod +x src/main.py
chmod +x install.sh

# 运行测试
echo "🧪 运行功能测试..."
node test/test_skill.js

echo ""
echo "✅ 安装完成！"
echo ""
echo "📖 使用说明:"
echo "  1. 将数据文件放入 data/ 目录"
echo "  2. 运行分析: node src/index.js data/你的数据.xlsx"
echo "  3. 查看结果: 生成结果信息图/ 目录"
echo ""
echo "🎯 示例命令:"
echo "  node src/index.js data/quarter_sales.xlsx --sheet=\"Q4\" --company=\"旗舰店A\" --time=\"月份\" --value=\"GMV\""
echo "  node src/index.js data/市场数据.csv --company=\"测试公司\""
echo ""
echo "📚 更多信息请查看 README.md"
