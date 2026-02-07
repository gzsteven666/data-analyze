#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
截图生成器
使用Playwright生成高质量截图
"""

import asyncio
import os
from playwright.async_api import async_playwright
from datetime import datetime

class ScreenshotGenerator:
    """截图生成器类"""

    def __init__(self, output_dir='生成结果信息图'):
        self.output_dir = output_dir
        self.create_output_dir()

    def create_output_dir(self):
        """创建输出目录"""
        os.makedirs(self.output_dir, exist_ok=True)

    async def generate_screenshot(self, html_path, output_name=None,
                                width=1920, height=1080, full_page=True):
        """
        生成截图

        Args:
            html_path: HTML文件路径
            output_name: 输出文件名
            width: 截图宽度
            height: 截图高度
            full_page: 是否截取整个页面

        Returns:
            str: 截图文件路径
        """
        if not output_name:
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            output_name = f"infographic_{timestamp}"

        try:
            async with async_playwright() as p:
                try:
                    # 启动浏览器
                    browser = await p.chromium.launch(
                        headless=True,
                        args=['--no-sandbox', '--disable-setuid-sandbox']
                    )
                except Exception as e:
                    print(f"截图生成失败（浏览器启动失败）: {e}")
                    return None

                # 创建新页面
                page = await browser.new_page(
                    viewport={'width': width, 'height': height},
                    device_scale_factor=2  # 高DPI以获得更清晰的截图
                )

                try:
                    # 加载HTML文件
                    file_url = f"file://{os.path.abspath(html_path)}"
                    await page.goto(file_url, wait_until='networkidle', timeout=30000)

                    # 等待页面完全加载
                    await page.wait_for_load_state('domcontentloaded')
                    await page.wait_for_timeout(2000)  # 额外等待动画完成

                    # 生成截图路径
                    screenshot_path = os.path.join(self.output_dir, f"{output_name}.png")

                    # 截取页面
                    await page.screenshot(
                        path=screenshot_path,
                        full_page=full_page,
                        type='png',
                        quality=100
                    )

                    print(f"截图已生成: {screenshot_path}")
                    return screenshot_path

                except Exception as e:
                    print(f"截图生成失败: {e}")
                    return None

                finally:
                    await browser.close()
        except Exception as e:
            print(f"截图生成失败（Playwright初始化失败）: {e}")
            return None

    def generate_screenshot_sync(self, html_path, output_name=None,
                               width=1920, height=1080, full_page=True):
        """
        同步生成截图

        Args:
            html_path: HTML文件路径
            output_name: 输出文件名
            width: 截图宽度
            height: 截图高度
            full_page: 是否截取整个页面

        Returns:
            str: 截图文件路径
        """
        return asyncio.run(self.generate_screenshot(
            html_path, output_name, width, height, full_page))

    async def generate_multiple_sizes(self, html_path, base_name):
        """
        生成多种尺寸的截图

        Args:
            html_path: HTML文件路径
            base_name: 基础文件名

        Returns:
            dict: 不同尺寸的截图路径
        """
        sizes = {
            'desktop': {'width': 1920, 'height': 1080},
            'tablet': {'width': 1024, 'height': 768},
            'mobile': {'width': 375, 'height': 667}
        }

        paths = {}
        async with async_playwright() as p:
            browser = await p.chromium.launch(headless=True)

            try:
                for size_name, dimensions in sizes.items():
                    page = await browser.new_page(
                        viewport={'width': dimensions['width'], 'height': dimensions['height']},
                        device_scale_factor=2
                    )

                    file_url = f"file://{os.path.abspath(html_path)}"
                    await page.goto(file_url, wait_until='networkidle')
                    await page.wait_for_timeout(2000)

                    screenshot_path = os.path.join(
                        self.output_dir, f"{base_name}_{size_name}.png"
                    )

                    await page.screenshot(
                        path=screenshot_path,
                        full_page=True,
                        type='png',
                        quality=100
                    )

                    paths[size_name] = screenshot_path
                    await page.close()

            finally:
                await browser.close()

        return paths

    def create_comparison_screenshot(self, html_paths, output_name="comparison"):
        """
        创建对比截图（多个HTML文件并排显示）

        Args:
            html_paths: HTML文件路径列表
            output_name: 输出文件名

        Returns:
            str: 对比截图路径
        """
        # 创建对比HTML
        comparison_html = self.create_comparison_html(html_paths)

        # 保存临时HTML文件
        temp_html_path = os.path.join(self.output_dir, f"temp_comparison.html")
        with open(temp_html_path, 'w', encoding='utf-8') as f:
            f.write(comparison_html)

        # 生成截图
        screenshot_path = self.generate_screenshot_sync(
            temp_html_path, output_name, width=2560, height=1440
        )

        # 清理临时文件
        try:
            os.remove(temp_html_path)
        except:
            pass

        return screenshot_path

    def create_comparison_html(self, html_paths):
        """创建对比HTML"""
        template = """
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>数据对比分析</title>
    <style>
        body {
            font-family: 'PingFang SC', 'Microsoft YaHei', sans-serif;
            margin: 0;
            padding: 20px;
            background: #f5f5f5;
        }
        .comparison-container {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(600px, 1fr));
            gap: 20px;
            max-width: 100%;
        }
        .comparison-item {
            background: white;
            border-radius: 10px;
            box-shadow: 0 4px 8px rgba(0,0,0,0.1);
            overflow: hidden;
        }
        .comparison-item iframe {
            width: 100%;
            height: 800px;
            border: none;
        }
        .comparison-title {
            background: #2C3E50;
            color: white;
            padding: 15px;
            text-align: center;
            font-weight: bold;
        }
        .header {
            text-align: center;
            margin-bottom: 30px;
            padding: 20px;
            background: white;
            border-radius: 10px;
            box-shadow: 0 4px 8px rgba(0,0,0,0.1);
        }
        .header h1 {
            color: #2C3E50;
            margin: 0;
        }
    </style>
</head>
<body>
    <div class="header">
        <h1>数据分析对比报告</h1>
        <p>多维度数据可视化对比</p>
    </div>

    <div class="comparison-container">
        {% for i, path in enumerate(html_paths) %}
        <div class="comparison-item">
            <div class="comparison-title">分析视图 {{ i + 1 }}</div>
            <iframe src="file://{{ path }}"></iframe>
        </div>
        {% endfor %}
    </div>
</body>
</html>
        """

        from jinja2 import Template
        tmpl = Template(template)
        return tmpl.render(html_paths=html_paths)

    def optimize_screenshot(self, image_path, quality=95, max_width=None):
        """
        优化截图

        Args:
            image_path: 图片路径
            quality: 质量参数
            max_width: 最大宽度

        Returns:
            str: 优化后的图片路径
        """
        try:
            from PIL import Image

            with Image.open(image_path) as img:
                # 如果需要调整尺寸
                if max_width and img.width > max_width:
                    ratio = max_width / img.width
                    new_height = int(img.height * ratio)
                    img = img.resize((max_width, new_height), Image.Resampling.LANCZOS)

                # 保存优化后的图片
                optimized_path = image_path.replace('.png', '_optimized.png')
                img.save(optimized_path, 'PNG', quality=quality, optimize=True)

                return optimized_path

        except ImportError:
            print("PIL库未安装，跳过图片优化")
            return image_path
        except Exception as e:
            print(f"图片优化失败: {e}")
            return image_path
