#!/usr/bin/env node

/**
 * 数据分析信息图Skill - Node.js入口
 * 整合Python数据分析模块，提供完整的CLI接口
 */

const { PythonShell } = require('python-shell');
const { execFile } = require('child_process');
const { promisify } = require('util');
const path = require('path');
const fs = require('fs-extra');

const execFileAsync = promisify(execFile);

class DataAnalysisSkill {
    constructor() {
        this.pythonScriptPath = path.join(__dirname, 'main.py');
        this.defaultOptions = {
            pythonPath: process.env.PYTHON || (process.platform === 'win32' ? 'python' : 'python3'),
            pythonOptions: ['-u'], // 无缓冲输出
            scriptPath: __dirname
        };
    }

    /**
     * 运行数据分析
     * @param {string} dataPath - 数据文件路径
     * @param {Object} options - 分析选项
     * @returns {Promise<Object>} 分析结果
     */
    async analyzeData(dataPath, options = {}) {
        const sheetName = options.sheetName || options.sheet || null;
        const targetBrand = options.targetBrand || options.target_brand || null;
        const companyName = options.companyName || options.company || targetBrand || '';
        const outputDir = options.outputDir || options.output || null;
        const timeColumn = options.timeColumn || options.time || null;
        const valueColumn = options.valueColumn || options.value || null;
        const enableCharts = options.enableCharts || options.enable_charts || false; // legacy
        const chartsMode = options.chartsMode || options.charts_mode || 'auto';
        const enableScreenshot = options.enableScreenshot || options.enable_screenshot || false;
        const coreDimension = options.coreDimension || options.core_dimension || null;

        // 检查数据文件
        if (!await fs.pathExists(dataPath)) {
            throw new Error(`数据文件不存在: ${dataPath}`);
        }

        // 准备Python脚本参数
        const args = [dataPath];
        if (sheetName) {
            args.push('--sheet');
            args.push(sheetName);
        }
        if (companyName) {
            args.push('--company');
            args.push(companyName);
        }
        if (timeColumn) {
            args.push('--time-column');
            args.push(timeColumn);
        }
        if (valueColumn) {
            args.push('--value-column');
            args.push(valueColumn);
        }
        if (outputDir) {
            args.push('--output-dir');
            args.push(outputDir);
        }
        if (enableCharts) {
            args.push('--enable-charts');
        }
        if (chartsMode && !enableCharts) {
            args.push('--charts-mode');
            args.push(chartsMode);
        }
        if (enableScreenshot) {
            args.push('--enable-screenshot');
        }
        if (coreDimension) {
            args.push('--core-dimension');
            args.push(coreDimension);
        }
        if (targetBrand) {
            args.push('--target-brand');
            args.push(targetBrand);
        }

        const pythonOptions = {
            ...this.defaultOptions,
            args: args
        };

        console.log(`开始数据分析: ${dataPath}`);
        console.log(`参数: ${JSON.stringify(args)}`);

        try {
            // 运行Python脚本
            const results = await PythonShell.run('main.py', pythonOptions);

            console.log('Python脚本执行完成');

            // 解析输出结果
            const output = this.parsePythonOutput(results);

            return {
                success: true,
                message: '数据分析完成',
                output: output,
                timestamp: new Date().toISOString()
            };

        } catch (error) {
            console.error('Python脚本执行失败:', error);
            throw new Error(`数据分析失败: ${error.message}`);
        }
    }

    /**
     * 解析Python脚本输出
     * @param {Array} outputLines - Python输出
     * @returns {Object} 解析后的结果
     */
    parsePythonOutput(outputLines) {
        const output = {
            charts: [],
            html: null,
            screenshot: null,
            csvFiles: [],
            excelFile: null,
            insights: null,
            wordReport: null,
            outputRoot: null
        };

        // 解析输出行
        const outputText = outputLines.join('\n');

        // 提取文件路径
        const filePatterns = {
            html: /信息图HTML:\s*(.+)/,
            screenshot: /截图文件:\s*(.+)/,
            excel: /Excel汇总:\s*(.+)/,
            csv: /CSV文件:\s*(.+)目录下/,
            charts: /图表文件:\s*(\d+)个/,
            word: /Word报告:\s*(.+)/,
            outputRoot: /输出根目录:\s*(.+)/
        };

        // 提取HTML文件路径
        const htmlMatch = outputText.match(filePatterns.html);
        if (htmlMatch) {
            output.html = htmlMatch[1].trim();
        }

        // 提取截图文件路径
        const screenshotMatch = outputText.match(filePatterns.screenshot);
        if (screenshotMatch) {
            output.screenshot = screenshotMatch[1].trim();
        }

        // 提取Excel文件路径
        const excelMatch = outputText.match(filePatterns.excel);
        if (excelMatch) {
            output.excelFile = excelMatch[1].trim();
        }

        const wordMatch = outputText.match(filePatterns.word);
        if (wordMatch) {
            output.wordReport = wordMatch[1].trim();
        }

        const rootMatch = outputText.match(filePatterns.outputRoot);
        if (rootMatch) {
            output.outputRoot = rootMatch[1].trim();
        }

        // 提取图表数量
        const chartsMatch = outputText.match(filePatterns.charts);
        if (chartsMatch) {
            output.chartCount = parseInt(chartsMatch[1]);
        }

        // 提取CSV目录
        const csvMatch = outputText.match(filePatterns.csv);
        if (csvMatch) {
            output.csvDir = csvMatch[1].trim();
        }

        return output;
    }

    /**
     * 获取支持的文件类型
     * @returns {Array} 支持的文件扩展名
     */
    getSupportedFileTypes() {
        return ['.csv', '.xlsx', '.xls'];
    }

    /**
     * 验证文件类型
     * @param {string} filePath - 文件路径
     * @returns {boolean} 是否支持
     */
    isValidFileType(filePath) {
        const ext = path.extname(filePath).toLowerCase();
        return this.getSupportedFileTypes().includes(ext);
    }

    /**
     * 获取Excel工作表列表
     * @param {string} excelPath - Excel文件路径
     * @returns {Promise<Array>} 工作表名称列表
     */
    async getExcelSheets(excelPath) {
        if (!await fs.pathExists(excelPath)) {
            throw new Error(`Excel文件不存在: ${excelPath}`);
        }

        const ext = path.extname(excelPath).toLowerCase();
        if (!['.xlsx', '.xls'].includes(ext)) {
            throw new Error('不是有效的Excel文件');
        }

        try {
            const code = [
                'import sys',
                'from openpyxl import load_workbook',
                'wb = load_workbook(sys.argv[1], read_only=True)',
                'for name in wb.sheetnames:',
                '    print(name)'
            ].join('\n');
            const { stdout } = await execFileAsync(
                this.defaultOptions.pythonPath,
                ['-c', code, excelPath],
                { cwd: __dirname }
            );
            const sheets = stdout.split(/\r?\n/).map(line => line.trim()).filter(Boolean);
            return sheets;

        } catch (error) {
            console.error('获取工作表失败:', error);
            return [];
        }
    }

    /**
     * 生成分析报告
     * @param {Object} analysisResult - 分析结果
     * @returns {string} 报告内容
     */
    generateReport(analysisResult) {
        const { output, timestamp } = analysisResult;

        let report = `# 数据分析报告\n\n`;
        report += `生成时间: ${new Date(timestamp).toLocaleString()}\n\n`;

        if (output.screenshot) {
            report += `## 信息图截图\n`;
            report += `![分析截图](${output.screenshot})\n\n`;
        }

        if (output.html) {
            report += `## 交互式报告\n`;
            report += `HTML文件: ${output.html}\n\n`;
        }

        if (output.excelFile) {
            report += `## Excel汇总\n`;
            report += `汇总文件: ${output.excelFile}\n\n`;
        }

        if (output.chartCount) {
            report += `## 图表统计\n`;
            report += `生成图表数量: ${output.chartCount}个\n\n`;
        }

        report += `## 文件输出\n`;
        report += `- CSV文件目录: ${output.csvDir || 'outputs/csv'}\n`;
        report += `- Excel汇总: ${output.excelFile || 'outputs/excel'}\n`;
        report += `- 图表文件: ${output.charts || 'outputs/figures'}\n`;
        report += `- HTML报告: ${output.html || 'outputs/html'}\n`;
        report += `- 截图文件: ${output.screenshot || 'outputs/screenshots'}\n`;

        return report;
    }
}

/**
 * CLI接口
 */
async function main() {
    const args = process.argv.slice(2);

    if (args.length === 0) {
        console.log(`
业务数据市场洞察分析包 - 使用说明

用法:
  node index.js <数据文件路径> [选项]

参数:
  数据文件路径          - CSV或Excel文件路径
  --sheet              - Excel工作表名称
  --company            - 可选：高亮或叙述的目标对象名称
  --time-column        - 可选：时间列名称
  --value-column       - 可选：核心指标列名称
  --output-dir         - 输出目录（默认：数据同级 outputs/）
  --charts-mode        - auto/on/off，默认 auto
  --core-dimension     - 可选：核心维度列
  --target-brand       - 可选：目标品牌/对象，用于白区与机会分析
  --enable-screenshot  - 可选：生成 HTML 截图

示例:
  node index.js data/quarter_sales.xlsx --sheet="Q4" --company="旗舰店A" --value-column="GMV" --charts-mode=on
  node index.js data/招采数据.xlsx --sheet="导出数据表" --core-dimension="申报企业名称" --value-column="协议采购量" --target-brand="林华"

支持的文件类型:
  - CSV文件 (.csv)
  - Excel文件 (.xlsx, .xls)
        `);
        return;
    }

    const dataPath = args[0];
    const options = {};

    // 解析命令行参数
    for (let i = 1; i < args.length; i++) {
        const token = args[i];
        if (!token.startsWith('--')) continue;

        let key = token.replace(/^--/, '');
        let value = true;

        if (token.includes('=')) {
            const [k, v] = token.replace(/^--/, '').split('=');
            key = k;
            value = v;
        } else {
            const next = args[i + 1];
            if (next && !next.startsWith('--')) {
                value = next;
                i++;
            }
        }

        options[key] = value;
    }

    try {
        const skill = new DataAnalysisSkill();

        // 验证文件类型
        if (!skill.isValidFileType(dataPath)) {
            console.error('错误: 不支持的文件类型');
            console.log('支持的文件类型:', skill.getSupportedFileTypes().join(', '));
            process.exit(1);
        }

        // 执行数据分析
        console.log('开始数据分析...');
        const result = await skill.analyzeData(dataPath, options);

        if (result.success) {
            console.log('\n✅ 数据分析完成！');

            // 生成并显示报告
            const report = skill.generateReport(result);
            console.log('\n' + report);

            console.log('\n📊 所有输出文件:');
            Object.entries(result.output).forEach(([key, value]) => {
                if (value) {
                    console.log(`  ${key}: ${value}`);
                }
            });

        } else {
            console.error('❌ 数据分析失败');
            process.exit(1);
        }

    } catch (error) {
        console.error('❌ 错误:', error.message);
        process.exit(1);
    }
}

// 如果直接运行此文件
if (require.main === module) {
    main().catch(error => {
        console.error('未捕获的错误:', error);
        process.exit(1);
    });
}

module.exports = DataAnalysisSkill;
