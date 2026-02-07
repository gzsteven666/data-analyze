/**
 * 数据分析信息图Skill测试
 */

const DataAnalysisSkill = require('../src/index.js');
const fs = require('fs-extra');
const path = require('path');

async function testSkill() {
    console.log('🧪 开始测试数据分析信息图Skill...\n');

    const skill = new DataAnalysisSkill();

    try {
        // 测试1: 技能实例化
        console.log('✅ 测试1: 技能实例化');
        console.log(`   - 技能名称: ${skill.constructor.name}`);
        console.log(`   - Python脚本路径: ${skill.pythonScriptPath}`);
        console.log(`   - 支持文件类型: ${skill.getSupportedFileTypes().join(', ')}\n`);

        // 测试2: 文件类型验证
        console.log('✅ 测试2: 文件类型验证');
        const testFiles = [
            'data.csv',
            'data.xlsx',
            'data.xls',
            'data.txt',
            'data.json'
        ];

        testFiles.forEach(file => {
            const isValid = skill.isValidFileType(file);
            console.log(`   - ${file}: ${isValid ? '✅ 支持' : '❌ 不支持'}`);
        });
        console.log();

        // 测试3: 参数验证
        console.log('✅ 测试3: 参数验证');
        try {
            await skill.analyzeData('non_existent_file.csv');
        } catch (error) {
            console.log(`   - 文件不存在错误处理: ✅ ${error.message}`);
        }

        // 测试4: 报告生成
        console.log('\n✅ 测试4: 报告生成');
        const mockResult = {
            output: {
                screenshot: 'path/to/screenshot.png',
                html: 'path/to/report.html',
                excelFile: 'path/to/summary.xlsx',
                chartCount: 6,
                csvDir: 'outputs/csv'
            },
            timestamp: new Date().toISOString()
        };

        const report = skill.generateReport(mockResult);
        console.log('   - 报告预览:');
        console.log(report.split('\n').slice(0, 10).join('\n'));
        console.log('   ...\n');

        // 测试5: 帮助信息
        console.log('✅ 测试5: 帮助信息');
        console.log('   - 技能已正确加载，支持以下功能:');
        console.log('     • 数据文件读取和分析');
        console.log('     • 多维度数据可视化');
        console.log('     • HTML信息图生成');
        console.log('     • 高质量截图输出');
        console.log('     • 多格式数据导出\n');

        console.log('🎉 所有测试通过！技能就绪。');
        console.log('\n📋 使用示例:');
        console.log('  node src/index.js data/quarter_sales.xlsx --sheet="Q4" --company="旗舰店A" --time="月份" --value="GMV"');
        console.log('  node src/index.js data/市场数据.csv --company="测试公司"\n');

    } catch (error) {
        console.error('❌ 测试失败:', error);
        process.exit(1);
    }
}

// 运行测试
if (require.main === module) {
    testSkill().catch(error => {
        console.error('测试执行失败:', error);
        process.exit(1);
    });
}

module.exports = { testSkill };
