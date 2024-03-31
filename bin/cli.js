#!/usr/bin/env node

const { Command } = require('commander'),
    fs = require('fs-extra'),
    path = require('path'),
    pkg = require('../package.json'),
    Text = require('../src/text/index'),
    Core = require('../src/index'),
    Tool = require('../src/tool/index'),
    Log = require('../src/log/index'),
    Config = require('../src/config/index'),
    program = new Command();

// 基础配置
program
    .name(pkg.name)
    .description(`${Text.descText('=> Manage your json data better with visual excel sheets')}`)
    .version(Text.descText(Config.LOGO), '-v', Text.infoText('View current version'))
    .helpOption('-h, --help', Text.infoText('View help'));


// 转换功能
program
    .option('-i, --input [path]', Text.infoText('Path of excel to be converted'))
    .option('-s, --start-row [number]', Text.infoText('Read data from what row of sheet'))
    .option('-r, --range [string]', Text.infoText('Select the table you want to read'))
    .action(function(options) {
        console.log(options);
        console.log(options.range)
        try {
            // 验证-i
            if (options.input === true || !options.input) {
                throw '"-i" or "--input" is required';
            } else if (!Tool.isPath(options.input)) {
                throw 'The path where excel is located is incorrect';
            } else if (!fs.existsSync(path.resolve('./', options.input))) {
                throw 'excel file does not exist';
            }
            // 验证-s
            else if (options.startRow && !Tool.getValueType(options.startRow).includes('str')) {
                throw 'The value of "-s" or "--start-row" needs to be a string or an integer';
            }
            // 验证-r
            else if (options.range && !Tool.getValueType(options.range).includes('str')) {
                throw 'The value of "-r" or "--range" needs to be a string or an integer';
            }
            // 开始生成
            else {
                Core.convertToJson(options);
            }
        } catch (error) {
            Log(error, 'fail');
        }
    });

// 下载模板excel文件
program
    .command('gt')
    .description(Text.infoText('Get the excel template file'))
    .argument('[path]', Text.infoText('Excel template file output path'))
    .action(function(option) {
        // console.log(option);
        try {
            if (option && !Tool.isPath(option)) {
                throw 'Incorrect output path format';
            } else {
                Core.getTemplate(option);
            }
        } catch (error) {
            Log(error, 'fail');
        }
    });

// 默认显示帮助（没有输入任何参数）
if (process.argv && process.argv.length < 3) {
    program.help();
}

program.parse(process.argv);