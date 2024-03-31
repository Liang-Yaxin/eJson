const
    xlsx = require('node-xlsx'),
    fs = require('fs-extra'),
    path = require('path'),
    Config = require('./config/index'),
    Tool = require('./tool/index'),
    Text = require('./text/index'),
    Log = require('./log/index'),
    Table = require('cli-table'),
    chalk = require('chalk');

// 将excel转换成json
function convertToJson(options) {
    /**
     * @options 用户终端输入的指令
     * @originalXlsxData 最原始的excel表格数据
     * @totalSheet 去掉空的 originalXlsxData 内可能存在的空数据
     * @totalSheetNum 总共有几个表（去掉空的sheet，有些表格里面没东西）
     * @finalOutPath 输出的json文件地址
     * @finalkeys excel每行对应的key
     * @finalJsonArr 用来存放生成的每条json数据的容器
     * @finalRange 用来存放读取表的范围
     * @table 用表格形式来描述输出文件的信息
     */
    console.log('options', options)
    const { input: userInput, startRow, range } = options,
    originalXlsxData = xlsx.parse(`${path.resolve('./', userInput)}`),
        totalSheet = originalXlsxData.filter(item => item.data.length !== 0),
        totalSheetNum = totalSheet.length,
        finalOutPath = Config.defaluOutPath,
        finalStartRow = Tool.getFinalStartRow(startRow, totalSheetNum), // 返回每个表的起始行 数组
        finalRange = range || Config.defaultRange,
        finalJsonArr = [],
        table = new Table({
            head: [Text.infoText('Json file name'), Text.infoText('File location')],
        });

    Log('Json data being generated...', 'info');
    console.log();

    // 最外层
    for (let index = 0; index < totalSheet.length; index++) {
        console.log(finalRange)
        if (finalRange.includes(index + 1) || finalRange === 'all') {
            const
                everySheetArr = totalSheet[index], // 每个表格
                { name, data } = everySheetArr,
                meterData = data[finalStartRow[index] - 2],
                totalRowData = data.splice(finalStartRow[index] - 1).filter(item => item.length !== 0), // 要开始读取数据做渲染的地方（所有行的数据）
                everyRowLengthArr = totalRowData.map(item => item.length), // 每一行的列数
                maxRowLength = Math.max(...everyRowLengthArr), // 求行最大的列数
                singleJsonObj = {}, // 每个表格作为一个数据（对象形式）
                totalRowJsonArr = []; // 用来存放excel每行组装成的对象数据（singleRowObj）

            console.log(chalk.hex('#A37FFF')(`Currently reading data from line ${Text.infoText(finalStartRow[index][0])} of ${Text.infoText('sheet ' + (index + 1))}`));

            // 把每行数组组装成对象形式
            for (let index2 = 0; index2 < totalRowData.length; index2++) {
                const
                    singleRowRenderArr = totalRowData[index2], // 每一行的值（形式为数组）
                    singleRowObj = {}, // 用对象形式存放每一行的值
                    MAXLENGTH = singleRowRenderArr.length, // 以用户输入的-k 长度做渲染
                    defalutValue = '';

                // 每行里的每个具体数据
                for (let index3 = 0; index3 < MAXLENGTH; index3++) {
                    const value = singleRowRenderArr[index3] || defalutValue;
                    const keyValue = meterData[index3];

                    singleRowObj[keyValue] = value;
                }

                totalRowJsonArr.push(singleRowObj);
            }

            // 添加每个表格名字
            singleJsonObj.xlsx_name = name;
            singleJsonObj.data = totalRowJsonArr;

            finalJsonArr.push(singleJsonObj);
        }
    }

    console.log(finalJsonArr);

    // 输出json文件
    finalJsonArr.forEach((item, index) => {
        console.log(item)
        let
            pathName = `${finalOutPath}`

        fs.outputJsonSync(`${pathName}/${item.xlsx_name}`, item);
        table.push([item.xlsx_name, pathName]);
    });


    // 打印最后的json文件地址
    Log('Json data generated successfully.', 'success');
    console.log(table.toString());
}

// 获取excel模板文件
function getTemplate(option) {
    const
        cliPath = __dirname,
        originalTemplate = fs.readdirSync(path.resolve(cliPath, `../template/`))[0],
        table = new Table({
            head: [Text.infoText('excel file name'), Text.infoText('File location')]
        });

    try {
        let pathName = '';

        if (!option) {
            pathName = path.resolve('./', Config.defaultOutTemplateName);
        } else {
            pathName = path.resolve('./', option);
        }

        fs.copySync(path.resolve(cliPath, `../template/`), `${path.resolve('./', pathName)}`);

        Log('Excel template file obtained successfully', 'success');
        table.push([originalTemplate, pathName]);
        console.log(table.toString());

    } catch (err) {
        log(err, 'error');
    }
}


module.exports = {
    convertToJson,
    getTemplate
};