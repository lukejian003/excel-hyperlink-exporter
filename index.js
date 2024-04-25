const ExcelJS = require('exceljs');
const fs = require('fs');
const readline = require('readline');
const path = require('path');
const glob = require('glob');

const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout
});

rl.question('请输入包含超链接的Excel文件所在的路径：', (inputPath) => {
    const excelFiles = glob.sync(path.join(inputPath, '*.xlsx'));
    if (excelFiles.length === 0) {
        console.log('指定路径下未找到任何Excel文件');
        rl.close();
    } else {
        const inputFileName = excelFiles[0]; // 选择第一个Excel文件
        rl.question('请输入要输出的txt文件名：', (outputFileName) => {
            const outputFilePath = path.join(inputPath, outputFileName + '.txt');
            const workbook = new ExcelJS.Workbook();
            workbook.xlsx.readFile(inputFileName)
                .then(function() {
                    const worksheet = workbook.getWorksheet(1);
                    let hyperlinks = [];
                    worksheet.eachRow(function(row) {
                        row.eachCell(function(cell) {
                            if (cell.isHyperlink) {
                                hyperlinks.push(cell.hyperlink);
                            }
                        });
                    });
                    fs.writeFileSync(outputFilePath, hyperlinks.join('\n'), 'utf-8');
                    console.log('超链接已成功写入到文件：' + outputFilePath);
                    rl.close();
                })
                .catch(function(error) {
                    console.log('发生错误：', error);
                    rl.close();
                });
        });
    }
});
