// 配置部分
const columnWidths = {
    '医疗机构编码': 6.25,
    '医疗机构名称': 10,
    '患者姓名': 4,
    '患者性别': 5.8,
    '险种类型': 8,
    '结算日期': 6.6,
    '医保目录名称': 4,
    '规则名称': 9,
    '疑似违规内容': 13.25,
    '疑似违规金额': 6,
    '初审意见': 15,
    '复审意见': 15,
    '终审意见': 15,
    '申诉意见': 6.8,
    '终审结论': 3.4,
    '扣款金额（元）': 5.7,
    '终审时间': 10,
    '二次反馈': 4,
    '备注': 6
};

const columnsToKeep = [
    '医疗机构编码', '医疗机构名称', '患者姓名', '患者性别', '险种类型',
    '结算日期', '医保目录名称', '规则名称', '疑似违规内容',
    '疑似违规金额', '初审意见', '申诉意见', '复审意见',
    '终审结论', '终审意见', '扣款金额（元）', '终审时间',
    '二次反馈', '备注'
];

// 创建状态更新函数生成器
function createStatusUpdater(button, input) {
    return function() {
        button.disabled = !input.files.length;
        if (input.files.length) {
            status.textContent = `已选择文件: ${input.files[0].name}`;
            status.className = 'success';
        } else {
            status.textContent = '';
        }
    };
}

// 创建文件处理处理器生成器
function createUploadHandler(dropZone, fileInput, updateButtonState) {
    dropZone.ondragover = (e) => {
        e.preventDefault();
        dropZone.classList.add('dragover');
    };

    dropZone.ondragleave = () => {
        dropZone.classList.remove('dragover');
    };

    dropZone.ondrop = (e) => {
        e.preventDefault();
        dropZone.classList.remove('dragover');
        const files = e.dataTransfer.files;
        if (files.length > 0) {
            fileInput.files = files;
            updateButtonState();
        }
    };

    dropZone.onclick = () => fileInput.click();
    fileInput.onchange = updateButtonState;
}

// 处理各家分表的函数
async function processExcel(file) {
    const fileName = file.name;
    const dateMatch = fileName.match(/(\d{4})年?(\d{1,2})月?/);
    let year = '', month = '';
    let mainTitle = "淮安经济技术开发区智能审核扣款明细统计表";
    
    if (dateMatch) {
        year = dateMatch[1];
        month = dateMatch[2].replace(/^0+/, '');
        mainTitle = `淮安经济技术开发区${year}年${month}月智能审核扣款明细统计表`;
    } else {
        const altMatch = fileName.match(/(\d{4})(\d{2})/);
        if (altMatch) {
            year = altMatch[1];
            month = altMatch[2].replace(/^0+/, '');
            mainTitle = `淮安经济技术开发区${year}年${month}月智能审核扣款明细统计表`;
        }
    }

    const data = await file.arrayBuffer();
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(data);
    const worksheet = workbook.getWorksheet(1);

    const groupedData = {};
    worksheet.eachRow((row, rowNumber) => {
        if (rowNumber <= 2) return;

        const rowData = {};
        row.eachCell((cell, colNumber) => {
            const header = worksheet.getRow(2).getCell(colNumber).value;
            if (header) rowData[header.trim()] = cell.value;
        });

        const institutionCode = rowData['医疗机构编码'];
        const institutionName = rowData['医疗机构名称'];
        // 直接使用原始金额，不过滤负数
        const amount = parseFloat(rowData['扣款金额']) || 0;

        // 只检查必要字段是否存在，不做金额判断
        if (institutionCode && institutionName) {
            if (!groupedData[institutionCode]) {
                groupedData[institutionCode] = {
                    name: institutionName,
                    data: []
                };
            }

            const newRow = {};
            columnsToKeep.forEach(col => {
                if (col === '扣款金额（元）') {
                    newRow[col] = amount.toFixed(2);
                } else if (col === '二次反馈') {
                    newRow[col] = '无异议';
                } else if (col === '备注') {
                    newRow[col] = '';
                } else {
                    newRow[col] = rowData[col];
                }
            });

            groupedData[institutionCode].data.push(newRow);
        }
    });

    const zip = new JSZip();

    for (const [code, institution] of Object.entries(groupedData)) {
        const wb = new ExcelJS.Workbook();
        const ws = wb.addWorksheet('Sheet1', {
            pageSetup: {
                orientation: 'landscape',
                fitToPage: true,
                fitToHeight: 1,
                fitToWidth: 1
            }
        });

        columnsToKeep.forEach((col, index) => {
            ws.getColumn(index + 1).width = columnWidths[col];
        });

        const titleRow = ws.addRow([mainTitle]);
        ws.mergeCells(1, 1, 1, columnsToKeep.length);
        titleRow.height = 65;
        titleRow.font = { name: '方正小标宋_GBK', size: 24, bold: true };
        titleRow.alignment = { vertical: 'middle', horizontal: 'center' };

        // 修改标题行的边框设置 - 需要设置合并区域内所有单元格的边框
        for (let i = 1; i <= columnsToKeep.length; i++) {
            const cell = ws.getCell(1, i);
            // 设置每个单元格的边框
            if (i === 1) {
                // 第一个单元格
                cell.border = {
                    top: { style: 'none' },
                    left: { style: 'none' },
                    bottom: { style: 'thin' },
                    right: { style: 'none' }
                };
            } else if (i === columnsToKeep.length) {
                // 最后一个单元格
                cell.border = {
                    top: { style: 'none' },
                    left: { style: 'none' },
                    bottom: { style: 'thin' },
                    right: { style: 'none' }
                };
            } else {
                // 中间的单元格
                cell.border = {
                    top: { style: 'none' },
                    left: { style: 'none' },
                    bottom: { style: 'thin' },
                    right: { style: 'none' }
                };
            }
        }

        const headerRow = ws.addRow(columnsToKeep);
        headerRow.font = { name: '黑体', size: 12 };
        headerRow.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };

        institution.data.forEach(data => {
            const row = ws.addRow(columnsToKeep.map(col => data[col]));
            row.eachCell((cell, colNumber) => {
                cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                
                const columnName = columnsToKeep[colNumber - 1];
                
                // 扩展使用 Times New Roman 字体的条件
                if (typeof cell.value === 'number' || 
                    columnName === '医疗机构编码' ||
                    columnName === '结算日期' ||
                    columnName === '扣款金额（元）' ||
                    columnName === '疑似违规金额' ||
                    columnName === '终审时间' ||
                    (typeof cell.value === 'string' && !isNaN(cell.value)) ||
                    cell.value instanceof Date ||
                    (typeof cell.value === 'string' && /^\d{4}[-/年]\d{1,2}[-/月]\d{1,2}日?$/.test(cell.value))) {
                    
                    cell.font = { name: 'Times New Roman', size: 11 };
                    if (typeof cell.value === 'number') {
                        cell.numFmt = '0.00';
                    }
                } else {
                    cell.font = { name: '方正仿宋_GBK', size: 11 };
                }
                
                // 为数据行设置所有边框
                cell.border = {
                    top: { style: 'thin' },
                    left: { style: 'thin' },
                    bottom: { style: 'thin' },
                    right: { style: 'thin' }
                };
            });
        });

        // 表头和数据行保持四周边框
        headerRow.eachCell(cell => {
            cell.border = {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            };
        });

        const amountColumnIndex = columnsToKeep.indexOf('扣款金额（元）') + 1;
        const totalAmount = institution.data.reduce((sum, row) => {
            return sum + parseFloat(row['扣款金额（元）']);
        }, 0).toFixed(2);

        const totalRow = ws.addRow(Array(columnsToKeep.length).fill(''));
        totalRow.getCell(amountColumnIndex - 1).value = '违规总金额：';
        totalRow.getCell(amountColumnIndex - 1).font = { name: '方正仿宋_GBK', size: 11 };
        totalRow.getCell(amountColumnIndex).value = totalAmount;
        totalRow.getCell(amountColumnIndex).numFmt = '0.00';
        totalRow.getCell(amountColumnIndex).font = { name: 'Times New Roman', size: 11 };

        ws.addRow(Array(columnsToKeep.length).fill(''));

        const signRow = ws.addRow(Array(columnsToKeep.length).fill(''));
        signRow.getCell(amountColumnIndex - 3).value = '经办人签字：';
        signRow.getCell(amountColumnIndex - 3).font = { name: '方正仿宋_GBK', size: 11 };
        signRow.getCell(amountColumnIndex + 1).value = '盖章：';
        signRow.getCell(amountColumnIndex + 1).font = { name: '方正仿宋_GBK', size: 11 };

        [totalRow, signRow].forEach(row => {
            row.eachCell(cell => {
                cell.border = null;
            });
        });

        const buffer = await wb.xlsx.writeBuffer();
        zip.file(`${institution.name}.xlsx`, buffer);
    }

    const content = await zip.generateAsync({type: "blob"});
    saveAs(content, year && month ? `医院分表_${year}${month}.zip` : "医院分表.zip");
}

// 处理汇总表的函数
async function processSummaryExcel(file) {
    const fileName = file.name;
    const dateMatch = fileName.match(/(\d{4})年?(\d{1,2})月?/);
    let year = '', month = '';
    
    if (dateMatch) {
        year = dateMatch[1];
        month = dateMatch[2].replace(/^0+/, '');
    } else {
        const altMatch = fileName.match(/(\d{4})(\d{2})/);
        if (altMatch) {
            year = altMatch[1];
            month = altMatch[2].replace(/^0+/, '');
        }
    }

    const data = await file.arrayBuffer();
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(data);
    const worksheet = workbook.getWorksheet(1);

    const rows = [];
    worksheet.eachRow((row, rowNumber) => {
        if (rowNumber <= 2) return;

        const rowData = {};
        row.eachCell((cell, colNumber) => {
            const header = worksheet.getRow(2).getCell(colNumber).value;
            if (header) rowData[header.trim()] = cell.value;
        });

        // 删除金额判断，处理所有记录
        if (rowData['医疗机构编码'] && rowData['医疗机构名称'] && 
            rowData['险种类型']) {
            const amount = parseFloat(rowData['扣款金额']) || 0;
            rows.push({
                '医疗机构编码': rowData['医疗机构编码'],
                '医疗机构名称': rowData['医疗机构名称'],
                '扣款金额': amount,
                '险种类型': rowData['险种类型'],
                '人次': 1  // 每条记录都计算一次人次
            });
        }
    });

    const hospitalNameMap = {
        '淮安经济技术开发区医院（淮安汉方医院管理有限公司）': '淮安经济技术开发区医院',
        '枚乘路社区卫生服务中心': '淮安经济技术开发区枚乘街道卫生院'
    };

    const categories = {
        '职工': '职工基本医疗保险',
        '居民': '城乡居民基本医疗保险'
    };

    const zip = new JSZip();

    for (const [category, insuranceType] of Object.entries(categories)) {
        const isEmployee = category === '职工';

        // 在数据分组时保持所有记录
        const typeData = rows.filter(row => {
            return isEmployee ? 
                row['险种类型'].includes('职工') : 
                row['险种类型'].includes('居民');
        });

        const groupedData = {};
        typeData.forEach(row => {
            const code = row['医疗机构编码'];
            if (!groupedData[code]) {
                groupedData[code] = {
                    '医疗机构编码': code,
                    '医疗机构名称': hospitalNameMap[row['医疗机构名称']] || row['医疗机构名称'],
                    '扣款金额': 0,
                    '人次': 0
                };
            }
            groupedData[code]['扣款金额'] += row['扣款金额'];
            groupedData[code]['人次'] += row['人次'];
        });

        const wb = new ExcelJS.Workbook();
        const ws = wb.addWorksheet('Sheet1');

        ws.getColumn(1).width = 8.4;   // 序号
        ws.getColumn(2).width = 20;    // 医疗机构编码
        ws.getColumn(3).width = 56;    // 医疗机构名称
        ws.getColumn(4).width = 12;    // 扣款金额
        ws.getColumn(5).width = 12;    // 人次

        const title = `淮安经济技术开发区智能审核${year}年${month}月扣款统计表（${insuranceType}）`;
        const titleRow = ws.addRow([title]);
        ws.mergeCells(1, 1, 1, 5);
        titleRow.height = 65;
        titleRow.font = { name: '方正小标宋_GBK', size: 22 };
        titleRow.alignment = { vertical: 'middle', horizontal: 'center' };

        const unitRow = ws.addRow(['单位：人次/元']);
        ws.mergeCells(2, 1, 2, 5);
        unitRow.font = { name: '方正仿宋_GBK', size: 11 };
        unitRow.alignment = { vertical: 'middle', horizontal: 'right' };

        const headers = ['序号', '医疗机构编码', '医疗机构名称', '扣款金额', '人次'];
        const headerRow = ws.addRow(headers);
        headerRow.font = { name: '黑体', size: 11 };
        headerRow.alignment = { vertical: 'middle', horizontal: 'center' };

        let totalAmount = 0;
        let totalCount = 0;
        Object.values(groupedData).forEach((data, index) => {
            totalAmount += data['扣款金额'];
            totalCount += data['人次'];
            
            const row = ws.addRow([
                index + 1,
                data['医疗机构编码'],
                data['医疗机构名称'],
                data['扣款金额'].toFixed(2),
                data['人次']
            ]);

            row.eachCell((cell, colNumber) => {
                if (colNumber === 2 || colNumber === 4 || colNumber === 5) {
                    cell.font = { name: 'Times New Roman', size: 11 };
                } else {
                    cell.font = { name: '方正仿宋_GBK', size: 11 };
                }
                
                cell.border = {
                    top: { style: 'thin' },
                    left: { style: 'thin' },
                    bottom: { style: 'thin' },
                    right: { style: 'thin' }
                };
                
                cell.alignment = { vertical: 'middle', horizontal: 'center' };
            });
        });

        const totalRow = ws.addRow(['合计', '', '', totalAmount.toFixed(2), totalCount]);
        ws.mergeCells(ws.rowCount, 1, ws.rowCount, 3);
        totalRow.eachCell((cell, colNumber) => {
            if (colNumber === 4 || colNumber === 5) {
                cell.font = { name: 'Times New Roman', size: 11 };
            } else {
                cell.font = { name: '方正仿宋_GBK', size: 11 };
            }
            cell.alignment = { vertical: 'middle', horizontal: 'center' };
            cell.border = {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            };
        });

        const signRow = ws.addRow(['', '审批人：', '复核：', '初审：', '']);
        signRow.eachCell(cell => {
            cell.font = { name: '方正仿宋_GBK', size: 11 };
            cell.alignment = { vertical: 'middle', horizontal: 'left' };
        });

        headerRow.eachCell(cell => {
            cell.border = {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            };
        });

        const buffer = await wb.xlsx.writeBuffer();
        const fileName = `淮安经济技术开发区智能审核${year}年${month}月扣款统计表（${insuranceType}）.xlsx`;
        zip.file(fileName, buffer);
    }

    const content = await zip.generateAsync({type: "blob"});
    saveAs(content, `职工居民汇总表_${year}${month}.zip`);
}

// 添加新的处理函数
async function processLocationSummaryExcel(file) {
    const fileName = file.name;
    const dateMatch = fileName.match(/(\d{4})年?(\d{1,2})月?/);
    let year = '', month = '';
    
    if (dateMatch) {
        year = dateMatch[1];
        month = dateMatch[2].replace(/^0+/, '');
    } else {
        const altMatch = fileName.match(/(\d{4})(\d{2})/);
        if (altMatch) {
            year = altMatch[1];
            month = altMatch[2].replace(/^0+/, '');
        }
    }

    const data = await file.arrayBuffer();
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(data);
    const worksheet = workbook.getWorksheet(1);

    const rows = [];
    worksheet.eachRow((row, rowNumber) => {
        if (rowNumber <= 2) return;

        const rowData = {};
        row.eachCell((cell, colNumber) => {
            const header = worksheet.getRow(2).getCell(colNumber).value;
            if (header) rowData[header.trim()] = cell.value;
        });

        // 只检查必要字段是否存在，不做金额判断
        if (rowData['医疗机构编码'] && rowData['医疗机构名称'] && 
            rowData['险种类型']) {
            // 直接使用原始金额，不过滤负数
            const amount = parseFloat(rowData['扣款金额']) || 0;
            rows.push({
                '医疗机构编码': rowData['医疗机构编码'],
                '医疗机构名称': rowData['医疗机构名称'],
                '扣款金额': amount,  // 保持原始金额
                '险种类型': rowData['险种类型'],
                '人次': 1,
                '是否本地': rowData['是否本地']
            });
        }
    });

    const categories = {
        '本地-职工': '职工基本医疗保险',
        '本地-居民': '城乡居民基本医疗保险',
        '非本地-职工': '职工基本医疗保险',
        '非本地-居民': '城乡居民基本医疗保险'
    };

    const zip = new JSZip();

    for (const [category, insuranceType] of Object.entries(categories)) {
        // 解析类别
        const isLocal = category.startsWith('本地');
        const isEmployee = category.includes('职工');

        // 根据"是否本地"列和险种类型筛选数据
        const typeData = rows.filter(row => {
            const matchesInsurance = isEmployee ? 
                row['险种类型'].includes('职工') : 
                row['险种类型'].includes('居民');
            const matchesLocation = isLocal ? 
                row['是否本地'] === '是' : 
                row['是否本地'] === '否';
            return matchesInsurance && matchesLocation;
        });

        const groupedData = {};
        typeData.forEach(row => {
            const code = row['医疗机构编码'];
            if (!groupedData[code]) {
                groupedData[code] = {
                    '医疗机构编码': code,
                    '医疗机构名称': row['医疗机构名称'],
                    '扣款金额': 0,
                    '人次': 0
                };
            }
            groupedData[code]['扣款金额'] += row['扣款金额'];
            groupedData[code]['人次'] += row['人次'];
        });

        const customOrder = ['H32087100006', 'H32087100010', 'H32087100021', 'H32087100196', 'H32087101766'];
        let sortedData = Object.values(groupedData);
        sortedData.sort((a, b) => {
            const aIndex = customOrder.indexOf(a['医疗机构编码']);
            const bIndex = customOrder.indexOf(b['医疗机构编码']);
            if (aIndex === -1 && bIndex === -1) return a['医疗机构编码'].localeCompare(b['医疗机构编码']);
            if (aIndex === -1) return 1;
            if (bIndex === -1) return -1;
            return aIndex - bIndex;
        });

        const wb = new ExcelJS.Workbook();
        const ws = wb.addWorksheet('Sheet1');

        ws.getColumn(1).width = 8.4;   // 序号
        ws.getColumn(2).width = 20;    // 医疗机构编码
        ws.getColumn(3).width = 56;    // 医疗机构名称
        ws.getColumn(4).width = 12;    // 扣款金额
        ws.getColumn(5).width = 12;    // 人次

        const title = `淮安经济技术开发区智能审核${year}年${month}月扣款统计表（${insuranceType}）`;
        const titleRow = ws.addRow([title]);
        ws.mergeCells(1, 1, 1, 5);
        titleRow.height = 65;
        titleRow.font = { name: '方正小标宋_GBK', size: 22 };
        titleRow.alignment = { vertical: 'middle', horizontal: 'center' };

        const unitRow = ws.addRow(['单位：人次/元']);
        ws.mergeCells(2, 1, 2, 5);
        unitRow.font = { name: '方正仿宋_GBK', size: 11 };
        unitRow.alignment = { vertical: 'middle', horizontal: 'right' };

        const headers = ['序号', '医疗机构编码', '医疗机构名称', '扣款金额', '人次'];
        const headerRow = ws.addRow(headers);
        headerRow.font = { name: '黑体', size: 11 };
        headerRow.alignment = { vertical: 'middle', horizontal: 'center' };

        let totalAmount = 0;
        let totalCount = 0;
        sortedData.forEach((data, index) => {
            totalAmount += data['扣款金额'];
            totalCount += data['人次'];
            
            const row = ws.addRow([
                index + 1,
                data['医疗机构编码'],
                data['医疗机构名称'],
                data['扣款金额'].toFixed(2),
                data['人次']
            ]);

            row.eachCell((cell, colNumber) => {
                if (colNumber === 2 || colNumber === 4 || colNumber === 5) {
                    cell.font = { name: 'Times New Roman', size: 11 };
                } else {
                    cell.font = { name: '方正仿宋_GBK', size: 11 };
                }
                
                cell.border = {
                    top: { style: 'thin' },
                    left: { style: 'thin' },
                    bottom: { style: 'thin' },
                    right: { style: 'thin' }
                };
                
                cell.alignment = { vertical: 'middle', horizontal: 'center' };
            });
        });

        const totalRow = ws.addRow(['合计', '', '', totalAmount.toFixed(2), totalCount]);
        ws.mergeCells(ws.rowCount, 1, ws.rowCount, 3);
        totalRow.eachCell((cell, colNumber) => {
            if (colNumber === 4 || colNumber === 5) {
                cell.font = { name: 'Times New Roman', size: 11 };
            } else {
                cell.font = { name: '方正仿宋_GBK', size: 11 };
            }
            cell.alignment = { vertical: 'middle', horizontal: 'center' };
            cell.border = {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            };
        });

        const signRow = ws.addRow(['', '审批人：', '复核：', '初审：', '']);
        signRow.eachCell(cell => {
            cell.font = { name: '方正仿宋_GBK', size: 11 };
            cell.alignment = { vertical: 'middle', horizontal: 'left' };
        });

        headerRow.eachCell(cell => {
            cell.border = {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            };
        });

        const buffer = await wb.xlsx.writeBuffer();
        const locationPrefix = isLocal ? '本地' : '非本地';
        const fileName = `淮安经济技术开发区智能审核${year}年${month}月${locationPrefix}扣款统计表（${insuranceType}）.xlsx`;
        zip.file(fileName, buffer);
    }

    const content = await zip.generateAsync({type: "blob"});
    saveAs(content, `按本地分类汇总表_${year}${month}.zip`);
}

// 获取 DOM 元素
const dropZone = document.getElementById('dropZone');
const fileInput = document.getElementById('fileInput');
const processButton = document.getElementById('processButton');
const dropZone2 = document.getElementById('dropZone2');
const fileInput2 = document.getElementById('fileInput2');
const processButton2 = document.getElementById('processButton2');
const status = document.getElementById('status');
const loading = document.getElementById('loading');

// 创建状态更新函数
const updateButtonState = createStatusUpdater(processButton, fileInput);
const updateButton2State = createStatusUpdater(processButton2, fileInput2);

// 应用上传处理器
createUploadHandler(dropZone, fileInput, updateButtonState);
createUploadHandler(dropZone2, fileInput2, updateButton2State);

// 添加标签切换功能
document.querySelectorAll('.tab').forEach(tab => {
    tab.addEventListener('click', () => {
        document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
        tab.classList.add('active');

        document.querySelectorAll('.function-area').forEach(area => area.classList.remove('active'));
        document.getElementById(tab.dataset.target).classList.add('active');

        status.textContent = '';
        loading.style.display = 'none';
    });
});

// 处理按钮点击事件
processButton.onclick = async () => {
    const file = fileInput.files[0];
    if (!file) return;

    loading.style.display = 'block';
    processButton.disabled = true;
    status.textContent = '';

    try {
        await processExcel(file);
        status.textContent = '文件处理成功！';
        status.className = 'success';
    } catch (error) {
        status.textContent = `错误：${error.message}`;
        status.className = 'error';
    } finally {
        loading.style.display = 'none';
        processButton.disabled = false;
        fileInput.value = '';
        updateButtonState();
    }
};

processButton2.onclick = async () => {
    const file = fileInput2.files[0];
    if (!file) return;

    loading.style.display = 'block';
    processButton2.disabled = true;
    status.textContent = '';

    try {
        await processSummaryExcel(file);
        status.textContent = '汇总表处理成功！';
        status.className = 'success';
    } catch (error) {
        status.textContent = `错误：${error.message}`;
        status.className = 'error';
    } finally {
        loading.style.display = 'none';
        processButton2.disabled = false;
        fileInput2.value = '';
        updateButton2State();
    }
};

// 添加新的 DOM 元素
const dropZone3 = document.getElementById('dropZone3');
const fileInput3 = document.getElementById('fileInput3');
const processButton3 = document.getElementById('processButton3');

// 创建新的状态更新函数
const updateButton3State = createStatusUpdater(processButton3, fileInput3);

// 应用上传处理器
createUploadHandler(dropZone3, fileInput3, updateButton3State);

// 添加新的处理按钮事件
processButton3.onclick = async () => {
    const file = fileInput3.files[0];
    if (!file) return;

    loading.style.display = 'block';
    processButton3.disabled = true;
    status.textContent = '';

    try {
        await processLocationSummaryExcel(file);
        status.textContent = '按本地分类汇总表处理成功！';
        status.className = 'success';
    } catch (error) {
        status.textContent = `错误：${error.message}`;
        status.className = 'error';
    } finally {
        loading.style.display = 'none';
        processButton3.disabled = false;
        fileInput3.value = '';
        updateButton3State();
    }
};