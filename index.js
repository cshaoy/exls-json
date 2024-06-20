const ExcelJS = require('exceljs');
const fs = require('fs');

async function readExcelToJson(filePath) {
    try {
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(filePath);
        const jsonData = [];
        workbook.eachSheet((worksheet, sheetId) => {

            console.log(sheetId)
            if (!worksheet) {
                throw new Error('Worksheet not found in the Excel file.');
            }



            worksheet.eachRow({ includeEmpty: true, firstRow: 2 }, (row, rowNum) => {
                const category = row.getCell(1).value;
                const item1 = row.getCell(2).value;
                const item2 = row.getCell(3).value;
                const item3 = row.getCell(4).value;
                function findOrCreate(arr, value) {
                    let item = arr.find(t => t.value === value);
                    if (!item) {
                        item = { value, label: value, children: [] };
                        arr.push(item);
                    }
                    return item;
                }
            
                if (category) {
                    let li = findOrCreate(jsonData, category);
                    
                    if (item1) {
                        let children1 = findOrCreate(li.children, item1);
                        
                        if (item2) {
                            let children2 = findOrCreate(children1.children, item2);
                            
                            if (item3) {
                                findOrCreate(children2.children, item3);
                            }
                        }
                    }
                }
                // if (category) {
                //     let li = jsonData.find(t => t.value == category);
                //     if (li) {
                //     } else {
                //         li = { value: category, label: category, children: [] }
                //         jsonData.push(li)
                //     }
                //     if (item1) {
                //         let children1 = li.children.find(t => t.value == item1);
                //         if (children1) {
                //             children1 = children1.children;
                //         } else {
                //             children1 = []
                //             li.children.push({ value: item1, label: item1, children: children1 })
                //         }
                //         if (item2) {
                //             let children2 = children1.find(t => t.value == item2);
                //             if (children2) {
                //                 children2 = children2.children;
                //             } else {
                //                 children2 = []
                //                 children1.push({ value: item2, label: item2, children: children2 });
                //             }
                //             if (item3) {
                //                 let children3 = children2.find(t => t.value == item3);
                //                 if (children3) {

                //                 } else {
                //                     children2.push({ value: item3, label: item3, })
                //                 }
                //             }
                //         }
                //     }
                // }
            });
        })


        return jsonData;
    } catch (error) {
        console.error('Error reading Excel:', error);
        return null;
    }
}

// Usage
readExcelToJson('./保险公司分类.xlsx')
    .then(jsonData => {
        if (jsonData) {
            // console.log(jsonData);
            const jsonFilePath = './output.json'; // JSON 文件路径
            fs.writeFile(jsonFilePath, JSON.stringify(jsonData, null, 2), err => {
                if (err) {
                    console.error('Error writing JSON file:', err);
                } else {
                    console.log('JSON file saved successfully.');
                }
            });
        } else {
            console.log('Failed to read Excel file.');
        }
    })
    .catch(error => {
        console.error('Error:', error);
    });
