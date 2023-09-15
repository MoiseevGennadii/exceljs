const axios = require('axios')
const Excel = require('exceljs')


axios.get('https://api.publicapis.org/entries')
  .then(res => {

    let workbook = new Excel.Workbook()
    let worksheet = workbook.addWorksheet('data')
    let data = res.data.entries[1];
    let data1 = res.data.entries.filter(err => err.HTTPS === true);
    let headerrow = worksheet.addRow(Object.keys(data));

    data1.sort((a, b) => {
      if (a.API < b.API) return -1;
      if (a.API > b.API) return 1;
      return 0;
    });

    data1.forEach(rowData => {
      const row = worksheet.addRow([rowData.API, rowData.Description, rowData.Auth, rowData.HTTPS, rowData.Cors, rowData.Link, rowData.Category]);
      row.eachCell((cell) => {
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' },
        };
      })
    })

    data1.forEach((row, index) => {
      const hyperlink = { text: row.Link, hyperlink: row.Link };
      worksheet.getCell(`F${index + 2}`).value = hyperlink;
    });

    worksheet.getColumn('F').font = {
      color: { argb: 'ff4f4fd9' },
      bold: true,
      underline: true
    }

    function setHeaderWidth(worksheet, width) {
      const numCols = worksheet.columns.length;
      for (let i = 1; i <= numCols; i++) {
        worksheet.getColumn(i).width = width;
      }
    }
    setHeaderWidth(worksheet, 16);

    headerrow.eachCell(cell => {
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' },
      }
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'BCF631' },

      }
      cell.alignment = {
        horizontal: 'center',
        vertical: 'center'
      }
      cell.font = {
        bold: false,
        name: 'Broadway'
      }
    })

    workbook.xlsx.writeFile('data.xlsx')
      .then(() => {
        console.log('File  saved saccessfully')
      })
      .catch(err => {
        console.log(err)
      })

  })
  .catch(err => {
    console.log(err)
  })
