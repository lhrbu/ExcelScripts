function main(workbook: ExcelScript.Workbook) {
    const worksheet = workbook.getWorksheet("Sample tracking list preview")
    const testOrder =  worksheet.getRange("G1").getValue()
    const url = `https://backtolordinate-sampletngsystem-eejlbdevhn.cn-shanghai.fcapp.run/${encodeURIComponent(testOrder)}`
    fetch(url,{method:"POST"})
  }