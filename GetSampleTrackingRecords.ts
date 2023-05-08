function main(workbook: ExcelScript.Workbook,sampleTrackingRecords:SampleTrackingRecord[]) {
    const worksheet = workbook.getWorksheet("Sample tracking list preview")
  
    let filter = worksheet.getAutoFilter();
    if(filter)
    {
      filter.clearCriteria();
    }
  
    const usagedRange = worksheet.getUsedRange(true)
    if(usagedRange){
      usagedRange.getOffsetRange(1, 0).delete(ExcelScript.DeleteShiftDirection.left)
    }
  
    const titleRange = worksheet.getRange('A1:D1')
    SetTitleRowStyle(titleRange)
    titleRange.setValues([["TestOrderNo", "SampleNo", "LocationName", "Timestamp"]]);
    // await FetchAndImportSampleTrackingRecords("https://gettestorders-sampletngsystem-mfectnvalm.cn-shanghai.fcapp.run",
    //   worksheet)
    ImportRecords(sampleTrackingRecords,worksheet)
  }
  
  function ImportRecords(records: SampleTrackingRecord[], worksheet: ExcelScript.Worksheet)
  {
    const rows: (string | number)[][] = [];
    for (const record of records) {
      rows.push([record.TestOrderNo, record.SampleNo, record.LocationName,
        ConvertJsDateToExcelDate(record.Timestamp)]);
    }
    if(rows.length>0){
      const range = worksheet.getRange('A2').getResizedRange(rows.length - 1, rows[0].length - 1);
      range.getColumn(2).setNumberFormatLocal("@");
      range.getColumn(3).setNumberFormatLocal("yyyy/mm/dd h:mm;@")
      range.setValues(rows);
    }
  }
  
  function ConvertJsDateToExcelDate(jsDate:string)
  {
    return Date.parse(jsDate) / 1000 / 86400 + 25569
  }
  
  interface SampleTrackingRecord
  {
    TestOrderNo:string
    SampleNo:string
    LocationName:string
    Timestamp:string
  }
  
  function SetTitleRowStyle(titleRange: ExcelScript.Range) {
    titleRange.getFormat().getFill().setColor("92D050");
    titleRange.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
    titleRange.getFormat().setIndentLevel(0);
    titleRange.getFormat().getFont().setColor("262626");
  
    titleRange.getFormat().getRangeBorder(ExcelScript.BorderIndex.insideHorizontal).setStyle(ExcelScript.BorderLineStyle.continuous);
    titleRange.getFormat().getRangeBorder(ExcelScript.BorderIndex.insideHorizontal).setColor("000000");
    titleRange.getFormat().getRangeBorder(ExcelScript.BorderIndex.insideHorizontal).setWeight(ExcelScript.BorderWeight.thin);
    titleRange.getFormat().getRangeBorder(ExcelScript.BorderIndex.insideVertical).setStyle(ExcelScript.BorderLineStyle.continuous);
    titleRange.getFormat().getRangeBorder(ExcelScript.BorderIndex.insideVertical).setColor("000000");
    titleRange.getFormat().getRangeBorder(ExcelScript.BorderIndex.insideVertical).setWeight(ExcelScript.BorderWeight.thin);
    titleRange.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeBottom).setStyle(ExcelScript.BorderLineStyle.continuous);
    titleRange.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeBottom).setColor("000000");
    titleRange.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeBottom).setWeight(ExcelScript.BorderWeight.thin);
    titleRange.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeTop).setStyle(ExcelScript.BorderLineStyle.continuous);
    titleRange.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeTop).setColor("000000");
    titleRange.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeTop).setWeight(ExcelScript.BorderWeight.thin);
    titleRange.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeLeft).setStyle(ExcelScript.BorderLineStyle.continuous);
    titleRange.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeLeft).setColor("000000");
    titleRange.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeLeft).setWeight(ExcelScript.BorderWeight.thin);
    titleRange.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeRight).setStyle(ExcelScript.BorderLineStyle.continuous);
    titleRange.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeRight).setColor("000000");
    titleRange.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeRight).setWeight(ExcelScript.BorderWeight.thin);
  }