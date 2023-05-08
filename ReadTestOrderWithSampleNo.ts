const worksheetName = "2023"

function main(workbook: ExcelScript.Workbook) {
    const worksheet = workbook.getWorksheet(worksheetName)
    const rowCount = worksheet.getUsedRange().getRowCount()

    const testOrderNos = worksheet.getRange(`A1:A${rowCount}`).getValues().map(item => item[0] as string)
    const samplesNoStart = worksheet.getRange(`I1:I${rowCount}`).getValues().map(item => item[0] as number)
    const samplesNoEnd = worksheet.getRange(`J1:J${rowCount}`).getValues().map(item => item[0] as number)
    const applicantNames = worksheet.getRange(`E1:E${rowCount}`).getValues().map(item => item[0] as string)
    const LabCoordinateNames = worksheet.getRange(`G1:G${rowCount}`).getValues().map(item => item[0] as string)
    const TestEngineerNames = worksheet.getRange(`M1:M${rowCount}`).getValues().map(item => item[0] as string)

    const initIndex = 3

    const result: Sample[] = []
    for (let i = initIndex; i <= rowCount; i++) {
        if (testOrderNos[i] && samplesNoStart[i] && samplesNoEnd[i]) {
            const sampleNoStart = samplesNoStart[i]
            const sampleNoEnd = samplesNoEnd[i]
            const applicantName = applicantNames[i]
            const labCoordinateName = RemoveBlankString(LabCoordinateNames[i])
            const testEngineerName = RemoveBlankString(TestEngineerNames[i])

            for (let sampleNo = sampleNoStart; sampleNo <= sampleNoEnd; sampleNo++) {
                result.push({
                    SampleNo: sampleNo.toString(),
                    TestOrderNo: testOrderNos[i],
                    ApplicantName: applicantName,
                    LabCoordinateName: labCoordinateName,
                    TestEngineerName: testEngineerName
                })
            }
        }
    }
    return result
}

interface Sample {
    SampleNo: string,
    TestOrderNo: string,
    ApplicantName?: string,
    LabCoordinateName?: string,
    TestEngineerName?: string
}

function RemoveBlankString(value: string) {
    if (value) {
        if (value.trim() === "") { return undefined }
        else { return value }
    } else {
        return undefined;
    }
}