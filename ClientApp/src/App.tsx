import React, { Component, useEffect } from 'react';
import { Button } from 'reactstrap';

export default () => {
  return (<div>
    <Button onClick={async () => {
      await Excel.run(async (context) => {
        await populateWeatherData().then(data => writeSheetData(context.workbook.worksheets.getActiveWorksheet(), data));
      })
    }}>插入数据</Button>
  </div>)

}

async function populateWeatherData() {
  const response = await fetch('weatherforecast');
  const data = await response.json();
  return data;
}

async function writeSheetData(sheet: Excel.Worksheet, data: any[]) {
  const titleCell = sheet.getCell(0, 0);
  titleCell.values = [["Weather Report"]];
  titleCell.format.font.name = "Century";
  titleCell.format.font.size = 26;
  // Create an array containing sample data
  const headerNames = ["Date", "TemperatureC", "TemperatureF", "Summary"];

  // Write the sample data to the specified range in the worksheet
  // and bold the header row
  const headerRow = titleCell.getOffsetRange(1, 0).getResizedRange(0, headerNames.length - 1);
  headerRow.values = [headerNames];
  headerRow.getRow(0).format.font.bold = true;

  const dataRange = headerRow.getOffsetRange(1, 0).getResizedRange(data.length - 1, 0);
  dataRange.values = data;

  titleCell.getResizedRange(0, headerNames.length - 1).merge();
  dataRange.format.autofitColumns();

  await sheet.context.sync();
}