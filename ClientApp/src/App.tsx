import React from 'react';
import { Button } from 'reactstrap';

const App = () => {
  return (<div>
    <Button style={{ margin: 20 }} onClick={async () => {
      await Excel.run(async (context) => {
        await populateWeatherData().then(data => writeSheetData(context.workbook.worksheets.getActiveWorksheet(), data));
      })
    }}>插入数据</Button>
  </div>)

}

async function populateWeatherData() {
  const response = await fetch('weatherforecast');
  const data = await response.json();

  return Array.from(data, (item: any) => {
    return [item.date, item.temperatureC, item.temperatureF, item.summary]
  });
}

async function writeSheetData(sheet: Excel.Worksheet, data: any[]) {

  const titleCell = sheet.getCell(0, 0);
  titleCell.values = [["Weather Report"]];
  titleCell.format.font.name = "Century";
  titleCell.format.font.size = 26;
  const headerNames = ["Date", "TemperatureC", "TemperatureF", "Summary"];

  const headerRow = titleCell.getOffsetRange(1, 0).getResizedRange(0, headerNames.length - 1);
  headerRow.values = [headerNames];
  headerRow.getRow(0).format.font.bold = true;

  const dataRange = headerRow.getOffsetRange(1, 0).getResizedRange(data.length - 1, 0);
  dataRange.values = data;

  titleCell.getResizedRange(0, headerNames.length - 1).merge();
  dataRange.format.autofitColumns();

  await sheet.context.sync();
}

export default App;