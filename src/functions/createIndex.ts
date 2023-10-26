import { Workbook } from 'exceljs';
import rowBorder from './rowBorder';
import rowStyle from './rowStyle';
import methodColor from '../config/methodColors';
import mainColor from '../config/mainColor';

export default function createIndex(
  wb: Workbook,
  servers: any[],
  securitySchemes: any,
  pathsByTag: any,
): void {
  const sheet = wb.addWorksheet('Index', {
    pageSetup: {
      paperSize: 9,
      fitToPage: true,
    },
  });

  let workRow = 1;

  /**
   * widget 설정
   */
  sheet.getColumn('A').width = 12;
  sheet.getColumn('B').width = 50;
  sheet.getColumn('C').width = 12;
  sheet.getColumn('D').width = 30;
  sheet.getColumn('E').width = 12;

  /**
   * Server 목록
   */
  sheet.mergeCells(`A${workRow}`, `E${workRow}`);
  const serverCell = sheet.getCell(`A${workRow}`);
  serverCell.value = 'Servers';
  serverCell.font = { bold: true, size: 15 };
  serverCell.style = {
    alignment: {
      horizontal: 'center',
    },
  };
  serverCell.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: mainColor },
  };
  rowBorder(sheet, workRow, 'A', 'E');

  workRow = 2;
  for (let index = 0; index < servers.length; index++) {
    const row = index + workRow;

    sheet.mergeCells(`A${row}`, `E${row}`);
    const url = sheet.getCell(`A${row}`);
    url.value = servers[index].url + ' - ' + servers[index].description;
    rowBorder(sheet, row, 'A', 'E');
  }
  workRow += servers.length;

  /**
   * 인증 정보
   */
  workRow++;
  sheet.mergeCells(`A${workRow}`, `E${workRow}`);
  const authCell = sheet.getCell(`A${workRow}`);
  authCell.value = 'Authroization';
  authCell.font = { bold: true, size: 15 };
  authCell.style = {
    alignment: {
      horizontal: 'center',
    },
  };
  authCell.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: mainColor },
  };
  rowBorder(sheet, workRow, 'A', 'E');

  workRow++;
  const security = Object.values(securitySchemes)[0] as any;
  sheet.mergeCells(`A${workRow}`, `E${workRow}`);
  sheet.getCell(
    `A${workRow}`,
  ).value = `${security.scheme} (${security.type}, ${security.bearerFormat})`;
  rowBorder(sheet, workRow, 'A', 'E');

  /**
   * tag별 summary
   */
  workRow += 2; // 한줄 띄우기
  for (const [tag, value] of Object.entries(pathsByTag)) {
    // 태그 추가
    sheet.mergeCells(`A${workRow}`, `E${workRow}`);
    const tagCell = sheet.getCell(`A${workRow}`);
    tagCell.value = tag;
    rowStyle(
      sheet,
      workRow,
      'A',
      'E',
      { bold: true, size: 15 },
      {
        alignment: {
          horizontal: 'center',
        },
      },
      {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: mainColor },
      },
      {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' },
      },
    );

    // Column명
    workRow++;
    sheet.getCell(`A${workRow}`).value = 'index';
    sheet.getCell(`B${workRow}`).value = 'Path';
    sheet.getCell(`C${workRow}`).value = 'Method';
    sheet.getCell(`D${workRow}`).value = 'Summary';
    sheet.getCell(`E${workRow}`).value = 'Etc';
    rowStyle(
      sheet,
      workRow,
      'A',
      'E',
      { size: 12 },
      {
        alignment: {
          horizontal: 'center',
        },
      },
      {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: mainColor },
      },
      {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' },
      },
    );

    // api
    const apis = value as any[];
    for (let index = 0; index < apis.length; index++) {
      workRow++;
      const api = apis[index];
      sheet.getCell(`A${workRow}`).value = index;
      sheet.getCell(`A${workRow}`).style = {
        alignment: {
          horizontal: 'center',
        },
      };
      sheet.getCell(`B${workRow}`).value = api.path;
      sheet.getCell(`C${workRow}`).value = api.method.toUpperCase();
      sheet.getCell(`C${workRow}`).style = {
        alignment: {
          horizontal: 'center',
        },
        fill: {
          type: 'pattern',
          pattern: 'solid',
          // eslint-disable-next-line @typescript-eslint/ban-ts-comment
          // @ts-ignore
          fgColor: { argb: methodColor[api.method] },
        },
      };
      sheet.getCell(`D${workRow}`).value = api.summary;
      sheet.getCell(`E${workRow}`).value =
        api.deprecated ?? false ? 'Deprecated' : '';

      rowBorder(sheet, workRow, 'A', 'E');
    }

    // 한줄 띄우기
    workRow++;
    workRow++;
  }
}
