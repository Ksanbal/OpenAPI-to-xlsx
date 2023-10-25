import { Workbook } from 'exceljs';

export default function createIndex(
  wb: Workbook,
  servers: any[],
  pathsByTag: any,
): void {
  const sheet = wb.addWorksheet('Index', {
    pageSetup: {
      paperSize: 9,
      fitToPage: true,
    },
  });

  // Server 목록
  sheet.mergeCells('A1', 'D1');
  const cellA1 = sheet.getCell('A1');
  cellA1.value = 'Servers';
  cellA1.font = { bold: true, size: 15 };
  cellA1.style = {
    alignment: {
      horizontal: 'center',
    },
  };
  cellA1.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: '00eef6fe' },
  };
  cellA1.border = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thin' },
    right: { style: 'thin' },
  };

  // sheet.addRows(
  //   servers.map((server) => {
  //     return [server.description, server.url];
  //   }),
  // );
}
