import { Workbook } from 'exceljs';

export default function createCover(wb: Workbook, info: any): void {
  const sheet = wb.addWorksheet('Sheet1', {
    pageSetup: {
      paperSize: 9,
      fitToPage: true,
    },
  });

  // Title
  const titleCell = sheet.getCell('A10');
  titleCell.value = info.title;
  titleCell.font = { size: 20, bold: true };

  // Description
  if (info.description) {
    sheet.getCell('A11').value = info.description;
  }

  // termsOfService
  if (info.termsOfService) {
    sheet.getCell('A12').value = info.termsOfService;
  }

  // Version
  sheet.getCell('A43').value = 'Version';
  sheet.getCell('B43').value = info.version;

  // License
  if (info.license && 0 < Object.keys(info.license).length) {
    sheet.getCell('A44').value = 'License';
    sheet.getCell('B44').value = info.license.name;
    sheet.getCell('C44').value = info.license.url;
  }

  // Contact
  if (info.contact && 0 < Object.keys(info.contact).length) {
    sheet.getCell('A45').value = 'Contact';

    for (let index = 0; index < Object.keys(info.contact).length; index++) {
      const key = Object.keys(info.contact)[index];
      const value = info.contact[key];

      const cellNum = index + 45;
      sheet.getCell(`B${cellNum}`).value = key;
      sheet.getCell(`C${cellNum}`).value = value;
    }
  }
}
