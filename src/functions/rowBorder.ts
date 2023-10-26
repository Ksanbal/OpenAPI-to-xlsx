import { Worksheet } from 'exceljs';

/**
 * 해당 열에 border를 생성
 *
 * @param {Worksheet} sheet - the worksheet object
 * @param {number} row - the row number
 * @param {string} from - the starting position like A
 * @param {string} to - the ending position like B
 */
export default function rowBorder(
  sheet: Worksheet,
  row: number,
  from: string,
  to: string,
) {
  for (let code = from.charCodeAt(0); code <= to.charCodeAt(0); code++) {
    sheet.getCell(`${String.fromCharCode(code)}${row}`).border = {
      top: { style: 'thin' },
      left: { style: 'thin' },
      bottom: { style: 'thin' },
      right: { style: 'thin' },
    };
  }
}
