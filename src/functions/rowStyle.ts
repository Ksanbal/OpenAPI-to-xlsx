import { Worksheet } from 'exceljs';

/**
 * 해당 열에 border를 생성
 *
 * @param {Worksheet} sheet - the worksheet object
 * @param {number} row - the row number
 * @param {string} from - the starting position like A
 * @param {string} to - the ending position like B
 */
export default function rowStyle(
  sheet: Worksheet,
  row: number,
  from: string,
  to: string,
  font?: any,
  style?: any,
  fill?: any,
  border?: any,
) {
  for (let code = from.charCodeAt(0); code <= to.charCodeAt(0); code++) {
    if (font !== undefined) {
      sheet.getCell(`${String.fromCharCode(code)}${row}`).font = font;
    }
    if (style !== undefined) {
      sheet.getCell(`${String.fromCharCode(code)}${row}`).style = style;
    }
    if (fill !== undefined) {
      sheet.getCell(`${String.fromCharCode(code)}${row}`).fill = fill;
    }
    if (border !== undefined) {
      sheet.getCell(`${String.fromCharCode(code)}${row}`).border = border;
    }
  }
}
