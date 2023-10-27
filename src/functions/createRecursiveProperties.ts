import { Worksheet } from 'exceljs';
import rowBorder from './rowBorder';

export default function createRecursiveProperties(
  sheet: Worksheet,
  workRow: number,
  depth: number,
  schema: any,
  schemas: any,
): number {
  let ref;
  if (schema.$ref) {
    ref = schema.$ref.split('/').pop();
  } else if (schema.allOf) {
    ref = schema.allOf[0].$ref.split('/').pop();
  }
  const sc = schemas[ref];

  for (const [key, value] of Object.entries(sc.properties)) {
    const anyValue = value as any;

    // Key
    if (depth === 1) {
      sheet.getCell(`A${workRow}`).value = key;
    } else if (depth === 2) {
      sheet.getCell(`B${workRow}`).value = key;
    } else {
      sheet.getCell(`C${workRow}`).value = key;
    }
    sheet.mergeCells(`E${workRow}`, `G${workRow}`);
    rowBorder(sheet, workRow, 'A', 'G');

    if (anyValue.type) {
      sheet.getCell(`D${workRow}`).value = `${anyValue.type}${
        anyValue.format ? `(${anyValue.format})` : ''
      }`;

      if (anyValue.type === 'array') {
        if (anyValue.items.type) {
          sheet.getCell(`D${workRow}`).value = `${anyValue.items.type}[]${
            anyValue.format ? `(${anyValue.format})` : ''
          }`;
        } else {
          workRow++;
          workRow = createRecursiveProperties(
            sheet,
            workRow,
            depth + 1,
            anyValue.items,
            schemas,
          );
        }
      }

      sheet.getCell(`E${workRow}`).value = `${
        anyValue.description ? anyValue.description : ''
      } ${anyValue.example ? `ex) ${anyValue.example}` : ''}`;

      workRow++;
    } else {
      // $ref인경우
      workRow++;
      workRow = createRecursiveProperties(
        sheet,
        workRow,
        depth + 1,
        anyValue,
        schemas,
      );
    }
  }
  return workRow;
}
