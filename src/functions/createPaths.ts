import mainColor from '../config/mainColor';
import { Workbook } from 'exceljs';
import rowBorder from './rowBorder';
import rowStyle from './rowStyle';
import statusMessage from '../config/statusMessage';
import createRecursiveProperties from './createRecursiveProperties';

export default function createPaths(
  wb: Workbook,
  pathsByTag: any,
  schemas: any,
) {
  // Tag별
  for (const [key, value] of Object.entries(pathsByTag)) {
    const tag = key;
    const paths = value as any[];

    // Path별
    for (let index = 0; index < paths.length; index++) {
      const path = paths[index];
      const sheet = wb.addWorksheet(`${tag}-${index}`, {
        pageSetup: {
          paperSize: 9,
          fitToPage: true,
        },
      });

      /**
       * Column width 설정
       */
      sheet.getColumn('A').width = 20;
      sheet.getColumn('B').width = 12;
      sheet.getColumn('C').width = 12;
      sheet.getColumn('D').width = 16;
      sheet.getColumn('E').width = 40;
      sheet.getColumn('F').width = 12;
      sheet.getColumn('G').width = 12;

      /**
       * 요약 정보
       */
      // Tag
      const tagCell = sheet.getCell('A1');
      tagCell.value = 'Tag';
      tagCell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: mainColor },
      };
      sheet.mergeCells(`B1`, `G1`);
      sheet.getCell('B1').value = tag;
      rowBorder(sheet, 1, 'A', 'G');

      // Summary
      const summaryCell = sheet.getCell('A2');
      summaryCell.value = 'Summary';
      summaryCell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: mainColor },
      };
      sheet.mergeCells(`B2`, `G2`);
      sheet.getCell('B2').value = path.summary ?? '';
      rowBorder(sheet, 2, 'A', 'G');

      // Description
      const descriptionCell = sheet.getCell('A3');
      descriptionCell.value = 'Description';
      descriptionCell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: mainColor },
      };
      sheet.mergeCells(`B3`, `G3`);
      sheet.getCell('B3').value = path.description ?? '';
      rowBorder(sheet, 3, 'A', 'G');

      // Path
      const pathCell = sheet.getCell('A4');
      pathCell.value = 'Path';
      pathCell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: mainColor },
      };
      sheet.mergeCells(`B4`, `G4`);
      sheet.getCell('B4').value = path.path;
      rowBorder(sheet, 4, 'A', 'G');

      // Method
      const methodCell = sheet.getCell('A5');
      methodCell.value = 'Method';
      methodCell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: mainColor },
      };
      sheet.mergeCells(`B5`, `G5`);
      sheet.getCell('B5').value = path.method.toUpperCase();
      rowBorder(sheet, 5, 'A', 'G');

      /**
       * Request
       */
      const reqCell = sheet.getCell('A8');
      reqCell.value = 'Request';
      reqCell.font = { bold: true, size: 15 };

      let workRow = 10;

      // Header
      sheet.mergeCells(`A${workRow}`, `G${workRow}`);
      const reqHeaderCell = sheet.getCell(`A${workRow}`);
      reqHeaderCell.value = 'Header';
      reqHeaderCell.font = { bold: true, size: 12 };
      rowStyle(
        sheet,
        workRow,
        'A',
        'G',
        undefined,
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

      workRow++;
      sheet.mergeCells(`C${workRow}`, `G${workRow}`);
      sheet.getCell(`A${workRow}`).value = 'key';
      sheet.getCell(`A${workRow}`).alignment = { horizontal: 'center' };
      sheet.getCell(`B${workRow}`).value = 'value';
      sheet.getCell(`B${workRow}`).alignment = { horizontal: 'center' };
      sheet.getCell(`C${workRow}`).value = 'description';
      sheet.getCell(`C${workRow}`).alignment = { horizontal: 'center' };
      // rowBorder(sheet, workRow, 'A', 'G');
      rowStyle(
        sheet,
        workRow,
        'A',
        'G',
        undefined,
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

      // Security에 따른 header 추가
      if (path.security) {
        for (const security of path.security) {
          const isBeaer = Object.keys(security).includes('bearer');

          workRow++;
          sheet.getCell(`A${workRow}`).value = 'Authorization';

          if (isBeaer) {
            sheet.getCell(`B${workRow}`).value = 'Beaer {Token}';
          }

          sheet.mergeCells(`C${workRow}`, `G${workRow}`);
          rowBorder(sheet, workRow, 'A', 'G');
        }
      }

      // Request body에 따른 header 추가
      if (path.requestBody) {
        workRow++;
        sheet.getCell(`A${workRow}`).value = 'Content-Type';
        sheet.getCell(`B${workRow}`).value = Object.keys(
          path.requestBody.content,
        )[0];

        sheet.mergeCells(`C${workRow}`, `G${workRow}`);
        rowBorder(sheet, workRow, 'A', 'G');
      }

      if (!path.security && !path.requestBody) {
        workRow++;
        sheet.mergeCells(`C${workRow}`, `G${workRow}`);
        rowBorder(sheet, workRow, 'A', 'G');
      }

      // Parameters
      workRow++;
      workRow++;
      sheet.mergeCells(`A${workRow}`, `G${workRow}`);
      const paramsCell = sheet.getCell(`A${workRow}`);
      paramsCell.value = 'Params';
      paramsCell.font = { bold: true, size: 12 };
      rowStyle(
        sheet,
        workRow,
        'A',
        'G',
        { bold: true, size: 12 },
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

      workRow++;
      sheet.mergeCells(`D${workRow}`, `E${workRow}`);
      sheet.getCell(`A${workRow}`).value = 'in';
      sheet.getCell(`A${workRow}`).alignment = { horizontal: 'center' };
      sheet.getCell(`B${workRow}`).value = 'key';
      sheet.getCell(`B${workRow}`).alignment = { horizontal: 'center' };
      sheet.getCell(`C${workRow}`).value = 'type';
      sheet.getCell(`C${workRow}`).alignment = { horizontal: 'center' };
      sheet.getCell(`D${workRow}`).value = 'description';
      sheet.getCell(`D${workRow}`).alignment = { horizontal: 'center' };
      sheet.getCell(`F${workRow}`).value = 'default';
      sheet.getCell(`F${workRow}`).alignment = { horizontal: 'center' };
      sheet.getCell(`G${workRow}`).value = 'required';
      sheet.getCell(`G${workRow}`).alignment = { horizontal: 'center' };
      rowStyle(
        sheet,
        workRow,
        'A',
        'G',
        undefined,
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

      if (0 < path.parameters.length) {
        for (const param of path.parameters) {
          workRow++;
          sheet.mergeCells(`D${workRow}`, `E${workRow}`);
          sheet.getCell(`A${workRow}`).value = param.in;
          sheet.getCell(`B${workRow}`).value = param.name;
          sheet.getCell(`C${workRow}`).value = param.schema.type;
          sheet.getCell(`D${workRow}`).value = param.description;
          sheet.getCell(`F${workRow}`).value = param.default;
          sheet.getCell(`G${workRow}`).value = param.required ? 'O' : 'X';
          sheet.getCell(`G${workRow}`).alignment = { horizontal: 'center' };
          rowBorder(sheet, workRow, 'A', 'G');
        }
      } else {
        workRow++;
        sheet.mergeCells(`D${workRow}`, `E${workRow}`);
        rowBorder(sheet, workRow, 'A', 'G');
      }

      // Body
      workRow++;
      workRow++;
      sheet.mergeCells(`A${workRow}`, `G${workRow}`);
      const bodyCell = sheet.getCell(`A${workRow}`);
      bodyCell.value = 'Body';
      bodyCell.font = { bold: true, size: 12 };
      rowStyle(
        sheet,
        workRow,
        'A',
        'G',
        { bold: true, size: 12 },
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

      workRow++;
      sheet.mergeCells(`C${workRow}`, `F${workRow}`);
      sheet.getCell(`A${workRow}`).value = 'key';
      sheet.getCell(`A${workRow}`).alignment = { horizontal: 'center' };
      sheet.getCell(`B${workRow}`).value = 'type';
      sheet.getCell(`B${workRow}`).alignment = { horizontal: 'center' };
      sheet.getCell(`C${workRow}`).value = 'description';
      sheet.getCell(`C${workRow}`).alignment = { horizontal: 'center' };
      // sheet.getCell(`F${workRow}`).value = 'default';
      // sheet.getCell(`F${workRow}`).alignment = { horizontal: 'center' };
      sheet.getCell(`G${workRow}`).value = 'required';
      sheet.getCell(`G${workRow}`).alignment = { horizontal: 'center' };
      rowStyle(
        sheet,
        workRow,
        'A',
        'G',
        undefined,
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

      if (path.requestBody) {
        const contentType = Object.keys(path.requestBody.content)[0];

        if (contentType === 'application/json') {
          // 스키마 불러오기
          // eslint-disable-next-line @typescript-eslint/ban-ts-comment
          //@ts-ignore
          const example = Object.values(path.requestBody.content)[0].schema
            .example;

          if (example) {
            workRow++;
            sheet.mergeCells(`A${workRow}`, `G${workRow}`);
            sheet.getCell(`A${workRow}`).value = 'Example';
            rowStyle(
              sheet,
              workRow,
              'A',
              'G',
              undefined,
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
            workRow++;

            sheet.mergeCells(`A${workRow}`, `G${workRow}`);
            sheet.getCell(`A${workRow}`).value = JSON.stringify(
              example,
              null,
              2,
            );
            rowBorder(sheet, workRow, 'A', 'G');
          } else {
            // eslint-disable-next-line @typescript-eslint/ban-ts-comment
            //@ts-ignore
            const ref = Object.values(path.requestBody.content)[0]
              .schema.$ref.split('/')
              .pop();
            const schema = schemas[ref];
            for (const [key, value] of Object.entries(schema.properties)) {
              const anyValue = value as any;

              workRow++;
              sheet.mergeCells(`C${workRow}`, `F${workRow}`);
              // key
              sheet.getCell(`A${workRow}`).value = key;

              // type
              if (anyValue.type === 'array') {
                sheet.getCell(`B${workRow}`).value = `${anyValue.items.type}[]`;
              } else {
                sheet.getCell(`B${workRow}`).value = anyValue.type;
              }

              // description
              sheet.getCell(`C${workRow}`).value = anyValue.description ?? '';

              // requried
              sheet.getCell(`G${workRow}`).value = schema.required?.includes(
                key,
              )
                ? 'O'
                : 'X';
              sheet.getCell(`G${workRow}`).alignment = { horizontal: 'center' };

              rowBorder(sheet, workRow, 'A', 'G');
            }
          }
        } else if (contentType === 'multipart/form-data') {
          const schema = path.requestBody.content['multipart/form-data'].schema;

          for (const [key, value] of Object.entries(schema.properties)) {
            const anyValue = value as any;

            workRow++;
            sheet.mergeCells(`C${workRow}`, `F${workRow}`);
            // key
            sheet.getCell(`A${workRow}`).value = key;

            // type
            if (anyValue.type === 'array') {
              sheet.getCell(`B${workRow}`).value = `${anyValue.items.type}[]`;
            } else {
              sheet.getCell(`B${workRow}`).value = anyValue.type;
            }

            // description
            sheet.getCell(`C${workRow}`).value = anyValue.description ?? '';

            // requried
            sheet.getCell(`G${workRow}`).value = schema.required?.includes(key)
              ? 'O'
              : 'X';
            sheet.getCell(`G${workRow}`).alignment = { horizontal: 'center' };

            rowBorder(sheet, workRow, 'A', 'G');
          }
        }
      } else {
        workRow++;
        sheet.mergeCells(`C${workRow}`, `F${workRow}`);
        rowBorder(sheet, workRow, 'A', 'G');
      }

      /**
       * Response
       */
      workRow++;
      workRow++;
      workRow++;
      const resCell = sheet.getCell(`A${workRow}`);
      resCell.value = 'Response';
      resCell.font = { bold: true, size: 15 };

      workRow++;
      for (const [code, value] of Object.entries(path.responses)) {
        const detail = value as any;

        workRow++;
        // Status Code
        sheet.mergeCells(`A${workRow}`, `G${workRow}`);
        const codeCell = sheet.getCell(`A${workRow}`);
        // eslint-disable-next-line @typescript-eslint/ban-ts-comment
        // @ts-ignore
        codeCell.value = `${code} ${statusMessage[code]} ${
          detail.description !== '' ? `- ${detail.description}` : ''
        }`;
        rowStyle(
          sheet,
          workRow,
          'A',
          'G',
          { bold: true, size: 12 },
          undefined,
          {
            type: 'pattern',
            pattern: 'solid',
            fgColor: {
              argb:
                200 <= Number(code) && Number(code) < 300
                  ? '00dce9d5'
                  : 300 <= Number(code) && Number(code) < 400
                  ? '00b8cdf7'
                  : '00eecdcd',
            },
          },
          {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' },
          },
        );

        // Body
        workRow++;
        sheet.mergeCells(`A${workRow}`, `G${workRow}`);
        const bodyCell = sheet.getCell(`A${workRow}`);
        bodyCell.value = 'Body';
        bodyCell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: {
            argb: mainColor,
          },
        };
        bodyCell.font = { bold: true, size: 12 };
        bodyCell.alignment = { horizontal: 'center' };
        rowBorder(sheet, workRow, 'A', 'G');

        workRow++;
        if (detail.content) {
          const schema = (Object.values(detail.content)[0] as any).schema;

          // Head 부분
          if (schema.example) {
            sheet.mergeCells(`A${workRow}`, `G${workRow}`);
            sheet.getCell(`A${workRow}`).value = 'Example';
            rowStyle(
              sheet,
              workRow,
              'A',
              'G',
              { bold: true, size: 12 },
              {
                alignment: {
                  horizontal: 'center',
                },
              },
              {
                type: 'pattern',
                pattern: 'solid',
                fgColor: {
                  argb: mainColor,
                },
              },
              {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' },
              },
            );
          } else {
            // Key
            sheet.getCell(`A${workRow}`).value = 'Key1';
            sheet.getCell(`B${workRow}`).value = 'Key2';
            sheet.getCell(`C${workRow}`).value = 'Key3';
            sheet.getCell(`D${workRow}`).value = 'type';
            sheet.mergeCells(`E${workRow}`, `G${workRow}`);
            sheet.getCell(`E${workRow}`).value = 'description';
            rowStyle(
              sheet,
              workRow,
              'A',
              'G',
              undefined,
              {
                alignment: {
                  horizontal: 'center',
                },
              },
              {
                type: 'pattern',
                pattern: 'solid',
                fgColor: {
                  argb: mainColor,
                },
              },
              {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' },
              },
            );
          }
          workRow++;

          if (schema.type === 'array') {
            // type
            if (schema.items.type) {
              // 단일 type의 배열
              sheet.mergeCells(`E${workRow}`, `G${workRow}`);
              sheet.getCell(`D${workRow}`).value = `${schema.items.type}[]`;
              rowBorder(sheet, workRow, 'A', 'G');
            } else {
              // ref
            }
          } else {
            // example
            if (schema.example) {
              const example = schema.example;
              sheet.mergeCells(`A${workRow}`, `G${workRow}`);
              sheet.getCell(`A${workRow}`).value = JSON.stringify(
                example,
                null,
                2,
              );
              rowBorder(sheet, workRow, 'A', 'G');
              workRow++;
            } else if (schema['$ref']) {
              // ref
              workRow = createRecursiveProperties(
                sheet,
                workRow,
                1,
                schema,
                schemas,
              );
            }
          }
        } else {
          sheet.mergeCells(`A${workRow}`, `G${workRow}`);
          rowBorder(sheet, workRow, 'A', 'G');
        }
      }
    }
  }
}
