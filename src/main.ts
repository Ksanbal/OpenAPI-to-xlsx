/**
 * OpenAPI to XLSX
 */
import ExcelJS from 'exceljs';

// load openapi json
import spec from '../.cache/example.json';
import createCover from './functions/createCover';
import createIndex from './functions/createIndex';

// OpenAPI Info 가져오기
const info = spec.info;

// 태그별 paths 묶음
const pathsByTag: Record<string, any[]> = {
  default: [],
};

// paths를 순회하면서 method를 검색
for (const [path, detail] of Object.entries(spec.paths)) {
  // method별로 탐색 후 tag가 있으면 해당 tag에 추가
  for (const [method, api] of Object.entries(detail)) {
    const apiDetail = { path, method, ...api };
    if (api.tags) {
      for (const tag of api.tags) {
        if (pathsByTag[tag] !== undefined) {
          pathsByTag[tag].push(apiDetail);
        } else {
          pathsByTag[tag] = [apiDetail];
        }
      }
    } else {
      pathsByTag['default'].push(apiDetail);
    }
  }
}

try {
  const workbook = new ExcelJS.Workbook();

  // 표지 생성
  createCover(workbook, info);

  // tag별 paths 묶음
  createIndex(
    workbook,
    spec.servers,
    spec.components.securitySchemes,
    pathsByTag,
  );

  // 엑셀을 파일로 export
  workbook.xlsx.writeFile('./output.xlsx');
} catch (error) {
  console.error(error);
}
