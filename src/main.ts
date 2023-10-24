/**
 * OpenAPI to XLSX
 */
// load openapi json
import bokuk from '../.cache/test.json';

// OpenAPI Info 가져오기
const info = bokuk.info;
console.log('info', info);

// 태그별 paths 묶음
const pathsByTag: Record<string, any[]> = {
  '': [],
};

// paths를 순회하면서 method를 검색
for (const [path, detail] of Object.entries(bokuk.paths)) {
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
      pathsByTag[''].push(apiDetail);
    }
  }
}

console.log(pathsByTag);
