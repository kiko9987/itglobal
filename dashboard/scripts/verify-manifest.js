#!/usr/bin/env node
/**
 * Manifest 검증 스크립트
 * 전문가 리뷰: "npm run build → manifest 검증 → Flask 서비스 재기동 순서를 자동화"
 */

import { readFileSync, existsSync } from 'fs';
import { resolve, dirname } from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

// 경로 설정 (Vite 5.x에서는 .vite 폴더에 생성됨)
const MANIFEST_PATH = resolve(__dirname, '../static/dist/.vite/manifest.json');
const DIST_DIR = resolve(__dirname, '../static/dist');

/**
 * Manifest 파일 검증
 */
function verifyManifest() {
  console.log('[Verify] Manifest 검증 시작...');

  // 1. Manifest 파일 존재 확인
  if (!existsSync(MANIFEST_PATH)) {
    console.error('❌ [ERROR] manifest.json 파일이 존재하지 않습니다:', MANIFEST_PATH);
    process.exit(1);
  }

  try {
    // 2. Manifest 파일 파싱
    const manifestContent = readFileSync(MANIFEST_PATH, 'utf8');
    const manifest = JSON.parse(manifestContent);

    console.log('✅ [OK] manifest.json 파싱 성공');

    // 3. 필수 엔트리 포인트 확인 (Vite 5.x 구조)
    const requiredEntries = [
      'js/pages/project-list.js',
      'css/main.css'
    ];

    const missingEntries = requiredEntries.filter(entry => !manifest[entry]);

    if (missingEntries.length > 0) {
      console.error('❌ [ERROR] 필수 엔트리 포인트가 누락되었습니다:', missingEntries);
      process.exit(1);
    }

    console.log('✅ [OK] 필수 엔트리 포인트 확인 완료');

    // 4. 생성된 파일들 존재 확인
    const manifestEntries = Object.values(manifest);
    const missingFiles = [];

    for (const entry of manifestEntries) {
      if (entry.file) {
        const filePath = resolve(DIST_DIR, entry.file);
        if (!existsSync(filePath)) {
          missingFiles.push(entry.file);
        }
      }
    }

    if (missingFiles.length > 0) {
      console.error('❌ [ERROR] Manifest에 명시된 파일들이 실제로 존재하지 않습니다:', missingFiles);
      process.exit(1);
    }

    console.log('✅ [OK] 빌드된 파일들 존재 확인 완료');

    // 5. 파일 크기 확인 (너무 작으면 빌드 실패일 가능성)
    const projectListEntry = manifest['js/pages/project-list.js'];
    const mainCssEntry = manifest['css/main.css'];

    if (projectListEntry && projectListEntry.file) {
      const jsFilePath = resolve(DIST_DIR, projectListEntry.file);
      const jsStats = readFileSync(jsFilePath, 'utf8');
      if (jsStats.length < 1000) { // 1KB 미만이면 의심스러움
        console.warn('⚠️ [WARNING] project-list.js 파일 크기가 너무 작습니다:', jsStats.length, 'bytes');
      }
    }

    if (mainCssEntry && mainCssEntry.file) {
      const cssFilePath = resolve(DIST_DIR, mainCssEntry.file);
      const cssStats = readFileSync(cssFilePath, 'utf8');
      if (cssStats.length < 500) { // 500bytes 미만이면 의심스러움
        console.warn('⚠️ [WARNING] main.css 파일 크기가 너무 작습니다:', cssStats.length, 'bytes');
      }
    }

    // 6. Manifest 내용 출력 (디버깅용)
    console.log('\n📋 [INFO] Manifest 내용:');
    Object.entries(manifest).forEach(([key, value]) => {
      console.log(`  ${key} -> ${value.file}`);
    });

    console.log('\n🎉 [SUCCESS] Manifest 검증 완료!');

    // 7. 환경 변수로 검증 결과 전달
    process.env.MANIFEST_VERIFIED = 'true';

  } catch (error) {
    console.error('❌ [ERROR] Manifest 검증 실패:', error.message);
    process.exit(1);
  }
}

// 스크립트 실행
if (import.meta.url.startsWith('file://') && process.argv[1] && import.meta.url.includes('verify-manifest.js')) {
  verifyManifest();
}

export default verifyManifest;