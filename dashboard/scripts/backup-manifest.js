#!/usr/bin/env node
/**
 * Manifest 백업 스크립트
 * 전문가 리뷰: "실패 시 레거시 CSS 로드로 롤백되는 안전장치를 마련하세요"
 */

import { copyFileSync, existsSync, mkdirSync, readFileSync, writeFileSync } from 'fs';
import { resolve, dirname } from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

// 경로 설정
const MANIFEST_PATH = resolve(__dirname, '../static/dist/manifest.json');
const BACKUP_DIR = resolve(__dirname, '../backups');
const BACKUP_MANIFEST_PATH = resolve(BACKUP_DIR, 'manifest.backup.json');
const BACKUP_DIST_DIR = resolve(BACKUP_DIR, 'dist');

/**
 * Manifest 및 빌드 파일 백업
 */
function backupManifest() {
  console.log('[Backup] Manifest 백업 시작...');

  try {
    // 1. 백업 디렉토리 생성
    if (!existsSync(BACKUP_DIR)) {
      mkdirSync(BACKUP_DIR, { recursive: true });
      console.log('📁 [INFO] 백업 디렉토리 생성:', BACKUP_DIR);
    }

    if (!existsSync(BACKUP_DIST_DIR)) {
      mkdirSync(BACKUP_DIST_DIR, { recursive: true });
      console.log('📁 [INFO] 백업 dist 디렉토리 생성:', BACKUP_DIST_DIR);
    }

    // 2. 기존 manifest.json 백업
    if (existsSync(MANIFEST_PATH)) {
      copyFileSync(MANIFEST_PATH, BACKUP_MANIFEST_PATH);
      console.log('✅ [OK] Manifest 파일 백업 완료');

      // 3. 기존 빌드 파일들 백업
      const manifest = JSON.parse(readFileSync(MANIFEST_PATH, 'utf8'));
      const distDir = resolve(__dirname, '../static/dist');

      let backedUpFiles = 0;

      Object.values(manifest).forEach(entry => {
        if (entry.file) {
          const sourceFile = resolve(distDir, entry.file);
          const backupFile = resolve(BACKUP_DIST_DIR, entry.file);

          if (existsSync(sourceFile)) {
            // 백업 파일의 디렉토리 생성
            const backupFileDir = dirname(backupFile);
            if (!existsSync(backupFileDir)) {
              mkdirSync(backupFileDir, { recursive: true });
            }

            copyFileSync(sourceFile, backupFile);
            backedUpFiles++;
          }
        }
      });

      console.log(`✅ [OK] 빌드 파일 ${backedUpFiles}개 백업 완료`);

      // 4. 백업 메타데이터 생성
      const backupMetadata = {
        timestamp: new Date().toISOString(),
        originalManifest: manifest,
        backedUpFiles: backedUpFiles,
        backupPath: BACKUP_DIR
      };

      writeFileSync(
        resolve(BACKUP_DIR, 'backup-metadata.json'),
        JSON.stringify(backupMetadata, null, 2)
      );

      console.log('📋 [INFO] 백업 메타데이터 생성 완료');

    } else {
      console.log('⚠️ [WARNING] 기존 manifest.json이 없습니다. 백업을 생략합니다.');

      // 빈 백업 메타데이터 생성 (첫 배포인 경우)
      const emptyBackupMetadata = {
        timestamp: new Date().toISOString(),
        originalManifest: null,
        backedUpFiles: 0,
        backupPath: BACKUP_DIR,
        isFirstDeploy: true
      };

      writeFileSync(
        resolve(BACKUP_DIR, 'backup-metadata.json'),
        JSON.stringify(emptyBackupMetadata, null, 2)
      );
    }

    console.log('\n🎉 [SUCCESS] 백업 완료!');
    console.log(`📂 [INFO] 백업 위치: ${BACKUP_DIR}`);

  } catch (error) {
    console.error('❌ [ERROR] 백업 실패:', error.message);
    process.exit(1);
  }
}

/**
 * 백업 상태 확인
 */
function checkBackupStatus() {
  const backupMetadataPath = resolve(BACKUP_DIR, 'backup-metadata.json');

  if (existsSync(backupMetadataPath)) {
    const metadata = JSON.parse(readFileSync(backupMetadataPath, 'utf8'));
    console.log('\n📊 [INFO] 백업 상태:');
    console.log(`  - 백업 시간: ${metadata.timestamp}`);
    console.log(`  - 백업된 파일 수: ${metadata.backedUpFiles}`);
    console.log(`  - 첫 배포 여부: ${metadata.isFirstDeploy || false}`);

    return metadata;
  }

  return null;
}

// 스크립트 실행
if (import.meta.url === `file://${process.argv[1]}`) {
  backupManifest();
  checkBackupStatus();
}

export { backupManifest, checkBackupStatus };