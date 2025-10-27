#!/usr/bin/env node
/**
 * Flask 재기동 스크립트
 * 전문가 리뷰: "manifest 검증 → Flask 서비스 재기동 순서를 자동화"
 */

import { execSync, spawn } from 'child_process';
import { resolve, dirname } from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

// 프로젝트 루트 경로
const PROJECT_ROOT = resolve(__dirname, '../../');

/**
 * Flask 서버 재기동
 */
async function deployFlask() {
  console.log('[Deploy] Flask 재기동 시작...');

  try {
    // 1. 기존 Flask 프로세스 확인 및 종료
    console.log('🔍 [INFO] 기존 Flask 프로세스 확인 중...');

    try {
      // Windows에서 Flask 프로세스 찾기
      const processes = execSync('tasklist /FI "IMAGENAME eq python.exe" /FO CSV', {
        encoding: 'utf8',
        cwd: PROJECT_ROOT
      });

      if (processes.includes('python.exe')) {
        console.log('🛑 [INFO] 기존 Flask 프로세스 발견, 종료 중...');
        // 포트 5000을 사용하는 프로세스 종료
        try {
          execSync('netstat -ano | findstr ":5000" | findstr "LISTENING"', { encoding: 'utf8' });
          execSync('for /f "tokens=5" %a in (\'netstat -ano ^| findstr ":5000" ^| findstr "LISTENING"\') do taskkill /PID %a /F', {
            encoding: 'utf8',
            shell: true
          });
          console.log('✅ [OK] 기존 프로세스 종료 완료');
        } catch (error) {
          console.log('ℹ️ [INFO] 포트 5000을 사용하는 프로세스가 없거나 이미 종료됨');
        }
      }
    } catch (error) {
      console.log('ℹ️ [INFO] 기존 프로세스 확인 완료 (실행 중인 프로세스 없음)');
    }

    // 2. 잠시 대기 (프로세스 완전 종료 대기)
    console.log('⏳ [INFO] 프로세스 정리 대기 중...');
    await new Promise(resolve => setTimeout(resolve, 3000));

    // 3. Flask 서버 재시작
    console.log('🚀 [INFO] Flask 서버 재시작 중...');

    // start_servers.py를 백그라운드에서 실행
    const flaskProcess = spawn('python', ['start_servers.py', '--env', 'dev', '--port', '5000'], {
      cwd: PROJECT_ROOT,
      detached: true,
      stdio: ['ignore', 'pipe', 'pipe']
    });

    // 서버 시작 확인
    let serverStarted = false;
    let startupOutput = '';

    flaskProcess.stdout.on('data', (data) => {
      const output = data.toString();
      startupOutput += output;
      console.log('📊 [Flask]', output.trim());

      if (output.includes('Running on') || output.includes('* Serving Flask app')) {
        serverStarted = true;
      }
    });

    flaskProcess.stderr.on('data', (data) => {
      const output = data.toString();
      console.log('⚠️ [Flask Error]', output.trim());
    });

    // 최대 30초 대기
    await new Promise((resolve, reject) => {
      const timeout = setTimeout(() => {
        if (!serverStarted) {
          flaskProcess.kill();
          reject(new Error('Flask 서버 시작 시간 초과 (30초)'));
        }
      }, 30000);

      flaskProcess.on('error', (error) => {
        clearTimeout(timeout);
        reject(error);
      });

      const checkInterval = setInterval(() => {
        if (serverStarted) {
          clearTimeout(timeout);
          clearInterval(checkInterval);
          resolve();
        }
      }, 1000);
    });

    // 4. 서버 상태 확인
    console.log('🔍 [INFO] 서버 상태 확인 중...');

    await new Promise(resolve => setTimeout(resolve, 5000)); // 5초 추가 대기

    try {
      // curl 대신 fetch 사용하여 서버 응답 확인
      const healthCheck = await fetch('http://localhost:5000/api/health', {
        method: 'GET',
        timeout: 10000
      }).catch(() => null);

      if (healthCheck && healthCheck.ok) {
        console.log('✅ [SUCCESS] Flask 서버 정상 동작 확인!');
      } else {
        console.log('⚠️ [WARNING] 서버 헬스체크 실패, 하지만 배포는 계속 진행됩니다.');
      }
    } catch (error) {
      console.log('⚠️ [WARNING] 서버 상태 확인 실패:', error.message);
      console.log('   서버가 시작되었지만 헬스체크에 실패했을 수 있습니다.');
    }

    // 5. 프로세스를 백그라운드로 전환
    flaskProcess.unref();

    console.log('\n🎉 [SUCCESS] Flask 재기동 완료!');
    console.log('📡 [INFO] 서버 주소: http://localhost:5000');
    console.log('📋 [INFO] 프로젝트 페이지: http://localhost:5000/projects');

  } catch (error) {
    console.error('❌ [ERROR] Flask 재기동 실패:', error.message);
    console.error('🔧 [DEBUG] 수동으로 서버를 시작하세요: python start_servers.py --env dev --port 5000');
    process.exit(1);
  }
}

// fetch 폴리필 (Node.js 18 미만 버전 대응)
if (!globalThis.fetch) {
  import('node-fetch').then(({ default: fetch }) => {
    globalThis.fetch = fetch;
  }).catch(() => {
    console.log('ℹ️ [INFO] fetch API 사용 불가, 서버 헬스체크 생략');
  });
}

// 스크립트 실행
if (import.meta.url === `file://${process.argv[1]}`) {
  deployFlask();
}

export default deployFlask;