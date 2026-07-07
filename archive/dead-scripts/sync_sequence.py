"""
프로젝트 코드 시퀀스를 스프레드시트 최대값과 동기화

이 스크립트는 SQLite 시퀀스를 현재 스프레드시트의 최대 프로젝트 번호와 일치시킵니다.
"""

import logging
import os
from pathlib import Path
from dotenv import load_dotenv
from dashboard.services.project_service import load_data, _extract_number
from dashboard.utils.user_database import get_user_database

# 환경 변수 로드
env_path = Path(__file__).parent / '.env'
load_dotenv(env_path)

logging.basicConfig(level=logging.INFO, format='%(asctime)s [%(levelname)s] %(message)s')
logger = logging.getLogger(__name__)

def sync_sequence():
    """시퀀스를 스프레드시트 최대값과 동기화"""

    logger.info("=" * 60)
    logger.info("프로젝트 코드 시퀀스 동기화 시작")
    logger.info("=" * 60)

    # 1. 스프레드시트에서 최대 번호 추출
    logger.info("스프레드시트 데이터 로드 중...")
    df = load_data(force_refresh=True)

    if df is None or df.empty:
        logger.error("스프레드시트 데이터를 불러올 수 없습니다.")
        return False

    max_number = 0
    if '프로젝트 코드' in df.columns:
        numbers = []
        for code in df['프로젝트 코드'].astype(str):
            number = _extract_number(code)
            if number is not None:
                numbers.append(number)
        max_number = max(numbers) if numbers else 0

    logger.info(f"스프레드시트 최대 프로젝트 번호: {max_number}")

    # 2. 현재 시퀀스 확인
    user_db = get_user_database()

    import sqlite3
    conn = sqlite3.connect(user_db.db_path)
    cursor = conn.execute('SELECT current_number FROM project_code_sequences WHERE pattern = ?', ('GLOBAL',))
    row = cursor.fetchone()
    current_sequence = row[0] if row else None

    logger.info(f"현재 시퀀스 번호: {current_sequence}")

    # 3. 시퀀스 업데이트
    if current_sequence is None:
        logger.info(f"시퀀스가 없습니다. {max_number}로 새로 생성합니다.")
        conn.execute('''
            INSERT INTO project_code_sequences (pattern, current_number, last_updated)
            VALUES ('GLOBAL', ?, CURRENT_TIMESTAMP)
        ''', (max_number,))
    else:
        if current_sequence < max_number:
            logger.info(f"시퀀스를 {current_sequence} → {max_number}로 업데이트합니다.")
            conn.execute('''
                UPDATE project_code_sequences
                SET current_number = ?, last_updated = CURRENT_TIMESTAMP
                WHERE pattern = 'GLOBAL'
            ''', (max_number,))
        elif current_sequence == max_number:
            logger.info(f"시퀀스가 이미 최신 상태입니다 (={max_number})")
        else:
            logger.warning(f"시퀀스({current_sequence})가 스프레드시트 최대값({max_number})보다 큽니다.")
            logger.warning(f"강제로 {current_sequence} → {max_number}로 업데이트합니다.")
            conn.execute('''
                UPDATE project_code_sequences
                SET current_number = ?, last_updated = CURRENT_TIMESTAMP
                WHERE pattern = 'GLOBAL'
            ''', (max_number,))

    conn.commit()

    # 4. 동기화 결과 확인
    cursor = conn.execute('SELECT current_number, last_updated FROM project_code_sequences WHERE pattern = ?', ('GLOBAL',))
    row = cursor.fetchone()
    updated_number, last_updated = row

    conn.close()

    logger.info("=" * 60)
    logger.info("동기화 완료!")
    logger.info(f"  - 시퀀스 번호: {updated_number}")
    logger.info(f"  - 마지막 업데이트: {last_updated}")
    logger.info(f"  - 다음 프로젝트 번호: {updated_number + 1}")
    logger.info("=" * 60)

    return True

if __name__ == '__main__':
    try:
        success = sync_sequence()
        if success:
            print("\n[SUCCESS] 시퀀스 동기화가 성공적으로 완료되었습니다.")
        else:
            print("\n[ERROR] 시퀀스 동기화에 실패했습니다.")
    except Exception as e:
        logger.error(f"오류 발생: {e}", exc_info=True)
        print(f"\n[ERROR] 오류 발생: {e}")
