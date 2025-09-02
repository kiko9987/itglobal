import smtplib
import requests
import schedule
import time
import logging
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime, timedelta
from typing import Dict, List, Any, Optional
import os
from dotenv import load_dotenv
import json

# 환경 변수 로드
load_dotenv()

# 로깅 설정
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class NotificationSystem:
    """알림 시스템 클래스"""
    
    def __init__(self):
        """알림 시스템 초기화"""
        # 이메일 설정
        self.email_host = os.getenv('EMAIL_HOST', 'smtp.gmail.com')
        self.email_port = int(os.getenv('EMAIL_PORT', 587))
        self.email_username = os.getenv('EMAIL_USERNAME')
        self.email_password = os.getenv('EMAIL_PASSWORD')
        
        # 슬랙 설정
        self.slack_webhook = os.getenv('SLACK_WEBHOOK_URL')
        
        # 알림 설정
        self.notification_interval = int(os.getenv('NOTIFICATION_INTERVAL_HOURS', 24))
        self.missing_fields_threshold = int(os.getenv('MISSING_FIELDS_THRESHOLD', 3))
        
        # 영업사원 이메일 매핑 (환경 변수나 설정 파일에서 로드)
        self.sales_emails = self._load_sales_emails()
        
        # 관리자 이메일
        self.admin_emails = os.getenv('ADMIN_EMAILS', '').split(',') if os.getenv('ADMIN_EMAILS') else []
    
    def _load_sales_emails(self) -> Dict[str, str]:
        """영업사원 이메일 매핑 로드"""
        # 실제로는 설정 파일이나 데이터베이스에서 로드해야 함
        # 여기서는 환경 변수에서 JSON 형태로 로드
        sales_emails_json = os.getenv('SALES_EMAILS', '{}')
        try:
            return json.loads(sales_emails_json)
        except json.JSONDecodeError:
            logger.warning("영업사원 이메일 설정을 불러올 수 없습니다. 기본값 사용.")
            return {
                '양곡': 'yangkok@company.com',
                '종로': 'jongno@company.com',
                '서현': 'seohyun@company.com',
                '목동': 'mokdong@company.com',
                '일산': 'ilsan@company.com'
            }
    
    def send_email(self, to_email: str, subject: str, body: str, html_body: str = None) -> bool:
        """이메일 발송"""
        try:
            if not self.email_username or not self.email_password:
                logger.error("이메일 계정 정보가 설정되지 않았습니다.")
                return False
            
            # MIME 메시지 생성
            msg = MIMEMultipart('alternative')
            msg['From'] = self.email_username
            msg['To'] = to_email
            msg['Subject'] = subject
            
            # 텍스트 버전
            text_part = MIMEText(body, 'plain', 'utf-8')
            msg.attach(text_part)
            
            # HTML 버전 (있는 경우)
            if html_body:
                html_part = MIMEText(html_body, 'html', 'utf-8')
                msg.attach(html_part)
            
            # SMTP 서버 연결 및 발송
            server = smtplib.SMTP(self.email_host, self.email_port)
            server.starttls()
            server.login(self.email_username, self.email_password)
            server.send_message(msg)
            server.quit()
            
            logger.info(f"이메일 발송 성공: {to_email}")
            return True
            
        except Exception as e:
            logger.error(f"이메일 발송 실패: {to_email}, 오류: {str(e)}")
            return False
    
    def send_slack_notification(self, message: str, channel: str = None) -> bool:
        """슬랙 알림 발송"""
        try:
            if not self.slack_webhook:
                logger.warning("슬랙 웹훅 URL이 설정되지 않았습니다.")
                return False
            
            payload = {
                'text': message,
                'username': '냉난방기 대시보드',
                'icon_emoji': ':thermometer:'
            }
            
            if channel:
                payload['channel'] = channel
            
            response = requests.post(
                self.slack_webhook,
                json=payload,
                headers={'Content-Type': 'application/json'}
            )
            
            if response.status_code == 200:
                logger.info("슬랙 알림 발송 성공")
                return True
            else:
                logger.error(f"슬랙 알림 발송 실패: {response.status_code}")
                return False
                
        except Exception as e:
            logger.error(f"슬랙 알림 발송 오류: {str(e)}")
            return False
    
    def check_missing_data(self, missing_analysis: Dict[str, Any]) -> List[Dict[str, Any]]:
        """누락 데이터 체크 및 알림 대상 추출"""
        notifications = []
        
        try:
            person_analysis = missing_analysis.get('person_analysis', {})
            
            for person, info in person_analysis.items():
                total_missing = info.get('total_missing', 0)
                critical_missing = info.get('critical_missing', [])
                
                # 임계값 초과 시 알림 대상
                if total_missing >= self.missing_fields_threshold:
                    notification = {
                        'person': person,
                        'total_missing': total_missing,
                        'critical_fields': [
                            {
                                'field': missing['field'],
                                'count': missing['missing_count'],
                                'projects': missing.get('projects', [])[:5]  # 최대 5개 프로젝트만
                            }
                            for missing in critical_missing[:5]  # 상위 5개 필드만
                        ],
                        'email': self.sales_emails.get(person),
                        'priority': 'high' if total_missing >= self.missing_fields_threshold * 2 else 'normal'
                    }
                    notifications.append(notification)
            
            return notifications
            
        except Exception as e:
            logger.error(f"누락 데이터 체크 오류: {str(e)}")
            return []
    
    def generate_missing_data_email(self, notification: Dict[str, Any]) -> tuple:
        """누락 데이터 알림 이메일 생성"""
        person = notification['person']
        total_missing = notification['total_missing']
        critical_fields = notification['critical_fields']
        priority = notification['priority']
        
        # 제목
        priority_text = "[긴급]" if priority == 'high' else ""
        subject = f"{priority_text} {person}님 - 공사 데이터 입력 요청 ({total_missing}건 누락)"
        
        # 본문 (텍스트)
        body = f"""
안녕하세요 {person}님,

공사 관리 시스템에서 다음과 같은 데이터 입력이 누락되어 있습니다.
빠른 시일 내에 입력 부탁드립니다.

총 누락 항목: {total_missing}건

상세 내역:
"""
        
        for field_info in critical_fields:
            body += f"\n• {field_info['field']}: {field_info['count']}건 누락"
            if field_info['projects']:
                body += f"\n  - 해당 프로젝트: {', '.join(field_info['projects'][:3])}"
                if len(field_info['projects']) > 3:
                    body += f" 외 {len(field_info['projects']) - 3}건"
        
        body += f"""

데이터 입력 완료 후 회신 부탁드립니다.

감사합니다.
냉난방기 설치 관리시스템
발송시간: {datetime.now().strftime('%Y년 %m월 %d일 %H:%M')}
"""
        
        # HTML 본문
        html_body = f"""
<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <style>
        body {{ font-family: 'Malgun Gothic', Arial, sans-serif; line-height: 1.6; margin: 0; padding: 20px; }}
        .container {{ max-width: 600px; margin: 0 auto; background: #f9f9f9; padding: 30px; border-radius: 10px; }}
        .header {{ background: {'#dc3545' if priority == 'high' else '#667eea'}; color: white; padding: 20px; border-radius: 5px; text-align: center; margin-bottom: 20px; }}
        .content {{ background: white; padding: 20px; border-radius: 5px; }}
        .missing-item {{ background: #fff3cd; border: 1px solid #ffeaa7; padding: 10px; margin: 10px 0; border-radius: 5px; }}
        .projects {{ font-size: 12px; color: #666; margin-left: 20px; }}
        .footer {{ text-align: center; margin-top: 20px; font-size: 12px; color: #666; }}
        .priority-high {{ color: #dc3545; font-weight: bold; }}
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h2>{'🚨 긴급 데이터 입력 요청' if priority == 'high' else '📋 데이터 입력 요청'}</h2>
        </div>
        
        <div class="content">
            <h3>안녕하세요 {person}님,</h3>
            <p>공사 관리 시스템에서 다음과 같은 데이터 입력이 누락되어 있습니다.</p>
            <p><strong class="{'priority-high' if priority == 'high' else ''}">총 누락 항목: {total_missing}건</strong></p>
            
            <h4>상세 내역:</h4>
"""
        
        for field_info in critical_fields:
            html_body += f"""
            <div class="missing-item">
                <strong>• {field_info['field']}: {field_info['count']}건 누락</strong>
"""
            if field_info['projects']:
                projects_text = ', '.join(field_info['projects'][:3])
                if len(field_info['projects']) > 3:
                    projects_text += f" 외 {len(field_info['projects']) - 3}건"
                html_body += f'<div class="projects">해당 프로젝트: {projects_text}</div>'
            
            html_body += "</div>"
        
        html_body += f"""
            <p style="margin-top: 20px;">데이터 입력 완료 후 회신 부탁드립니다.</p>
            <p><strong>구글 시트 바로가기:</strong> <a href="https://docs.google.com/spreadsheets/d/{os.getenv('GOOGLE_SHEET_ID', '')}" target="_blank">여기를 클릭하세요</a></p>
        </div>
        
        <div class="footer">
            <p>냉난방기 설치 관리시스템<br>
            발송시간: {datetime.now().strftime('%Y년 %m월 %d일 %H:%M')}</p>
        </div>
    </div>
</body>
</html>
"""
        
        return subject, body, html_body
    
    def send_missing_data_notifications(self, missing_analysis: Dict[str, Any]) -> Dict[str, int]:
        """누락 데이터 알림 발송"""
        notifications = self.check_missing_data(missing_analysis)
        
        results = {
            'email_sent': 0,
            'email_failed': 0,
            'slack_sent': 0,
            'slack_failed': 0,
            'total_notifications': len(notifications)
        }
        
        for notification in notifications:
            person = notification['person']
            email = notification['email']
            
            # 이메일 발송
            if email:
                subject, body, html_body = self.generate_missing_data_email(notification)
                if self.send_email(email, subject, body, html_body):
                    results['email_sent'] += 1
                else:
                    results['email_failed'] += 1
            else:
                logger.warning(f"{person}님의 이메일 주소가 설정되지 않았습니다.")
                results['email_failed'] += 1
        
        # 슬랙 요약 알림
        if notifications:
            slack_message = f"""
🔔 데이터 입력 알림 발송 완료

총 {len(notifications)}명의 영업사원에게 알림 발송
• 이메일 발송 성공: {results['email_sent']}건
• 이메일 발송 실패: {results['email_failed']}건

알림 대상:
{chr(10).join([f"• {notif['person']}: {notif['total_missing']}건 누락" for notif in notifications[:5]])}
{"..." if len(notifications) > 5 else ""}

발송시간: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
"""
            if self.send_slack_notification(slack_message):
                results['slack_sent'] = 1
            else:
                results['slack_failed'] = 1
        
        return results
    
    def send_daily_summary(self, summary_stats: Dict[str, Any], outstanding_analysis: Dict[str, Any]) -> bool:
        """일일 요약 보고서 발송"""
        try:
            # 관리자용 일일 요약
            subject = f"[일일보고] 냉난방기 공사 현황 요약 - {datetime.now().strftime('%Y.%m.%d')}"
            
            body = f"""
일일 공사 현황 요약 보고서

📊 전체 현황:
• 총 프로젝트: {summary_stats.get('total_projects', 0)}건
• 완료 프로젝트: {summary_stats.get('completed_projects', 0)}건
• 진행중 프로젝트: {summary_stats.get('in_progress_projects', 0)}건
• 총 매출: {summary_stats.get('total_amount', 0):,}원
• 총 미수금: {summary_stats.get('total_outstanding', 0):,}원
• 회수율: {summary_stats.get('collection_rate', 0):.1f}%

💰 미수금 현황:
• 미수금 건수: {outstanding_analysis.get('total_cases', 0)}건
• 미수금 총액: {outstanding_analysis.get('total_amount', 0):,}원

발송시간: {datetime.now().strftime('%Y년 %m월 %d일 %H:%M')}
"""
            
            # 관리자들에게 발송
            sent_count = 0
            for admin_email in self.admin_emails:
                if admin_email.strip() and self.send_email(admin_email.strip(), subject, body):
                    sent_count += 1
            
            # 슬랙 알림
            slack_message = f"""
📈 일일 현황 요약 ({datetime.now().strftime('%Y.%m.%d')})

총 프로젝트: {summary_stats.get('total_projects', 0)}건
완료: {summary_stats.get('completed_projects', 0)}건 | 진행중: {summary_stats.get('in_progress_projects', 0)}건
총 매출: {summary_stats.get('total_amount', 0):,}원
미수금: {outstanding_analysis.get('total_amount', 0):,}원 ({outstanding_analysis.get('total_cases', 0)}건)
회수율: {summary_stats.get('collection_rate', 0):.1f}%
"""
            self.send_slack_notification(slack_message)
            
            logger.info(f"일일 요약 보고서 발송 완료: {sent_count}명")
            return sent_count > 0
            
        except Exception as e:
            logger.error(f"일일 요약 보고서 발송 오류: {str(e)}")
            return False

def run_notification_scheduler():
    """알림 스케줄러 실행"""
    from dashboard.utils.google_sheets import GoogleSheetsManager
    from dashboard.utils.data_analyzer import DataAnalyzer
    
    notification_system = NotificationSystem()
    
    def daily_check():
        """일일 체크 작업"""
        try:
            logger.info("일일 데이터 체크 시작")
            
            # 데이터 로드
            sheet_id = os.getenv('GOOGLE_SHEET_ID')
            if not sheet_id:
                logger.error("GOOGLE_SHEET_ID가 설정되지 않았습니다.")
                return
            
            manager = GoogleSheetsManager()
            df = manager.get_sheet_data(sheet_id)
            
            if df.empty:
                logger.warning("데이터가 없어 체크를 건너뜁니다.")
                return
            
            analyzer = DataAnalyzer(df)
            
            # 누락 데이터 분석
            missing_analysis = analyzer.check_missing_data()
            if missing_analysis:
                results = notification_system.send_missing_data_notifications(missing_analysis)
                logger.info(f"누락 데이터 알림 발송 결과: {results}")
            
            # 일일 요약 (관리자용)
            summary_stats = analyzer.get_summary_stats()
            outstanding_analysis = analyzer.get_outstanding_analysis()
            notification_system.send_daily_summary(summary_stats, outstanding_analysis)
            
            logger.info("일일 데이터 체크 완료")
            
        except Exception as e:
            logger.error(f"일일 체크 작업 오류: {str(e)}")
    
    # 스케줄 설정
    schedule.every().day.at("09:00").do(daily_check)  # 매일 오전 9시
    schedule.every().day.at("18:00").do(daily_check)  # 매일 오후 6시
    
    logger.info("알림 스케줄러 시작 (09:00, 18:00 실행)")
    
    while True:
        schedule.run_pending()
        time.sleep(60)  # 1분마다 체크

if __name__ == "__main__":
    run_notification_scheduler()