import os
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from dotenv import load_dotenv
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import re
import requests
import pickle
from datetime import datetime, timedelta
from selenium.webdriver.chrome.options import Options
import undetected_chromedriver as uc
import atexit
import signal
import sys

opts = Options()
opts.add_argument("--ignore-certificate-errors")
opts.add_argument("--disable-features=UseChromeRootStore")
opts.add_argument("--ignore-ssl-errors")
opts.add_argument("--allow-running-insecure-content")
opts.add_argument("--allow-insecure-localhost")
opts.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36")
# 브라우저를 보이게 설정 (로그인 확인용)
opts.add_argument("--start-maximized")

# undetected_chromedriver 종료 중 예외 출력 억제
try:
    uc.Chrome.__del__ = lambda self: None  # type: ignore[attr-defined]
except Exception:
    pass

driver = None

def _safe_driver_quit():
    try:
        if driver is not None:
            driver.quit()
    except BaseException:
        pass

atexit.register(_safe_driver_quit)

# CTRL+C 등 강제 종료 시에도 브라우저를 정상 종료
def _graceful_shutdown(signum, frame):
    try:
        if driver is not None:
            driver.quit()
    except BaseException:
        pass
    sys.exit(0)

try:
    signal.signal(signal.SIGINT, _graceful_shutdown)
except Exception:
    pass
try:
    signal.signal(signal.SIGTERM, _graceful_shutdown)
except Exception:
    pass

# 1. 환경변수(.env)에서 Stibee 로그인 정보 불러오기
load_dotenv('C:/Users/7040_64bit/Desktop/20250708.env')
STIBEE_EMAIL = os.getenv('STIBEE_EMAIL')
STIBEE_PASSWORD = os.getenv('STIBEE_PASSWORD')
STIBEE_API_KEY = os.getenv('STIBEE_API_KEY')  # API 키 추가

print("환경변수 로드 확인:")
print(f"이메일: {STIBEE_EMAIL}")
print(f"비밀번호: {'*' * len(STIBEE_PASSWORD) if STIBEE_PASSWORD else 'None'}")
print(f"API 키: {'있음' if STIBEE_API_KEY else '없음'}")

# 2. 구글 시트 인증
json_path = 'C:/Users/7040_64bit/Desktop/python/single-arcadia-436209-g2-ebfcbea8518b.json'
sheet_url = 'https://docs.google.com/spreadsheets/d/1kAO3oRGPuLSEPx_4nYuSYiUf0m3xdbqtI1hR42WNfjU/edit?gid=0'
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name(json_path, scope)
client = gspread.authorize(creds)

## Selenium 기본 드라이버 생성 코드 제거(undetected_chromedriver만 사용)

def stibee_login():
    print("스티비 로그인 중...")
    global driver
    if driver is None:
        driver = uc.Chrome(options=opts, headless=False)
    
    # 에러 페이지 우회를 위해 여러 방법 시도
    print("로그인 페이지로 이동 시도...")
    
    # 방법 1: 직접 로그인 페이지로 이동
    try:
        driver.get('https://stibee.com/login')
        time.sleep(5)
        print(f"현재 URL: {driver.current_url}")
        
        # 에러 페이지인지 확인
        if 'error' in driver.current_url:
            print("❌ 에러 페이지로 리다이렉트됨. 다른 방법 시도...")
            
            # 방법 2: 메인 페이지 먼저 방문 후 로그인 페이지로
            driver.get('https://stibee.com/')
            time.sleep(3)
            driver.get('https://stibee.com/login')
            time.sleep(5)
            print(f"방법 2 시도 후 URL: {driver.current_url}")
            
            # 여전히 에러 페이지면 방법 3 시도
            if 'error' in driver.current_url:
                print("방법 3: 다른 경로로 시도...")
                driver.get('https://stibee.com/auth/login')
                time.sleep(5)
                print(f"방법 3 시도 후 URL: {driver.current_url}")
                
                # 마지막 방법: JavaScript로 직접 이동
                if 'error' in driver.current_url:
                    print("JavaScript로 직접 이동 시도...")
                    driver.execute_script("window.location.href = 'https://stibee.com/login'")
                    time.sleep(5)
                    print(f"JavaScript 이동 후 URL: {driver.current_url}")
        
    except Exception as e:
        print(f"페이지 이동 중 오류: {e}")
        # 오류 발생 시 다시 시도
        driver.get('https://stibee.com/login')
        time.sleep(5)
    
    # 이미 로그인되어 있는지 확인
    if 'dashboard' in driver.current_url or 'emails' in driver.current_url:
        print("이미 로그인되어 있습니다.")
        return True
    
    print(f"현재 URL: {driver.current_url}")
    
    try:
        # 페이지 로딩 완료 대기
        wait = WebDriverWait(driver, 20)
        
        # 이메일 입력란 찾기 (다양한 방법 시도)
        email_input = None
        selectors = [
            (By.NAME, "userId"),
            (By.NAME, "email"),
            (By.CSS_SELECTOR, 'input[type="email"]'),
            (By.CSS_SELECTOR, 'input[type="text"]'),
            (By.XPATH, '//form//input[1]'),
            (By.XPATH, '//input[@type="text"]'),
            (By.XPATH, '//input[@type="email"]')
        ]
        
        for by_type, selector in selectors:
            try:
                email_input = wait.until(EC.presence_of_element_located((by_type, selector)))
                print(f"이메일 입력란 찾음: {by_type}={selector}")
                break
            except:
                continue
        
        if not email_input:
            print("❌ 이메일 입력란을 찾을 수 없습니다.")
            return False
        
        email_input.clear()
        email_input.send_keys(STIBEE_EMAIL)
        print("이메일 입력 완료")
        
        # 비밀번호 입력란 찾기
        password_input = None
        password_selectors = [
            (By.NAME, "password"),
            (By.CSS_SELECTOR, 'input[type="password"]'),
            (By.XPATH, '//form//input[2]'),
            (By.XPATH, '//input[@type="password"]')
        ]
        
        for by_type, selector in password_selectors:
            try:
                password_input = driver.find_element(by_type, selector)
                print(f"비밀번호 입력란 찾음: {by_type}={selector}")
                break
            except:
                continue
        
        if not password_input:
            print("❌ 비밀번호 입력란을 찾을 수 없습니다.")
            return False
            
        password_input.clear()
        password_input.send_keys(STIBEE_PASSWORD)
        print("비밀번호 입력 완료")
        
        # 로그인 버튼 찾기 및 클릭
        login_button = None
        button_selectors = [
            (By.CSS_SELECTOR, 'button[type="submit"]'),
            (By.XPATH, '//button[contains(text(), "로그인")]'),
            (By.XPATH, '//form//button'),
            (By.CSS_SELECTOR, 'input[type="submit"]'),
            (By.XPATH, '//button[@type="submit"]')
        ]
        
        for by_type, selector in button_selectors:
            try:
                login_button = driver.find_element(by_type, selector)
                print(f"로그인 버튼 찾음: {by_type}={selector}")
                break
            except:
                continue
        
        if not login_button:
            print("❌ 로그인 버튼을 찾을 수 없습니다.")
            return False
            
        login_button.click()
        print("로그인 버튼 클릭")
        
        # 로그인 완료 대기
        time.sleep(20)  # 대기 시간 더 증가
        print("로그인 후 현재 URL:", driver.current_url)
        
        # 로그인 성공 여부 확인 (더 정확한 확인)
        if 'login' in driver.current_url:
            print("❌ 로그인 실패 - 여전히 로그인 페이지에 있습니다.")
            return False
        
        # 대시보드나 이메일 페이지로 이동했는지 확인
        if 'dashboard' in driver.current_url or 'emails' in driver.current_url:
            print("✅ 로그인 성공")
            return True
        else:
            print("⚠️ 로그인 상태 불확실, 추가 확인 필요")
            # 현재 페이지 새로고침하여 상태 확인
            driver.refresh()
            time.sleep(5)
            if 'login' not in driver.current_url:
                print("✅ 로그인 성공 (새로고침 후 확인)")
                return True
            else:
                print("❌ 로그인 실패")
                return False
        
    except Exception as e:
        print(f"로그인 중 오류 발생: {e}")
        return False

def get_email_list_from_website():
    """스티비 웹사이트에서 발송완료된 이메일 목록 스크래핑 (최신순)"""
    print("발송완료 이메일 목록 페이지로 이동 중...")
    global driver
    if driver is None:
        driver = uc.Chrome(options=opts, headless=False)  # 브라우저를 보이게 설정
    
    # 먼저 로그인 확인 및 수행
    if 'login' in driver.current_url or driver.current_url == 'data:,':
        print("로그인이 필요합니다. 로그인을 수행합니다.")
        if not stibee_login():
            print("❌ 로그인 실패")
            return []
    
    # 발송완료 페이지로 직접 이동
    driver.get('https://stibee.com/emails/1?emailStatus=3')
    time.sleep(10)  # 로딩 대기 시간 증가
    
    # 현재 URL 확인
    current_url = driver.current_url
    print(f"현재 URL: {current_url}")
    
    # 로그인 페이지로 리다이렉트되었는지 확인
    if 'login' in current_url:
        print("❌ 로그인 세션이 만료되었습니다. 다시 로그인이 필요합니다.")
        return []
    
    # 디버깅용: 페이지 HTML 저장
    print("파일 저장 시도: stibee_completed_emails_page.html")
    try:
        with open('stibee_completed_emails_page.html', 'w', encoding='utf-8') as f:
            f.write(driver.page_source)
        print("발송완료 이메일 목록 페이지 HTML을 stibee_completed_emails_page.html로 저장했습니다.")
    except Exception as e:
        print(f"파일 저장 실패: {e}")
    
    try:
        # JavaScript 렌더링 완료 대기
        print("JavaScript 렌더링 완료 대기 중...")
        wait = WebDriverWait(driver, 20)
        
        # 메인 컨테이너 찾기 (여러 방법 시도)
        main_container = None
        container_selectors = [
            '/html/body/div[1]/div/div[1]/div/section/main',
            '//main',
            '//div[contains(@class, "main-content")]',
            '//div[contains(@class, "content")]',
            '//section[contains(@class, "main")]'
        ]
        
        for selector in container_selectors:
            try:
                main_container = wait.until(EC.presence_of_element_located((By.XPATH, selector)))
                print(f"메인 컨테이너 찾음: {selector}")
                break
            except:
                continue
        
        if not main_container:
            print("메인 컨테이너를 찾을 수 없습니다.")
            return []
        
        # 테이블 로딩 대기
        print("발송완료 이메일 목록 테이블 로딩 대기 중...")
        time.sleep(8)
        
        # 방법 1: 테이블 행 찾기 (더 정확한 패턴 우선)
        rows = []
        table_patterns = [
            '//table/tbody/tr',
            '//div[contains(@class, "ant-table")]//tbody//tr',
            '//div[contains(@class, "ant-table-tbody")]//tr',
            '//div[contains(@class, "ant-table-row")]',
            '//table//tr[position()>1]',
            '//div[contains(@class, "email-list")]//tr',
            '//div[contains(@class, "email-item")]',
            '//div[contains(@class, "list-item")]',
            '//div[contains(@class, "main")]//table//tr'
        ]
        
        for pattern in table_patterns:
            try:
                rows = driver.find_elements(By.XPATH, pattern)
                if rows and len(rows) > 0:
                    print(f"XPATH 패턴 '{pattern}'으로 {len(rows)}개 행 발견")
                    break
            except Exception as e:
                print(f"XPATH 패턴 '{pattern}' 실패: {e}")
                continue
        
        # 방법 1-1: 더 구체적인 이메일 관련 행 찾기
        if not rows:
            print("일반 테이블 행을 찾을 수 없어 이메일 관련 행을 찾습니다...")
            email_row_patterns = [
                '//div[contains(@class, "email-row")]',
                '//div[contains(@class, "email-item")]',
                '//div[contains(@class, "list-item")]',
                '//div[contains(@class, "row") and contains(@class, "email")]',
                '//div[contains(@class, "item") and contains(@class, "email")]'
            ]
            
            for pattern in email_row_patterns:
                try:
                    rows = driver.find_elements(By.XPATH, pattern)
                    if rows and len(rows) > 0:
                        print(f"이메일 관련 패턴 '{pattern}'으로 {len(rows)}개 행 발견")
                        break
                except:
                    continue
        
        # 방법 1-2: 실제 이메일 링크가 있는 행만 필터링
        if rows and len(rows) > 0:
            print(f"발견된 행 수: {len(rows)}")
            print("실제 이메일 링크가 있는 행만 필터링합니다...")
            
            filtered_rows = []
            for i, row in enumerate(rows):
                try:
                    # 이메일 링크가 있는지 확인
                    email_link = row.find_element(By.XPATH, './/a[contains(@href, "/email/")]')
                    if email_link:
                        filtered_rows.append(row)
                        print(f"행 {i+1}: 이메일 링크 발견")
                except:
                    continue
            
            if filtered_rows:
                rows = filtered_rows
                print(f"필터링 후 실제 이메일 행 수: {len(rows)}")
            else:
                print("실제 이메일 링크가 있는 행을 찾을 수 없습니다.")
                rows = []
        
        # 방법 2: 테이블 행을 찾을 수 없는 경우 대안 방법 사용
        if not rows:
            print("테이블 행을 찾을 수 없어 대안 방법을 시도합니다...")
            
            # 2-1. 모든 링크에서 이메일 관련 링크 찾기
            all_links = driver.find_elements(By.TAG_NAME, "a")
            email_links = []
            
            for link in all_links:
                try:
                    href = link.get_attribute('href')
                    if href and '/email/' in href:
                        email_links.append(link)
                except:
                    continue
            
            print(f"이메일 관련 링크 수: {len(email_links)}")
            
            # 2-2. 페이지 전체에서 이메일 ID 패턴 찾기
            page_text = driver.page_source
            email_ids = re.findall(r'/email/(\d+)', page_text)
            unique_email_ids = list(set(email_ids))
            
            print(f"페이지에서 발견된 고유 이메일 ID 수: {len(unique_email_ids)}")
            if unique_email_ids:
                print(f"발견된 이메일 ID들: {unique_email_ids[:10]}")
                
                recent_emails = []
                for email_id in unique_email_ids[:20]:  # 최대 20개
                    email_data = {
                        'id': email_id,
                        'subject': f'이메일 {email_id}',
                        'url': f'https://stibee.com/email/{email_id}',
                        'send_date': 'unknown'
                    }
                    recent_emails.append(email_data)
                
                print(f"이메일 데이터 생성 완료: {len(recent_emails)}개")
                return recent_emails
        
        if not rows:
            print("발견된 이메일 행 수: 0")
            print("발송완료 이메일 목록을 찾을 수 없습니다. 페이지 구조를 확인해주세요.")
            return []
        
        print(f"발견된 이메일 행 수: {len(rows)}")
        
        # 현재 날짜 기준 한달 전 계산
        current_date = datetime.now()
        one_month_ago = current_date - timedelta(days=30)
        print(f"현재 날짜: {current_date.strftime('%Y-%m-%d')}")
        print(f"한달 전 날짜: {one_month_ago.strftime('%Y-%m-%d')}")
        
        recent_emails = []
        
        for i, row in enumerate(rows[:20]):  # 최대 20개만 처리 (최신순)
            try:
                print(f"\n--- 행 {i+1} 처리 중 ---")
                
                # 디버깅: 행의 HTML 구조 확인
                try:
                    row_html = row.get_attribute('outerHTML')
                    print(f"행 {i+1} HTML: {row_html[:200]}...")
                except:
                    print(f"행 {i+1} HTML 추출 실패")
                
                # 이메일 링크와 제목 추출 (더 다양한 방법 시도)
                link_selectors = [
                    './/a[contains(@href, "/email/")]',  # 이메일 링크 우선
                    './/td[1]//a',                        # 첫 번째 열의 링크
                    './/a',                               # 모든 링크
                    './/td//a',                           # 테이블 셀의 링크
                    './/div//a',                          # div 내의 링크
                    './/span//a'                          # span 내의 링크
                ]
                
                title_element = None
                for selector in link_selectors:
                    try:
                        title_element = row.find_element(By.XPATH, selector)
                        print(f"제목 요소 찾음: {selector}")
                        break
                    except:
                        continue
                
                if not title_element:
                    print(f"행 {i+1}: 제목 요소를 찾을 수 없음")
                    continue
                
                title = title_element.text.strip()
                href = title_element.get_attribute('href')
                print(f"행 {i+1}: 제목='{title}', 링크='{href}'")
                
                # 이메일 ID 추출 (URL에서)
                email_id_match = re.search(r'/email/(\d+)', href)
                email_id = email_id_match.group(1) if email_id_match else None
                
                if not email_id or not title:
                    print(f"행 {i+1}: 이메일 ID 또는 제목이 없음")
                    continue
                
                print(f"행 {i+1}: 이메일 ID={email_id}")
                
                # 발송일 추출 (더 정확한 방법)
                send_date = None
                date_selectors = [
                    './/td[contains(@class, "date")]',           # date 클래스가 있는 셀
                    './/td[contains(@class, "send-date")]',      # send-date 클래스가 있는 셀
                    './/td[contains(@class, "created")]',        # created 클래스가 있는 셀
                    './/td[3]',                                  # 3번째 열
                    './/td[4]',                                  # 4번째 열
                    './/td[2]',                                  # 2번째 열
                    './/td[5]',                                  # 5번째 열
                    './/div[contains(@class, "date")]',          # date 클래스가 있는 div
                    './/span[contains(@class, "date")]',         # date 클래스가 있는 span
                    './/div[contains(text(), "-")]',             # 하이픈이 포함된 div (날짜 형식)
                    './/span[contains(text(), "-")]'             # 하이픈이 포함된 span (날짜 형식)
                ]
                
                for date_selector in date_selectors:
                    try:
                        date_element = row.find_element(By.XPATH, date_selector)
                        date_text = date_element.text.strip()
                        print(f"행 {i+1}: 날짜 선택자 '{date_selector}'에서 '{date_text}' 발견")
                        
                        # 날짜 형식 파싱 (더 다양한 형식 지원)
                        if re.match(r'\d{4}-\d{2}-\d{2}', date_text):
                            send_date = datetime.strptime(date_text, '%Y-%m-%d')
                            print(f"행 {i+1}: YYYY-MM-DD 형식으로 파싱됨: {send_date}")
                            break
                        elif re.match(r'\d{2}/\d{2}/\d{4}', date_text):
                            send_date = datetime.strptime(date_text, '%m/%d/%Y')
                            print(f"행 {i+1}: MM/DD/YYYY 형식으로 파싱됨: {send_date}")
                            break
                        elif re.match(r'\d{2}-\d{2}-\d{4}', date_text):
                            send_date = datetime.strptime(date_text, '%m-%d-%Y')
                            print(f"행 {i+1}: MM-DD-YYYY 형식으로 파싱됨: {send_date}")
                            break
                        elif '분 전' in date_text or '시간 전' in date_text or '일 전' in date_text:
                            # 상대적 시간 표현 처리
                            send_date = current_date  # 현재 시간으로 처리
                            print(f"행 {i+1}: 상대적 시간으로 처리됨: {send_date}")
                            break
                        elif re.match(r'\d{4}년\d{1,2}월\d{1,2}일', date_text):
                            # 한국어 날짜 형식
                            send_date = datetime.strptime(date_text, '%Y년%m월%d일')
                            print(f"행 {i+1}: 한국어 형식으로 파싱됨: {send_date}")
                            break
                    except Exception as e:
                        print(f"행 {i+1}: 날짜 파싱 실패 (선택자: {date_selector}): {e}")
                        continue
                
                if send_date:
                    # 한달 이내 발송된 이메일만 포함
                    if send_date >= one_month_ago:
                        email_data = {
                            'id': email_id,
                            'subject': title,
                            'url': href,
                            'send_date': send_date.strftime('%Y-%m-%d')
                        }
                        recent_emails.append(email_data)
                        print(f"✅ 행 {i+1}: 포함된 이메일 - ID={email_id}, 발송일={send_date.strftime('%Y-%m-%d')}, 제목={title}")
                    else:
                        print(f"⏭️ 행 {i+1}: 제외된 이메일 - ID={email_id}, 발송일={send_date.strftime('%Y-%m-%d')} (한달 이전)")
                else:
                    # 발송일을 추출할 수 없는 경우 포함 (안전장치)
                    email_data = {
                        'id': email_id,
                        'subject': title,
                        'url': href,
                        'send_date': 'unknown'
                    }
                    recent_emails.append(email_data)
                    print(f"⚠️ 행 {i+1}: 발송일 미확인 이메일 포함 - ID={email_id}, 제목={title}")
                
            except Exception as e:
                print(f"행 {i+1} 처리 오류: {e}")
                continue
        
        # 발송일 기준 최신순 정렬
        recent_emails.sort(key=lambda x: x.get('send_date', '1900-01-01'), reverse=True)
        
        print(f"한달 이내 발송완료 이메일 수: {len(recent_emails)}")
        return recent_emails
                
    except Exception as e:
        print(f"발송완료 이메일 목록 스크래핑 오류: {e}")
        return []

def get_email_list_from_api():
    """스티비 API로 최근 이메일 목록 가져오기"""
    # 웹 스크래핑으로 발송일 기준 필터링 사용
    print("발송일 기준 필터링을 위해 웹 스크래핑을 사용합니다.")
    return get_email_list_from_website()
    
    # 기존 API 코드는 주석 처리
    # if not STIBEE_API_KEY or STIBEE_API_KEY == 'your_api_key_here':
    #     print("API 키가 없거나 설정되지 않아서 웹사이트에서 스크래핑합니다.")
    #     return get_email_list_from_website()
    
    # 스티비 API 문서에 따른 정확한 엔드포인트 (v2)
    # api_url = 'https://api.stibee.com/v2/emails'
    # headers = {
    #     'AccessToken': STIBEE_API_KEY,
    #     'Content-Type': 'application/json'
    # }
    
    # # API 파라미터 (전체 이메일 가져오기)
    # params = {
    #     'limit': 1000,  # 충분히 큰 값으로 설정
    #     'offset': 0
    # }
    
    # try:
    #     print(f"API 요청: {api_url}")
    #     print(f"헤더: AccessToken={STIBEE_API_KEY[:10]}...")
    #     print(f"파라미터: {params}")
        
    #     response = requests.get(api_url, headers=headers, params=params)
    #     print(f"응답 상태 코드: {response.status_code}")
        
    #     if response.status_code == 200:
    #         data = response.json()
    #         print(f"응답 데이터 구조: {list(data.keys()) if isinstance(data, dict) else 'Not a dict'}")
            
    #         # API 응답 구조에 따라 이메일 목록 추출
    #         if 'items' in data:
    #             emails = data['items']
    #             print(f"총 이메일 수: {data.get('total', 0)}")
    #         elif 'data' in data:
    #             emails = data['data']
    #         elif isinstance(data, list):
    #             emails = data
    #         else:
    #             print(f"예상치 못한 응답 구조: {data}")
    #             return []
            
    #         # 최근 순으로 정렬 (created_at 기준)
    #         if emails and 'created_at' in emails[0]:
    #             emails.sort(key=lambda x: x.get('created_at', ''), reverse=True)
            
    #         print(f"API에서 {len(emails)}개의 이메일을 가져왔습니다.")
    #         for email in emails[:5]:  # 처음 5개만 출력
    #             print(f"  - ID: {email.get('id')}, 제목: {email.get('subject', 'N/A')}")
    #         return emails
    #     else:
    #         print(f"API 오류: {response.status_code}")
    #         print(f"오류 내용: {response.text}")
            
    #         # 권한 관련 오류인지 확인
    #         if response.status_code == 403:
    #             print("❌ 권한 오류: 프로 또는 엔터프라이즈 요금제가 필요합니다.")
    #         elif response.status_code == 401:
    #             print("❌ 인증 오류: API 키가 유효하지 않습니다.")
            
    #         print("웹사이트에서 스크래핑으로 전환합니다.")
    #         return get_email_list_from_website()
    # except Exception as e:
    #     print(f"API 요청 오류: {e}")
    #     print("웹사이트에서 스크래핑으로 전환합니다.")
    #     return get_email_list_from_website()

def ensure_logged_in():
    # 로그인 페이지로 리다이렉트되었는지 확인하고, 필요시 재로그인
    global driver
    if driver is None:
        driver = uc.Chrome(options=opts, headless=False)  # 브라우저를 보이게 설정
    if 'login' in driver.current_url:
        print("세션 만료, 재로그인 시도")
        stibee_login()

def extract_unsubscribes(email_id):
    """특정 이메일 ID의 수신거부 리스트를 수집 (모든 열 포함)"""
    deny_url = f'https://stibee.com/email/{email_id}/logs/deny'
    print(f"수신거부 페이지 접속: {deny_url}")
    
    try:
        global driver
        if driver is None:
            driver = uc.Chrome(options=opts, headless=False)  # 브라우저를 보이게 설정
        driver.get(deny_url)
        
        # 페이지 로딩 대기 (최대 30초)
        wait = WebDriverWait(driver, 30)
        
        # 로그인 페이지로 리다이렉트되었는지 확인
        if 'login' in driver.current_url:
            print(f"❌ {email_id}: 세션 만료로 로그인 페이지로 이동됨")
            # 재로그인 시도
            if stibee_login():
                driver.get(deny_url)
                wait = WebDriverWait(driver, 30)
            else:
                raise Exception("재로그인 실패")
        
        # 제목 추출 시도
        try:
            subject_element = wait.until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="root"]/div/div[1]/div/section/main/div/div/div[1]/div[1]/h1'))
            )
            subject = subject_element.text.strip()
        except Exception as e:
            print(f"⚠️ {email_id}: 제목 추출 실패 - {e}")
            subject = f"이메일ID_{email_id}"
        
        # 테이블 로딩 대기
        try:
            wait.until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="root"]/div/div[1]/div/section/main/div/div/div[5]/div/div/div/div/div/table/tbody/tr'))
            )
        except Exception as e:
            print(f"⚠️ {email_id}: 테이블 로딩 실패 - {e}")
            # 테이블이 없는 경우 빈 리스트 반환
            return subject, []
        
        unsubscribes = []
        
        # 테이블의 모든 행을 찾기
        try:
            rows = driver.find_elements(By.XPATH, '//*[@id="root"]/div/div[1]/div/section/main/div/div/div[5]/div/div/div/div/div/table/tbody/tr')
            
            if not rows:
                print(f"⚠️ {email_id}: 수신거부 테이블 행을 찾을 수 없음")
                return subject, []
            
            print(f"📋 {email_id}: {len(rows)}개의 수신거부 행 발견")
            
            for i, row in enumerate(rows, 1):
                try:
                    # 필요한 5개 열만 수집: [이메일주소, 이름, 구독자ID, 구독일, 발송일, 수신일, 수신거부일]
                    row_data = []
                    
                    # 1번째 열: 이메일 주소
                    try:
                        email_element = row.find_element(By.XPATH, f'./td[1]/a')
                        email = email_element.text.strip()
                        row_data.append(email)
                    except:
                        email = ""
                        row_data.append(email)
                    
                    # 2번째 열: 이름
                    try:
                        name_element = row.find_element(By.XPATH, f'./td[2]')
                        name = name_element.text.strip()
                        row_data.append(name)
                    except:
                        name = ""
                        row_data.append(name)
                    
                    # 3번째 열: 구독자 ID
                    try:
                        subscriber_id_element = row.find_element(By.XPATH, f'./td[3]')
                        subscriber_id = subscriber_id_element.text.strip()
                        row_data.append(subscriber_id)
                    except:
                        subscriber_id = ""
                        row_data.append(subscriber_id)
                    
                    # 4번째 열: 구독일
                    try:
                        subscribe_date_element = row.find_element(By.XPATH, f'./td[4]')
                        subscribe_date = subscribe_date_element.text.strip()
                        row_data.append(subscribe_date)
                    except:
                        subscribe_date = ""
                        row_data.append(subscribe_date)
                    
                    # 5번째 열: 발송일
                    try:
                        send_date_element = row.find_element(By.XPATH, f'./td[5]')
                        send_date = send_date_element.text.strip()
                        row_data.append(send_date)
                    except:
                        send_date = ""
                        row_data.append(send_date)
                    
                    # 6번째 열: 수신일
                    try:
                        receive_date_element = row.find_element(By.XPATH, f'./td[6]')
                        receive_date = receive_date_element.text.strip()
                        row_data.append(receive_date)
                    except:
                        receive_date = ""
                        row_data.append(receive_date)
                    
                    # 7번째 열: 수신거부일 (중요!)
                    try:
                        unsubscribe_date_element = row.find_element(By.XPATH, f'./td[7]')
                        unsubscribe_date = unsubscribe_date_element.text.strip()
                        row_data.append(unsubscribe_date)
                    except:
                        unsubscribe_date = ""
                        row_data.append(unsubscribe_date)
                    
                    if email:  # 이메일이 있는 경우만 추가
                        unsubscribes.append(row_data)
                        print(f"  수신거부자 {i}: {email} ({name}) - 수신거부일: {unsubscribe_date}")
                    else:
                        print(f"  ⚠️ 행 {i}: 이메일 주소가 비어있음")
                        
                except Exception as e:
                    print(f"  ❌ 행 {i} 추출 실패: {e}")
                    continue
            
        except Exception as e:
            print(f"❌ {email_id}: 테이블 행 추출 실패 - {e}")
            return subject, []
        
        print(f"✅ {email_id}: 총 {len(unsubscribes)}명의 수신거부자 수집 완료 (모든 열 포함)")
        return subject, unsubscribes
        
    except Exception as e:
        print(f"❌ {email_id}: 수신거부 페이지 처리 실패 - {e}")
        return f"이메일ID_{email_id}", []

def extract_failures(email_id):
    """특정 이메일 ID의 발송실패 리스트를 수집 (페이지네이션 처리)"""
    fail_url = f'https://stibee.com/email/{email_id}/logs/fail'
    print(f"발송실패 페이지 접속: {fail_url}")
    
    try:
        global driver
        if driver is None:
            driver = uc.Chrome(options=opts, headless=False)  # 브라우저를 보이게 설정
        driver.get(fail_url)
        time.sleep(5)  # 페이지 로딩 대기
        
        # 페이지 로딩 대기 (최대 30초)
        wait = WebDriverWait(driver, 30)
        
        # 로그인 페이지로 리다이렉트되었는지 확인
        if 'login' in driver.current_url:
            print(f"❌ {email_id}: 세션 만료로 로그인 페이지로 이동됨")
            # 재로그인 시도
            if stibee_login():
                driver.get(fail_url)
                time.sleep(5)
                wait = WebDriverWait(driver, 30)
            else:
                raise Exception("재로그인 실패")
        
        # 제목 추출 시도
        try:
            subject_element = wait.until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="root"]/div/div[1]/div/section/main/div/div/div[1]/div[1]/h1'))
            )
            subject = subject_element.text.strip()
            print(f"📧 {email_id}: 제목 - {subject}")
        except Exception as e:
            print(f"⚠️ {email_id}: 제목 추출 실패 - {e}")
            subject = f"이메일ID_{email_id}"
        
        # 발송실패 수 확인 (정확한 XPATH 사용)
        fail_count = 0
        try:
            fail_count_element = wait.until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="root"]/div/div[1]/div/section/main/div/div/div[4]/div/div[1]/div/em'))
            )
            fail_count_text = fail_count_element.text.strip()
            print(f"발송실패 수 텍스트: '{fail_count_text}'")
            # 숫자만 추출
            import re
            numbers = re.findall(r'\d+', fail_count_text)
            if numbers:
                fail_count = int(numbers[0])
                print(f"📊 {email_id}: 총 발송실패 수 {fail_count}건")
            else:
                print(f"⚠️ {email_id}: 발송실패 수를 추출할 수 없습니다. 텍스트: {fail_count_text}")
        except Exception as e:
            print(f"⚠️ {email_id}: 발송실패 수 추출 실패 - {e}")
            # 테이블이 있는지 확인
            try:
                rows = driver.find_elements(By.XPATH, '//*[@id="root"]/div/div[1]/div/section/main/div/div/div[5]/div/div/div/div/div/table/tbody/tr')
                if rows:
                    fail_count = len(rows)
                    print(f"📊 {email_id}: 테이블에서 {fail_count}개의 발송실패 행 발견")
                else:
                    print(f"⏭️ {email_id}: 발송실패가 없습니다.")
                    return subject, []
            except:
                print(f"⏭️ {email_id}: 발송실패가 없습니다.")
                return subject, []
        
        if fail_count == 0:
            print(f"⏭️ {email_id}: 발송실패가 없습니다.")
            return subject, []
        
        # 대용량 데이터 처리 준비
        if fail_count > 100:
            print(f"�� {email_id}: 대용량 발송실패 데이터 감지 ({fail_count}건)")
            print(f"   예상 처리 시간: 약 {fail_count // 20 * 3}초")
        
        failures = []
        page = 1
        collected_count = 0
        max_pages = (fail_count + 19) // 20  # 페이지당 20개씩, 올림 처리
        consecutive_empty_pages = 0
        max_consecutive_empty = 3  # 연속 3페이지가 비어있으면 중단
        
        print(f"📄 {email_id}: 예상 총 페이지 수 {max_pages}페이지")
        
        while collected_count < fail_count and consecutive_empty_pages < max_consecutive_empty:
            print(f"📄 {email_id}: {page}페이지 처리 중... (수집된 수: {collected_count}/{fail_count}, 진행률: {collected_count/fail_count*100:.1f}%)")
            
            # 현재 페이지의 테이블 로딩 대기
            try:
                wait.until(
                    EC.presence_of_element_located((By.XPATH, '//*[@id="root"]/div/div[1]/div/section/main/div/div/div[5]/div/div/div/div/div/table/tbody/tr'))
                )
                time.sleep(3)  # 테이블 완전 로딩 대기
            except Exception as e:
                print(f"⚠️ {email_id}: {page}페이지 테이블 로딩 실패 - {e}")
                consecutive_empty_pages += 1
                continue
            
            # 현재 페이지의 모든 행 수집
            try:
                rows = driver.find_elements(By.XPATH, '//*[@id="root"]/div/div[1]/div/section/main/div/div/div[5]/div/div/div/div/div/table/tbody/tr')
                
                if not rows:
                    print(f"⚠️ {email_id}: {page}페이지에 발송실패 행이 없음")
                    consecutive_empty_pages += 1
                    if consecutive_empty_pages >= max_consecutive_empty:
                        print(f"❌ {email_id}: 연속 {max_consecutive_empty}페이지가 비어있어 수집을 중단합니다.")
                        break
                    continue
                else:
                    consecutive_empty_pages = 0  # 데이터가 있으면 카운터 리셋
                
                print(f"📋 {email_id}: {page}페이지에서 {len(rows)}개의 발송실패 행 발견")
                
                page_failures = 0
                for i, row in enumerate(rows, 1):
                    try:
                        # 각 행에서 이메일, 이름, 날짜, 실패 이유 추출
                        email_element = row.find_element(By.XPATH, f'./td[1]/a')
                        email = email_element.text.strip()
                        
                        name_element = row.find_element(By.XPATH, f'./td[2]')
                        name = name_element.text.strip()
                        
                        date_element = row.find_element(By.XPATH, f'./td[5]')
                        date = date_element.text.strip()
                        
                        # 실패 이유 추출 (6번째 열)
                        try:
                            reason_element = row.find_element(By.XPATH, f'./td[6]')
                            reason = reason_element.text.strip()
                        except:
                            reason = "알 수 없음"
                        
                        if email:  # 이메일이 있는 경우만 추가
                            failure_data = [email, name, date, reason]
                            failures.append(failure_data)
                            page_failures += 1
                            if fail_count <= 50 or page_failures % 10 == 0:  # 대용량일 때는 10개마다만 출력
                                print(f"  발송실패 {collected_count + i}: {email} ({name}) - {date} - {reason}")
                        else:
                            print(f"  ⚠️ 행 {i}: 이메일 주소가 비어있음")
                            
                    except Exception as e:
                        print(f"  ❌ 행 {i} 추출 실패: {e}")
                        continue
                
                collected_count += len(rows)
                print(f"현재까지 수집된 발송실패: {collected_count}건 (이번 페이지: {page_failures}건)")
                
                # 다음 페이지로 이동 (발송실패 수만큼 수집했으면 종료)
                if collected_count >= fail_count:
                    print(f"✅ {email_id}: 모든 발송실패 수집 완료 ({collected_count}/{fail_count})")
                    break
                
                # 다음 페이지 버튼 클릭
                try:
                    next_button = driver.find_element(By.XPATH, '//*[@id="root"]/div/div[1]/div/section/main/div/div/div[7]/div/div[3]/button')
                    if 'disabled' in next_button.get_attribute('class') or 'disabled' in next_button.get_attribute('aria-disabled'):
                        print(f"⏭️ {email_id}: 다음 페이지가 없습니다.")
                        break
                    
                    next_button.click()
                    time.sleep(3)  # 페이지 로딩 대기
                    page += 1
                    
                    # 대용량 데이터 처리 시 추가 대기
                    if fail_count > 100 and page % 5 == 0:
                        print(f"⏸️ {email_id}: 대용량 데이터 처리 중 잠시 대기...")
                        time.sleep(5)
                    
                except Exception as e:
                    print(f"⚠️ {email_id}: 다음 페이지 버튼 클릭 실패 - {e}")
                    break
                    
            except Exception as e:
                print(f"❌ {email_id}: {page}페이지 처리 실패 - {e}")
                consecutive_empty_pages += 1
                continue
        
        print(f"✅ {email_id}: 총 {len(failures)}건의 발송실패 수집 완료 (예상: {fail_count}건)")
        return subject, failures
        
    except Exception as e:
        print(f"❌ {email_id}: 발송실패 페이지 처리 실패 - {e}")
        return f"이메일ID_{email_id}", []



def get_existing_email_ids():
    """구글시트에서 이미 기록된 이메일 ID 목록을 가져오기 (발송일자 열이 추가된 경우 2번째 열 기준)"""
    try:
        sheet1 = client.open_by_url(sheet_url).worksheet('메일 발송 성과')
        # B열에서 이메일 ID 목록 가져오기 (헤더 제외)
        existing_ids = sheet1.col_values(2)[1:]  # 첫 번째 행(헤더) 제외
        existing_ids = [str(id).strip() for id in existing_ids if str(id).strip()]
        print(f"이미 기록된 이메일 ID 수: {len(existing_ids)}")
        return set(existing_ids)
    except Exception as e:
        print(f"기존 이메일 ID 조회 오류: {e}")
        return set()

def get_existing_unsubscribes():
    """구글시트에서 이미 기록된 수신거부 데이터를 가져오기"""
    try:
        unsub_sheet = client.open_by_url(sheet_url).worksheet('수신거부리스트')
        existing_unsub = set()
        for row in unsub_sheet.get_all_values()[1:]:  # 헤더 제외
            if len(row) >= 3:
                existing_unsub.add((row[0], row[2]))  # (이메일ID, 수신거부자 이메일)
        print(f"기존 수신거부 데이터 수: {len(existing_unsub)}")
        return existing_unsub
    except Exception as e:
        print(f"기존 수신거부 데이터 조회 오류: {e}")
        return set()

def update_unsubscribes_last_month():
    """한 달 이내 모든 이메일의 수신거부리스트를 확인하고, 신규만 추가"""
    emails = get_email_list_from_api()  # 이제 웹 스크래핑 기반으로 동작
    if not emails:
        print("이메일 목록을 가져올 수 없습니다.")
        return

    # 웹 스크래핑에서 이미 한달 필터링이 적용되어 있음
    print(f"한달 이내 이메일 수(수신거부 체크): {len(emails)}")

    existing_unsub = get_existing_unsubscribes()

    for i, email in enumerate(emails, start=1):
        try:
            email_id = str(email['id'])
            subject, unsubscribes = extract_unsubscribes(email_id)
            
            # 새로운 수신거부 데이터만 필터링 (이메일 ID + 수신거부자 이메일 기준)
            new_unsubs = []
            for unsub in unsubscribes:
                if len(unsub) >= 1:  # 최소 이메일 주소는 있어야 함
                    unsubscribe_email = unsub[0]  # 첫 번째 열이 이메일 주소
                    if (email_id, unsubscribe_email) not in existing_unsub:
                        # 구글 독스 열 순서: [이메일ID, 메일 제목, 수신거부자 이메일, 이름, 날짜(수신거부일)]
                        # unsub[0]=이메일주소, unsub[1]=이름, unsub[6]=수신거부일
                        row_data = [
                            email_id,                    # 1열: 이메일ID
                            subject,                     # 2열: 메일 제목
                            unsub[0],                    # 3열: 수신거부자 이메일
                            unsub[1],                    # 4열: 이름
                            unsub[6] if len(unsub) > 6 else ""  # 5열: 수신거부일
                        ]
                        new_unsubs.append(row_data)
            
            if new_unsubs:
                try:
                    sheet2 = client.open_by_url(sheet_url).worksheet('수신거부리스트')
                    for row in new_unsubs:
                        sheet2.append_row(row)
                    print(f"{email_id}의 새로운 수신거부 {len(new_unsubs)}건 추가")
                except Exception as e:
                    print(f"수신거부리스트 업데이트 오류: {e}")
            else:
                print(f"{email_id}의 새로운 수신거부 없음")
        except Exception as e:
            print(f"❌ 오류: {email_id} - {str(e)}")
            continue
        time.sleep(2)
    print(f'🎉 한달 이내 수신거부리스트 업데이트 완료!')




if __name__ == '__main__':
    # 1) 로그인 시도
    try:
        if stibee_login():
            print('로그인 성공')
        else:
            print('로그인 실패 또는 세션 만료 상태')
            exit(1)  # 로그인 실패 시 종료
    except Exception as e:
        print(f'로그인 중 오류: {e}')
        exit(1)

    # 2) 최근 한달 수신거부 업데이트 실행
    try:
        update_unsubscribes_last_month()
    except Exception as e:
        print(f'업데이트 실행 중 오류: {e}')

