import requests
import pandas as pd
from io import StringIO
from datetime import datetime, timedelta
import time
import sys

def get_market_growth_data(market_name, sosok_code):
    # 1. 네이버 금융 항목 설정 (19:매출액증가율, 20:영업이익증가율, 27:외국인비율, 6:거래량, 4:전일비)
    target_fields = "019|020|027|006|004|"
    
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
        'Referer': f'https://finance.naver.com/sise/sise_market_sum.naver?sosok={sosok_code}',
        'Cookie': f'field_list={target_fields}',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8',
        'Accept-Language': 'ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7',
        'Connection': 'keep-alive'
    }

    all_data = []
    print(f"\n🚀 {market_name} 데이터 수집을 시작합니다... (증가율 데이터 포함)")

    # 2. 데이터 수집 (코스피, 코스닥 모두 대략 45페이지 내외)
    for page in range(1, 45):
        url = f"https://finance.naver.com/sise/sise_market_sum.naver?sosok={sosok_code}&page={page}"
        try:
            res = requests.get(url, headers=headers, timeout=10)
            
            if res.status_code != 200:
                print(f"\n🚫 [오류] 페이지 {page} 접근 실패! HTTP 상태 코드: {res.status_code}")
                sys.exit(1) # 에러 발생 시 강제 종료하여 GitHub Actions 실패 처리
                
            # 데이터가 없는 페이지거나 장이 열리지 않은 날의 빈 페이지 체크
            if "데이터가 없습니다" in res.text or "등록된 종목이 없습니다" in res.text:
                break
                
            # HTML 테이블 파싱
            df_list = pd.read_html(StringIO(res.text), encoding='euc-kr')
            if len(df_list) < 2: 
                continue # 유효한 테이블이 없으면 건너뜀
            
            df = df_list[1]
            df = df[df['종목명'].notna()] # 빈 줄 제거
            df = df.loc[:, ~df.columns.str.contains('Unnamed|토론')] # 불필요 컬럼 제거
            
            all_data.append(df)
            print(f"📡 {market_name} {page}페이지 수집 완료", end='\r')
            
            # Rate Limit 회피
            time.sleep(1.0) 
            
        except Exception as e:
            print(f"\n⚠️ {page}페이지에서 크롤링 중 예외 오류 발생: {e}")
            sys.exit(1) # 에러 발생 시 강제 종료

    # 3. 결과 합치기 및 저장
    if all_data:
        final_df = pd.concat(all_data, ignore_index=True)
        
        # 컬럼명 정리
        if 'N' in final_df.columns:
            final_df.rename(columns={'N': '시총순위'}, inplace=True)

        # KST(한국 시간) 변환
        kst_time = datetime.utcnow() + timedelta(hours=9)
        file_name = f"{market_name}_FINAL_{kst_time.strftime('%Y%m%d_%H%M')}.xlsx"
        
        # ✨ 명시적 openpyxl 엔진 사용
        final_df.to_excel(file_name, index=False, engine='openpyxl')
        
        print("\n" + "="*50)
        print(f"✅ {market_name} 수집 성공! 파일명: {file_name}")
        print("="*50)
    else:
        print(f"\n❌ {market_name} 수집된 데이터가 없습니다.")

if __name__ == "__main__":
    # 코스피 (sosok=0) 수집
    get_market_growth_data("KOSPI", "0")
    
    # 네이버 서버에 부담을 주지 않기 위해 시장 변경 전 잠시 대기
    time.sleep(3)
    
    # 코스닥 (sosok=1) 수집
    get_market_growth_data("KOSDAQ", "1")