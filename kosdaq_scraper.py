import requests
import pandas as pd
from io import StringIO
from datetime import datetime, timedelta
import time

def get_kosdaq_growth_data():
    # 1. 네이버 금융 항목 설정 (19:매출액증가율, 20:영업이익증가율, 27:외국인비율, 6:거래량, 4:전일비)
    target_fields = "019|020|027|006|004|"
    
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
        'Referer': 'https://finance.naver.com/sise/sise_market_sum.naver?sosok=1',
        'Cookie': f'field_list={target_fields}'
    }

    all_data = []
    print("🚀 코스닥 데이터 수집을 시작합니다... (증가율 데이터 포함)")

    # 2. 데이터 수집 (최대 45페이지까지 확인)
    for page in range(1, 45):
        url = f"https://finance.naver.com/sise/sise_market_sum.naver?sosok=1&page={page}"
        try:
            res = requests.get(url, headers=headers, timeout=10)
            
            # 데이터가 없는 페이지거나 장이 열리지 않은 날의 빈 페이지 체크
            if "데이터가 없습니다" in res.text or "등록된 종목이 없습니다" in res.text:
                break
                
            df_list = pd.read_html(StringIO(res.text), encoding='euc-kr')
            if len(df_list) < 2: continue # 유효한 테이블이 없으면 건너뜀
            
            df = df_list[1]
            df = df[df['종목명'].notna()] # 빈 줄 제거
            df = df.loc[:, ~df.columns.str.contains('Unnamed|토론')] # 불필요 컬럼 제거
            
            all_data.append(df)
            print(f"📡 {page}페이지 수집 완료", end='\r')
            time.sleep(0.2)
        except Exception as e:
            print(f"\n⚠️ {page}페이지에서 오류 발생: {e}")
            break

    # 3. 결과 합치기 및 저장
    if all_data:
        final_df = pd.concat(all_data, ignore_index=True)
        
        # 컬럼명 정리
        if 'N' in final_df.columns:
            final_df.rename(columns={'N': '시총순위'}, inplace=True)

        # ⭐️수정된 부분: UTC 시간에 9시간을 더해 한국 시간(KST)으로 맞춤⭐️
        kst_time = datetime.utcnow() + timedelta(hours=9)
        file_name = f"KOSDAQ_FINAL_{kst_time.strftime('%Y%m%d_%H%M')}.xlsx"
        
        final_df.to_excel(file_name, index=False)
        
        print("\n" + "="*50)
        print(f"✅ 수집 성공! 파일명: {file_name}")
        print(f"📊 수집된 항목: {list(final_df.columns)}")
        print("="*50)
        
        # 성공 여부 체크
        if '영업이익증가율' in final_df.columns:
            print("✨ 축하합니다! 모든 데이터가 정상적으로 수집되었습니다.")
        else:
            print("❗ 경고: 증가율 데이터가 누락되었습니다. 네이버 쿠키 정책을 다시 확인해야 합니다.")
    else:
        print("\n❌ 수집된 데이터가 없습니다. (오늘은 시장이 열리지 않는 날일 수 있습니다)")

if __name__ == "__main__":
    get_kosdaq_growth_data()
