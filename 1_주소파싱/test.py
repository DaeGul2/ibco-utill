import requests
import urllib.parse
import pandas as pd

# 네가 발급받은 실제 API 키
JUSO_API_KEY = "devU01TX0FVVEgyMDI1MDIyMDExMDM1NDExNTQ4ODc="

def get_address_info(address):
    """ 도로명주소 API를 호출하여 주소 정보를 가져오는 함수 """
    url = "https://www.juso.go.kr/addrlink/addrLinkApi.do"

    params = {
        "currentPage": "1",
        "countPerPage": "10",
        "resultType": "json",
        "confmKey": JUSO_API_KEY,
        "keyword": address
    }

    response = requests.post(url, data=params, headers={"Content-Type": "application/x-www-form-urlencoded"})

    try:
        data = response.json()
    except requests.exceptions.JSONDecodeError:
        return None  # 응답이 JSON 형식이 아니면 None 반환

    if not data or "results" not in data or "juso" not in data["results"] or len(data["results"]["juso"]) == 0:
        return None  # 주소 검색 결과 없음

    return data["results"]["juso"][0]  # 첫 번째 결과 반환

def parse_address(address):
    """ 주소를 공백 기준으로 줄여가면서 검색하여 가장 적절한 주소를 찾는 함수 """
    address_parts = address.split()  # 공백 기준으로 주소 분할

    while address_parts:
        search_address = " ".join(address_parts)  # 현재 남은 주소 조합
        print(f"🔍 주소 검색 시도: {search_address}")

        result = get_address_info(search_address)

        if result:  # 결과가 나오면 멈춤
            print("✅ 주소 검색 성공!")
            return result

        address_parts.pop()  # 마지막 요소 제거 후 다시 검색

    return None  # 끝까지 돌았는데도 검색 결과 없음

def get_bunji(address_info):
    """ API 응답에서 '번지' 값을 올바르게 추출 """
    jibun_main = address_info.get("lnbrMnnm", "").strip()  # 지번 본번
    jibun_sub = address_info.get("lnbrSlno", "").strip()   # 지번 부번

    if not jibun_main:  # 본번이 없으면 "번지 알 수 없음"
        return "번지 알 수 없음"
    
    if jibun_sub and jibun_sub != "0":  # 부번이 있으면 "본번-부번"
        return f"{jibun_main}-{jibun_sub}"
    
    return jibun_main  # 부번이 없으면 본번만 반환

def extract_address_details(original_address, address_info):
    """ API 결과에서 필요한 정보를 추출하여 포맷팅 """
    if not address_info:
        return [original_address, "주소 없음", "주소 없음", "주소 없음", "주소 없음", "번지 알 수 없음"]

    zip_code = address_info.get("zipNo", "없음")
    si_do = address_info.get("siNm", "없음")
    gu_gun = address_info.get("sggNm", "없음")
    dong_eup_myeon = address_info.get("emdNm", "없음")
    bunji = get_bunji(address_info)  # 여기서 올바른 '번지' 추출

    return [original_address, zip_code, si_do, gu_gun, dong_eup_myeon, bunji]

# 엑셀 파일 읽기
input_file = "ibco.xlsx"
output_file = "ibco_juso_output.xlsx"

# 모든 시트 로드
sheets = pd.read_excel(input_file, sheet_name=None)  # sheet_name=None으로 모든 시트 불러오기

# 결과 저장용 딕셔너리
output_sheets = {}

# 각 시트별로 처리
for sheet_name, df in sheets.items():
    print(f"\n📌 처리 중: {sheet_name} 시트")

    # '주소' 컬럼이 존재하는지 확인
    if "주소" not in df.columns:
        print(f"⚠️ {sheet_name} 시트에 '주소' 컬럼이 없습니다. 스킵합니다.")
        continue

    output_data = []

    # 주소 데이터 처리
    for idx, row in df.iterrows():
        original_address = row["주소"]
        best_address = parse_address(original_address)
        output_data.append(extract_address_details(original_address, best_address))

    # 시트별 데이터프레임 생성
    output_sheets[sheet_name] = pd.DataFrame(output_data, columns=["원본주소", "우편번호", "시/도", "구/군", "동", "번지"])

# 결과를 엑셀 파일로 저장 (시트별로 저장)
with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
    for sheet_name, df in output_sheets.items():
        df.to_excel(writer, sheet_name=sheet_name, index=False)

print("\n✅ 변환 완료! 결과가 'output.xlsx'에 저장되었습니다.")
