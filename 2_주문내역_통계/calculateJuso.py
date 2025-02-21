import pandas as pd
from openpyxl import Workbook

print("🚀 실행 시작...")

# ✅ 1. 엑셀 데이터 불러오기
file_path = "ibco_juso_output.xlsx"  # 📂 엑셀 파일명을 입력하세요
print(f"📂 엑셀 파일 로드 중: {file_path}")

df = pd.read_excel(file_path)

# ✅ 2. 데이터 확인 (컬럼명 변경 필요 시 수정)
print("🔍 실제 컬럼명:", df.columns)  # 컬럼 확인
df.columns = ["시/도", "구/군", "동"]
print("✅ 컬럼명 변경 완료:", df.columns)

# ✅ 3. 시/도별 개수 및 비율 계산
sido_counts = df["시/도"].value_counts().reset_index()
sido_counts.columns = ["시/도", "개수"]
sido_counts["비율"] = (sido_counts["개수"] / sido_counts["개수"].sum()) * 100

print("📊 시/도별 개수 계산 완료!")
print(sido_counts.head())  # 상위 5개만 출력

# ✅ 4. 구/군별 개수 및 비율 계산
sido_gu_stats = {}  # 시/도별 구/군 데이터 저장용 딕셔너리

for sido in df["시/도"].unique():
    gu_counts = df[df["시/도"] == sido]["구/군"].value_counts().reset_index()
    gu_counts.columns = ["구/군", "개수"]
    gu_counts["비율"] = (gu_counts["개수"] / gu_counts["개수"].sum()) * 100
    sido_gu_stats[sido] = gu_counts
    print(f"📊 {sido} - 구/군 데이터 계산 완료!")

# ✅ 5. 결과 엑셀 파일 생성
output_file = "통계_결과.xlsx"
wb = Workbook()
ws_sido = wb.active
ws_sido.title = "시도별 통계"

# ✅ 6. 시/도별 통계 저장
ws_sido.append(["시/도", "개수", "비율"])
for _, row in sido_counts.iterrows():
    ws_sido.append(row.tolist())

print("✅ 시/도별 통계 엑셀 저장 완료!")

# ✅ 7. 각 시/도별 구/군 통계 저장
for sido, gu_df in sido_gu_stats.items():
    ws_gu = wb.create_sheet(title=sido)
    ws_gu.append(["구/군", "개수", "비율"])
    for _, row in gu_df.iterrows():
        ws_gu.append(row.tolist())
    print(f"✅ {sido} 구/군별 통계 엑셀 저장 완료!")

# ✅ 8. 엑셀 파일 저장
wb.save(output_file)
print(f"✅ 엑셀 파일 저장 완료: {output_file}")

print("🎉 실행 완료!")
