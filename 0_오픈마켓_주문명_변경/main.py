import pandas as pd

# CSV 파일 읽기
df = pd.read_csv('./input.csv', encoding='utf-8')

# '상품명' 열 가져오기
if '상품명' in df.columns:
    product_names = df['상품명'].dropna().tolist()  # NaN 값 제거 후 리스트로 변환
    
else:
    print("CSV 파일에 '상품명' 열이 없습니다.")

cnt = 0
for p in product_names:
    if '키즈' in p and '트리트' in p and '세트' in p:
        print('모두바른 키즈 샴푸 트리트먼트 세트')
        continue
    if '키즈' in p and '리필' in p and '세트' in p:
        print("모두바른 키즈 샴푸 리필팩 세트")
        continue
    if '키즈' in p and '리필' in p:
        print("모두바른 키즈 리필팩")
        continue
    if '키즈' in p:
        print("모두바른 키즈 샴푸")
        continue
    if ('청소' in p or '틴에이' in p)  and '트리트' in p and '세트' in p:
        print('모두바른 틴에이저 샴푸 트리트먼트 세트')
        continue
    if ('청소' in p or '틴에이' in p)  and '리필' in p and '세트' in p:
        print("모두바른 틴에이저 샴푸 리필팩 세트")
        continue
    if ('청소' in p or '틴에이' in p)  and '리필' in p:
        print("모두바른 틴에이저 리필팩")
        continue
    if ('청소' in p or '틴에이' in p) :
        print("모두바른 틴에이저 샴푸")
        continue
    if ('선크' in p or '썬크' in p):
        print("모두바른 선크림")
        continue
    if '페이셜' in p or '폼' in p:
        print('정말바른 페이셜폼')    
        continue
    if ('워시' in p or '바디' in p) and '로션' in p:
        print("모두바른 바디워시 바디로션 세트")
        continue
    if '로션' in p:
        print("모두바른 바디로션")
        continue
    if ('워시' in p or '바디' in p) or '리필' in p:
        print('모두바른 바디워시 리필팩')
        continue
    if '워시' in p or '바디' in p:
        print('모두바른 바디워시')
        continue
    if '트리트' in p:
        print('모두바른 트리트먼트')
        continue
    if '임산부' in p:
        print('모두바른 임산부 샴푸')
        continue
    print("잘못")
    cnt+=1

print(cnt)
