import pandas as pd
import numpy as np

# 예시 데이터 (group_data)
group_data = {
    '경기1': ['가평군', '고양시', '과천시', '광명시', '광주시', '구리시', '군포시', '화성시'],
    '경기2': ['김포시', '남양주시', '동두천시', '부천시', '성남시', '수원시', '시흥시', '안산시', '안성시', '안양시', '양주시', '양평군', '여주시', '연천군', '오산시', '용인시', '의왕시', '의정부시', '이천시', '파주시', '평택시', '포천시', '하남시']
}

max_len = max(len(v) for v in group_data.values())

# 길이가 짧은 리스트는 NaN으로 채워서 맞추기
for key in group_data:
    while len(group_data[key]) < max_len:
        group_data[key].append(np.nan)

# dict를 DataFrame으로 변환
temp_df = pd.DataFrame(group_data)

# 결과 출력
print(temp_df)
print(temp_df.shape)