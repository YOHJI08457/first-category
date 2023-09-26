import pandas as pd

# エクセルファイルの読み込み
df = pd.read_excel('カテゴリーのみ２.xlsx')

# TOPICごとにカテゴリーの数を数えて多数決を行う
topics_categories = {}
current_topic = None

for index, row in df.iterrows():
    if row['カテゴリー'].startswith('TOPIC-'):
        current_topic = row['カテゴリー']
    elif current_topic and current_topic not in topics_categories:
        topics_categories[current_topic] = row['カテゴリー']

# 結果の表示
for topic, category in topics_categories.items():
    print(f"{topic}: {category}")

# 結果をエクセルファイルに書き込む
result_df = pd.DataFrame(list(topics_categories.items()), columns=['トピック', 'カテゴリー'])
result_df.to_excel('結果データ.xlsx', index=False)