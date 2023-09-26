import pandas as pd

df = pd.read_excel('カテゴリーのみ２.xlsx')

topics_categories = {}
current_topic = None

for index, row in df.iterrows():
    if row['カテゴリー'].startswith('TOPIC-'):
        if current_topic:
            topics_categories[current_topic] = current_category
        current_topic = row['カテゴリー']
        current_category = None
    else:
        current_category = row['カテゴリー']

if current_topic and current_category:
    topics_categories[current_topic] = current_category

for topic, category in topics_categories.items():
    print(f"{topic}: {category}")

result_df = pd.DataFrame(list(topics_categories.items()), columns=['トピック', 'カテゴリー'])
result_df.to_excel('結果データ　最後の出来事.xlsx', index=False)