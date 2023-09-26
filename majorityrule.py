import pandas as pd

df = pd.read_excel('カテゴリーのみ２.xlsx')

topics_categories = {}
current_topic = None
current_categories = []

for index, row in df.iterrows():
    if row['カテゴリー'].startswith('TOPIC-'):
        if current_topic and current_categories:
            topic_category = max(set(current_categories), key=current_categories.count)
            topics_categories[current_topic] = topic_category
        current_topic = row['カテゴリー']
        current_categories = []
    else:
        current_categories.append(row['カテゴリー'])

if current_topic and current_categories:
    topic_category = max(set(current_categories), key=current_categories.count)
    topics_categories[current_topic] = topic_category

for topic, category in topics_categories.items():
    print(f"{topic}: {category}")

result_df = pd.DataFrame(list(topics_categories.items()), columns=['トピック', 'カテゴリー'])
result_df.to_excel('結果データ　多数決.xlsx', index=False)