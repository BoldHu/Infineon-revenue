import pandas as pd
from datetime import datetime, timedelta

# 创建示例数据集
data = {
    'Proposed PGI': [datetime.today().date() - timedelta(days=2), '2023-12-17', None, '2023-01-15'],
    # 添加更多的列作为示例数据
    'Other_Column1': [1, 2, 3, 4],
    'Other_Column2': ['A', 'B', 'C', 'D']
}

revord_df = pd.DataFrame(data)

# 将你提供的代码嵌入一个类或函数
class YourClass:
    def __init__(self, revord_df):
        self.revord_df = revord_df

    def add_proposed_pgi_day(self):
        today = datetime.today().date()

        for index, row in self.revord_df.iterrows():
            print(type(row['Proposed PGI']))
            if pd.isnull(row['Proposed PGI']):
                continue
            elif isinstance(row['Proposed PGI'], pd.Timestamp):
                proposed_pgi = row['Proposed PGI'].date()
            elif isinstance(row['Proposed PGI'], str):
                proposed_pgi = datetime.strptime(row['Proposed PGI'], '%Y-%m-%d').date()
            # if it is datetime.date
            else:
                proposed_pgi = row['Proposed PGI']
            if proposed_pgi < today:
                self.revord_df.at[index, 'Proposed PGI Day'] = today
            else:
                self.revord_df.at[index, 'Proposed PGI Day'] = proposed_pgi

# 使用类
your_instance = YourClass(revord_df)
your_instance.add_proposed_pgi_day()

# 打印结果
print(your_instance.revord_df)
