import pandas as pd
import numpy as np

# df = pd.read_excel(r"C:\Users\m\Desktop\2019年毕节市公安局50名事业编招聘信息表及照片\领取准考证确认单.xls",encoding="gb18030")
# split_columns = df['工作岗位'].unique()
#
# for split_column in split_columns:
#     df[df['工作岗位']==split_column].to_excel("{}.xls".format(split_column),index=False)
# print("工作表拆分结束")

ts = pd.Series(np.random.randn(1000),index=pd.date_range('1/1/2000', periods=1000))
ts = ts.cumsum()
ts.plot()
