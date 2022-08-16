import requests
import pandas as pd
import os

# 存储中英文对应的变量的中文名
word_dict = {"poiAddv": "行政区",
             "poiBsnm": "流域",
             "ql": "流量(立方米/秒)",
             "rvnm": "河名",
             "stcd": "站点代码",
             "stnm": "站名",
             "tm": "时间",
             "webStlc": "站址",
             "wrz": "警戒水位(米)",
             "zl": "水位(米)",
             "dateTime": "日期"}

# 爬取大江大河实时水情
url = 'http://xxfb.mwr.cn/hydroSearch/greatRiver'
return_data = requests.get(url, verify=False)
js = return_data.json()
river_info = dict(js)["result"]["data"]
river_table = pd.DataFrame(river_info[0], index=[0])
for i in range(1, len(river_info)):
    river_table = pd.concat([river_table, pd.DataFrame(river_info[i], index=[0])])
for str_col_name in ["poiAddv", "poiBsnm", "rvnm", "stnm", "webStlc"]:
    river_table[str_col_name] = river_table[str_col_name].apply(lambda s: s.strip())
river_table.columns = [word_dict.get(i) for i in river_table.columns]
river_table.index = range(len(river_table.index))

"""
功能
保存爬取后得到的数据集river_table为.csv文件
输入
out_path:表格文件的输出路径，默认为桌面
"""
out_path = os.path.join(os.path.expanduser("~"), 'Desktop')  # 输出路径，默认为桌面
date_suffix = river_table["时间"].values[0].split(" ")[0]
out_csv = os.path.join(out_path, "全国大江大河_{0}.csv".format(date_suffix))  # 输出文件名
river_table.to_csv(out_csv, encoding='utf_8_sig')

# 打印待保存的数据
print(river_table)
a = input("press any key to quit")
