import os, argparse
import pandas as pd

parser = argparse.ArgumentParser()
parser.add_argument('-f', '--filepath', type=str, default=None)
args = parser.parse_args()

# 输入文件路径
if args.filepath is None:
    filepath = input("请输入文件路径: ")
else:
    filepath = args.filepath
filepath = filepath.strip()
filepath = filepath.replace('"', '')
if os.path.exists(filepath):
    print(f'文件路径: {filepath}')
else:
    print(f'文件路径不存在: {filepath}')
    input("按任意键退出...")
    exit()

# 读入表格
text_df = pd.read_excel(filepath, sheet_name="正文表")[['集数', '标题', '角色', '正文']] # 集数，标题，角色，正文
characters_df = pd.read_excel(filepath, sheet_name="角色表")[['角色', '主播', '性别', '年龄', '人设']] # 角色，主播，性别，年龄，人设
text_df.dropna(how='all', inplace=True)
text_df.fillna(method='ffill', inplace=True)
text_count_df = text_df.groupby(['集数', '标题', '角色']).count().reset_index()[['集数', '标题', '角色', '正文']]
# display(text_count_df)
# display(characters_df)

# 合并表格
merged_df = pd.merge(characters_df, text_count_df, on='角色')
merged_df.rename(columns={'正文': '句数'}, inplace=True)
print(merged_df.columns)
merged_df = merged_df.reindex(columns=['集数', '标题', '角色', '主播', '性别', '年龄', '人设', '句数'])

# 自定义排序序列
custom_order = dict(zip(characters_df['角色'], range(len(characters_df['角色']))))
# 按照A列和B列排序
merged_df['角色'] = pd.Categorical(merged_df['角色'], categories=custom_order, ordered=True)
merged_df = merged_df.sort_values(by=['集数', '角色'])
print(merged_df.head())
print('角色顺序', custom_order)

# 保存表格
xls = pd.ExcelFile(filepath)
sheet_name = "出场表"
new_id = 1
while sheet_name in xls.sheet_names:
    sheet_name = f'出场表_{new_id}'
    new_id += 1

try:
    with pd.ExcelWriter(filepath, engine='openpyxl', mode='a') as writer:
        merged_df.to_excel(writer, sheet_name=sheet_name, index=False)
    print(f'表格已保存至{filepath}，表名为{sheet_name}')
except:
    postfix = filepath.split('.')[-1]
    outfile = filepath.replace(f'.{postfix}', f'out.{postfix}')
    with pd.ExcelWriter(outfile, engine='openpyxl') as writer:
        merged_df.to_excel(writer, sheet_name=sheet_name, index=False)
    print(f'表格已保存至{outfile}，表名为{sheet_name}')
message = f'表格已保存至{filepath}，表名为{sheet_name}'

input("按任意键退出...")