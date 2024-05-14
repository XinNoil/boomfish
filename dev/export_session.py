import os, argparse
import pandas as pd
from tools import get_sheet_name, save_to_xlsx

def export_session(filepath, sheet_name = "出场表", chekcbox_var=None):
    # 读入表格
    text_df = pd.read_excel(filepath, sheet_name="正文表")[['集数', '标题', '角色', '正文']] # 集数，标题，角色，正文
    characters_df = pd.read_excel(filepath, sheet_name="角色表")[['角色', '主播', '性别', '年龄', '人设']] # 角色，主播，性别，年龄，人设
    text_df.dropna(how='all', inplace=True)
    text_df.fillna(method='ffill', inplace=True)
    text_count_df = text_df.groupby(['集数', '标题', '角色']).count().reset_index()[['集数', '标题', '角色', '正文']]

    # 合并表格
    merged_df = pd.merge(characters_df, text_count_df, on='角色')
    merged_df.rename(columns={'正文': '句数'}, inplace=True)
    merged_df = merged_df.reindex(columns=['集数', '标题', '角色', '主播', '性别', '年龄', '人设', '句数'])

    # 自定义排序序列
    custom_order = dict(zip(characters_df['角色'], range(len(characters_df['角色']))))
    merged_df['角色'] = pd.Categorical(merged_df['角色'], categories=custom_order, ordered=True)
    merged_df = merged_df.sort_values(by=['集数', '角色'])

    # 保存表格
    xls = pd.ExcelFile(filepath)
    sheet_name = get_sheet_name(sheet_name, xls.sheet_names)
    return save_to_xlsx(merged_df, filepath, sheet_name)

if __name__ == '__main__':
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
    message = export_session(filepath)
    print(message)
    input("按任意键退出...")