import os, sys, json, ctypes, pystray, chardet, re
import pandas as pd
from itertools import chain

real_path = os.path.realpath(sys.argv[0])
install_path = os.path.dirname(real_path)

def dialog(icon, item):
    MessageBox = ctypes.windll.user32.MessageBoxW
    MessageBox(None, 'Hello', 'Window title', 0)

def notify(icon: pystray.Icon, msg, title):
    icon.notify(msg, title)

def tojson(o, ensure_ascii=True):
    return json.dumps(o, default=lambda obj: obj.__dict__, sort_keys=True,ensure_ascii=ensure_ascii)

def toobj(strjson):
    json.loads(strjson)

def write_file(file_name, str_list, encoding=None):
    file_=open(file_name, 'w', encoding=encoding)
    file_.writelines(['%s\n'%s for s in str_list])
    file_.close()

def read_file(file_name):
    with open(file_name, 'rb') as f:
        result = chardet.detect(f.read())
        encoding = result['encoding']
    file_=open(file_name, 'r', encoding=encoding, errors='ignore')
    str_list = file_.read().splitlines()
    file_.close()
    return str_list

def load_json(filename, encoding=None):
    json_file=open(filename, 'r', encoding=encoding)
    json_strings=json_file.readlines()
    json_string=''.join(json_strings)
    json_file.close()
    return json.loads(json_string)

def save_json(filename, obj, ensure_ascii=True, encoding=None):
    str_json=tojson(obj, ensure_ascii)
    with open(filename, 'w', encoding=encoding) as f:
        f.write(str_json)
        f.close()

def save_to_xlsx(df, outfile, sheet_name, mode='a'):
    try:
        with pd.ExcelWriter(outfile, engine='openpyxl', mode=mode) as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    except:
        return f'保存失败：保存至 “{outfile}”，表名为 “{sheet_name}”，请关闭文件后重试'
    return f'表格已保存至 “{outfile}”，表名为 “{sheet_name}”'

def list_con(l):
    return list(chain(*l))

def split_text(text, char_to_split):
    cleaned_text = re.sub(str(char_to_split[0]) + r'\s+' + str(char_to_split[1]), char_to_split[:2], text)
    splited_text = cleaned_text.split(char_to_split[:2])
    splited_text = [splited_text[0]] + list_con([s.split(char_to_split[2], 1) for s in splited_text[1:]])
    return splited_text

def get_outfile(filepath):
    postfix = filepath.split('.')[-1]
    return filepath.replace(f'.{postfix}', 'out.xlsx')

def get_sheet_name(sheet_name, sheet_names):
    new_id = 1
    while sheet_name in sheet_names:
        sheet_name = f'出场表_{new_id}'
        new_id += 1
    return sheet_name

def get_resource_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)