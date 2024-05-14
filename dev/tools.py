import os, sys, json, ctypes, pystray

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

def read_file(file_name, encoding=None):
    file_=open(file_name, 'r', encoding=encoding)
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
