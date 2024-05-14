# 功能
该程序功能是按照角色表和正文表，导出出场表。其中角色的排序与角色表中排序一致。

# 使用
注：使用前需要关闭文件。如果未关闭文件则可能会写入失败，这时会输出到$filename$out.xlsx。

## 带托盘GUI版
1. 双击boomfish.exe，之后会常驻托盘。右键可莉图标可以退出。
2. 双击托盘中的可莉图标后启动新窗口，在新窗口中打开要处理的文件，显示处理成功后打开文件检查结果。

## 无托盘GUI版
1. 双击setup/boomfish_only.exe，在新窗口中打开要处理的文件，显示处理成功后打开文件检查结果。

## 黑框框版
1. 双击setup/boomfish_clean.exe，在新窗口中粘贴文件路径，显示处理成功后打开文件检查结果。

## python脚本版
### 命令行运行
```
python boomfish_clean.py -f filepath
```
### 双击运行
将*.py的默认打开程序设置为python，双击boomfish_clean.py后输入文件路径

# 开发安装
如需进行开发，请安装以下软件
1. 安装miniconda3
2. 安装依赖包
3. 安装VSCode
4. 安装VSCode插件
```
pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple
```

# TODO