# 功能
该程序功能有：
1. 养鱼：选择多个小说txt文件批量导出正文白表。对非对话部分自动标记为旁白。
2. 炸鱼：选择多个演出表格按照角色表和正文表，批量导出出场表。其中角色的排序与角色表中排序一致。

# 使用
注：使用前需要关闭文件。如果未关闭文件则可能会写入失败，这时会输出到$filename$out.xlsx。

## GUI版
1. 双击boomfish.exe，之后会常驻托盘。
2. 左键单击托盘中的可莉图标后启动新炸鱼窗口，右键托盘显示菜单。
3. 在新窗口中打开要处理的文件，显示处理成功后打开文件检查结果

## python脚本版
### 命令行运行
```
python boomfish.py
python segment_text.py -f filepath
python export_session.py -f filepath
```
### 双击运行
将*.py的默认打开程序设置为python，双击boomfish.py后输入文件路径

# 开发安装
如需进行开发，请安装以下软件
1. 安装miniconda3
2. 安装依赖包
3. 安装VSCode
4. 安装VSCode插件
```
pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple
```
