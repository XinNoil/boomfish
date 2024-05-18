import re

# 从用户那里获取正则表达式模式
# pattern_input = input("请输入正则表达式模式: ")
pattern_input = r"^##(.*)第\s*(\d+)\s*章\s*(.*)$"

# 编译正则表达式模式
pattern = re.compile(pattern_input)

# 定义待匹配的字符串
test_string = "##这是一段测试第 123 章 内容"

# 使用模式进行匹配
matches = pattern.findall(test_string)

# 检查是否匹配成功
if matches:
    for match in matches:
        if isinstance(match, tuple):
            # 当匹配结果是一个元组时，表示有多个捕获组
            part1, chapter_number, part3 = match
            print(f"匹配成功!")
            print(f"部分1: {part1}")
            print(f"章节号: {chapter_number}")
            print(f"部分3: {part3}")
        else:
            # 当匹配结果不是元组时，表示没有捕获组
            print(f"匹配成功: {match}")
else:
    print("没有匹配到符合的模式。")
