def extract_letter_type(subject):
    """
    从主题字符串中提取信件类型（如A123、B456、C789、A*123、A+B456等）。
    自动去除类型中的空格。
    """
    import re
    # 去除类型中的空格（只针对类型部分）
    subject_nospace = re.sub(r'(Fw[:：]\s*)(A\s*\*?\s*\d+|B\s*\d+|C\s*\d+|A\s*\+\s*B\s*\d+|A\s*\+\s*C\s*\d+)',
                             lambda m: m.group(1) + re.sub(r'\s+', '', m.group(2)),
                             subject, flags=re.IGNORECASE)
    subject_nospace = re.sub(r'^(A\s*\*?\s*\d+|B\s*\d+|C\s*\d+|A\s*\+\s*B\s*\d+|A\s*\+\s*C\s*\d+)',
                             lambda m: re.sub(r'\s+', '', m.group(1)),
                             subject_nospace, flags=re.IGNORECASE)
    # 正则提取类型
    m = re.search(r'Fw[:：]\s*(A\d+|A\*\d+|B\d+|C\d+|A\+B\d+|A\+C\d+)', subject_nospace, re.IGNORECASE)
    if m:
        return m.group(1)
    m2 = re.match(r'(A|B|C.*|A\+B.*|A\+C.*)', subject_nospace)
    if m2:
        return m2.group(1)
    return ""

# 测试用例
test_subjects = [
    "FW:C 2396 来自尹浦的信",
    "Fw:B 456_来自李四的信",
    "Fw:C 789_来自王五的信",
    "Fw:A* 321_来自赵六的信",
    "Fw:A+B 654_来自钱七的信",
    "Fw:A+C 987_来自孙八的信",
    "A 123_普通邮件",
    "B 456_普通邮件",
    "C 789_普通邮件",
    "A* 321_普通邮件",
    "A+B 654_普通邮件",
    "A+C 987_普通邮件",
    "Fw:A123_无空格",
    "Fw:B456_无空格",
    "Fw:C789_无空格",
    "Fw:A*123_无空格",
    "Fw:A+B654_无空格",
    "Fw:A+C987_无空格",
]

for subj in test_subjects:
    print(f"{subj} -> {extract_letter_type(subj)}")