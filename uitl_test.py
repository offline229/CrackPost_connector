from pathlib import Path
import re

# 文字附件在 TSV 中的编码：多条用 \x1e 分隔，条内换行/制表转义，避免撑破 TSV 行
TEXT_ATT_SEP = "\x1e"


def encode_text_attachments(atts):
    parts = []
    for a in atts or []:
        s = str(a).replace("\\", "\\\\").replace("\r\n", "\n").replace("\r", "\n")
        s = s.replace("\n", "\\n").replace("\t", "\\t")
        if s.strip():
            parts.append(s.strip())
    return TEXT_ATT_SEP.join(parts)


def decode_text_attachments(raw):
    """兼容：新编码(\\x1e + \\n)、旧版真实换行、以及「一行一条」。"""
    if raw is None:
        return []
    s = str(raw)
    if not s.strip():
        return []
    if TEXT_ATT_SEP in s:
        chunks = s.split(TEXT_ATT_SEP)
    elif "\\n" in s and s.count("\t") == 0:
        # 单字段内用 \\n 转义（且无真实 tab）
        chunks = [s]
    else:
        # 旧格式：真实换行分隔多条
        chunks = s.split("\n")

    out = []
    for c in chunks:
        c = c.replace("\\n", "\n").replace("\\t", "\t").replace("\\\\", "\\")
        c = c.strip()
        if c:
            out.append(c)
    return out


def extract_text_attachment_blocks(content_text):
    """从正文提取文字附件块，兼容单行/多行与多种括号写法。

    例：
      【附件：一支花】
      「附件：\\n……\\n」
      （附件：……）
      附件：一张化羽符。
      附件：   ← 单独一行，正文在后续行
      【附件：几张照片，里面有：\\nA\\nB】

    排除叙述性提及（如「那些附件，而不是你了」「打开附件，万一…」）：
      - 仅认「行首（可带括号）的附件标记」，不用「附件，」作触发（逗号太常见）
      - 「附件」左侧若紧贴 ≥2 个汉字/英文数字（非符号），视为正文连写，跳过
    """
    if not content_text:
        return []
    lines = content_text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    blocks = []
    i = 0
    # 真附件标记：【附件 / 「附件 / （附件 / 行首 附件：/附件: / 行首仅「附件」
    # 故意不含「附件，」——叙述里「那些附件，…」极常见
    start_re = re.compile(
        r"(?:^[【「（(]\s*附件|^附件\s*[:：]|^附件\s*$)"
    )
    word_char_re = re.compile(r"[\u4e00-\u9fffA-Za-z0-9]")

    def _open_close_ok(text):
        pairs = (("【", "】"), ("「", "」"), ("（", "）"), ("(", ")"))
        for a, b in pairs:
            if a in text:
                return text.count(a) <= text.count(b) and text.rstrip().endswith(b)
        # 无括号的「附件：…」单行视为完整
        return True

    def _left_glued_word_len(stripped):
        """附件左侧紧贴的汉字/英文数字个数（跳过开括号与空白）。"""
        m = re.search(r"附件", stripped)
        if not m:
            return 0
        before = stripped[: m.start()]
        before = re.sub(r"^[【「（(\s]+", "", before)
        n = 0
        for ch in reversed(before):
            if word_char_re.match(ch):
                n += 1
            else:
                break
        return n

    while i < len(lines):
        raw = lines[i]
        stripped = raw.strip()
        if not stripped or "附件" not in stripped or not start_re.search(stripped):
            i += 1
            continue

        # 同属正文叙述、不是附件登记：跳过「没有附件」类
        if re.search(r"(没有|无|不含|打开|拿到|做成|一种).{0,4}附件", stripped):
            i += 1
            continue
        # 左侧紧贴 ≥2 个非符号字符（那些附件 / 打开附件）→ 叙述，非登记
        if _left_glued_word_len(stripped) >= 2:
            i += 1
            continue

        buf = [stripped]
        # 已在同行闭合 → 单块
        if _open_close_ok(stripped) and not (
            stripped.endswith(("：", ":", "，", ",", "有：", "有:"))
            or (stripped.startswith(("【", "「", "（", "(")) and not stripped.endswith(("】", "」", "）", ")")))
        ):
            blocks.append(stripped)
            i += 1
            continue

        # 多行块：括号未闭合时忽略空行（地缝信常见中间空多行再写】）
        i += 1
        while i < len(lines):
            nxt = lines[i]
            ns = nxt.strip()
            joined = "\n".join(buf)
            if not ns:
                if _open_close_ok(joined):
                    break
                # 未闭合：保留空行（折叠为最多一个），继续往后找】
                if buf and buf[-1] != "":
                    buf.append("")
                i += 1
                continue
            if re.match(r"^(主题|发件邮箱|收件邮箱|日期|发件人|发件者|--------)", ns):
                break
            # 新的附件块开头 → 当前块结束（即使未闭合也先收口）
            if len(buf) > 0 and start_re.search(ns) and re.match(r"^[【「（(]?附件", ns):
                break
            buf.append(ns)
            joined = "\n".join(buf)
            if _open_close_ok(joined):
                i += 1
                break
            i += 1

        block = "\n".join(buf).strip()
        # 去掉块尾为对齐留下的空行
        block = re.sub(r"\n{3,}", "\n\n", block).strip()
        if block:
            blocks.append(block)
        else:
            i += 1

    # 去重保序
    seen = set()
    uniq = []
    for b in blocks:
        if b not in seen:
            seen.add(b)
            uniq.append(b)
    return uniq


def _resolve_email_dir(directory, tsv_path):
    """把 global_result 里的信件路径解析成实际文件夹（兼容新旧目录名）。"""
    if not directory:
        return None
    raw = str(directory).strip().strip('"')
    if not raw:
        return None

    email_root = Path(tsv_path).resolve().parent          # .../email
    project_root = email_root.parent                      # 项目根

    p = Path(raw)
    if p.is_absolute() and p.exists():
        return p

    rel = raw[2:] if raw.startswith(".\\") or raw.startswith("./") else raw
    rel = rel.replace("\\", "/")
    parts = Path(rel).parts

    candidates = [
        project_root / rel,
        email_root / rel,
    ]
    # .\email\发\xxx → email_root/发/xxx
    if parts and parts[0].lower() == "email":
        candidates.append(email_root.joinpath(*parts[1:]))
    # .\test\email\收\xxx（或任意前缀\email\收\xxx）→ 从 email 段之后接到 email_root
    if "email" in [p.lower() for p in parts]:
        for i, p in enumerate(parts):
            if p.lower() == "email" and i + 1 < len(parts):
                candidates.append(email_root.joinpath(*parts[i + 1:]))
                break
    # 仅给了 发\xxx 或 收\xxx
    if parts and parts[0] in ("发", "收"):
        candidates.append(email_root.joinpath(*parts))
    # 最后一段目录名在 email_root/收|发 下模糊匹配
    if parts:
        leaf = parts[-1]
        for side in ("收", "发"):
            candidates.append(email_root / side / leaf)

    seen = set()
    for c in candidates:
        try:
            c = c.resolve()
        except Exception:
            continue
        key = str(c).lower()
        if key in seen:
            continue
        seen.add(key)
        if c.exists() and c.is_dir():
            return c
    return None


def extract_attachments_from_tsv(tsv_path):
    """
    从 global_result.tsv 中提取所有带附件的邮件信息
    
    Args:
        tsv_path: global_result.tsv 的路径
        
    Returns:
        list: 包含附件信息的字典列表
    """
    tsv_path = Path(tsv_path)
    with open(tsv_path, 'r', encoding='utf-8') as f:
        lines = f.readlines()

    # 兼容：oc_registerdata / 可选 oc_aliases / 表头 / 数据
    data_start = 0
    for i, line in enumerate(lines):
        if line.startswith("收/发类型") or line.startswith("收/发\t"):
            data_start = i + 1
            break
    else:
        # 旧假设：前两行是 OC + 表头
        data_start = 2 if len(lines) >= 2 else 0

    results = []
    project_root = tsv_path.resolve().parent.parent

    for line in lines[data_start:]:
        cols = line.rstrip("\n").split('\t')
        if len(cols) < 6:
            continue

        send_type = cols[0]
        date = cols[1]
        sender = cols[2]
        receiver = cols[3]
        letter_type = cols[4]
        directory = cols[5]

        email_dir = _resolve_email_dir(directory, tsv_path)
        if not email_dir:
            continue

        text_attachments = []
        file_attachments = []

        text_attachments, file_attachments = scan_letter_attachments(
            email_dir, project_root=project_root
        )

        if text_attachments or file_attachments:
            results.append({
                'send_type': send_type,
                'date': date,
                'sender': sender,
                'receiver': receiver,
                'letter_type': letter_type,
                'directory': directory,
                'text_attachments': text_attachments,
                'file_attachments': file_attachments
            })

    return results


def print_attachment_summary(results):
    """打印附件统计摘要"""
    print(f"\n{'='*80}")
    print(f"附件统计摘要")
    print(f"{'='*80}\n")
    
    print(f"总共找到 {len(results)} 封带附件的邮件\n")
    
    # 统计
    text_only = len([r for r in results if r['text_attachments'] and not r['file_attachments']])
    file_only = len([r for r in results if r['file_attachments'] and not r['text_attachments']])
    both = len([r for r in results if r['text_attachments'] and r['file_attachments']])
    
    print(f"- 仅有文字附件: {text_only} 封")
    print(f"- 仅有文件附件: {file_only} 封")
    print(f"- 同时有文字和文件附件: {both} 封\n")
    
    # 详细列表
    print(f"{'='*80}")
    print(f"详细列表")
    print(f"{'='*80}\n")
    
    for i, item in enumerate(results, 1):
        print(f"[{i}] {item['send_type']} | {item['date']} | {item['sender']} → {item['receiver']} | {item['letter_type']}")
        print(f"    路径: {item['directory']}")
        
        if item['text_attachments']:
            print(f"    文字附件 ({len(item['text_attachments'])} 行):")
            for j, att in enumerate(item['text_attachments'], 1):
                # 限制显示长度
                display_text = att[:100] + '...' if len(att) > 100 else att
                print(f"      [{j}] {display_text}")
        
        if item['file_attachments']:
            print(f"    文件附件 ({len(item['file_attachments'])} 个):")
            for j, att in enumerate(item['file_attachments'], 1):
                print(f"      [{j}] {att}")
        
        print()


def normalize_letter_path_key(directory):
    """统一信件路径键，便于附件报告按信合并/替换。

    一律压成 ``email\\收|发\\信夹``（小写），兼容：
    ``.\\email\\收\\…``、``email/收/…``、``D:\\…\\email\\收\\…``。
    """
    s = str(directory or "").strip().strip('"').replace("/", "\\")
    while s.startswith(".\\"):
        s = s[2:]
    s = s.lower()
    marker = "\\email\\"
    i = s.find(marker)
    if i >= 0:
        return s[i + 1 :]  # email\收\…
    if s.startswith("email\\"):
        return s
    return s


def scan_letter_attachments(email_dir, project_root=None):
    """扫描单个信夹：正文文字附件块（支持多行）+ 非 content.txt 文件。"""
    email_dir = Path(email_dir)
    text_attachments = []
    file_attachments = []
    content_file = email_dir / "content.txt"
    if content_file.exists():
        try:
            content = content_file.read_text(encoding="utf-8")
            text_attachments = extract_text_attachment_blocks(content)
        except Exception as e:
            print(f"读取 {content_file} 失败: {e}")
    if project_root is None:
        project_root = email_dir
        # 尝试上溯到含 email 的项目根
        for p in [email_dir.parent, email_dir.parent.parent, email_dir.parent.parent.parent]:
            if p.name.lower() == "email" or (p / "email").exists():
                project_root = p.parent if p.name.lower() == "email" else p
                break
    project_root = Path(project_root).resolve()
    try:
        for file in sorted(email_dir.iterdir(), key=lambda p: p.name.lower()):
            if file.is_file() and file.name != "content.txt":
                # 跳过明显非附件的系统垃圾
                if file.name.lower() in ("thumbs.db", "desktop.ini", ".ds_store"):
                    continue
                try:
                    rel_file_path = file.resolve().relative_to(project_root)
                    file_attachments.append(str(rel_file_path).replace("\\", "/"))
                except ValueError:
                    file_attachments.append(str(file))
    except Exception as e:
        print(f"读取文件夹 {email_dir} 失败: {e}")
    return text_attachments, file_attachments


def build_attachment_item_from_meta(meta, email_dir, project_root=None):
    """用库存元数据 + 信夹扫描，构造一条 attachment_report 记录。"""
    text_attachments, file_attachments = scan_letter_attachments(email_dir, project_root=project_root)
    return {
        "send_type": meta.get("send_type") or meta.get("direction") or "",
        "date": meta.get("date") or "",
        "sender": meta.get("sender") or "",
        "receiver": meta.get("receiver") or "",
        "letter_type": meta.get("letter_type") or "",
        "directory": meta.get("directory") or meta.get("path") or "",
        "text_attachments": text_attachments,
        "file_attachments": file_attachments,
    }


def load_attachment_report(report_path):
    """读取已有 attachment_report.tsv（兼容文字附件多行）。"""
    report_path = Path(report_path)
    if not report_path.exists():
        return []
    data = report_path.read_text(encoding="utf-8")
    raw_lines = data.split("\n")
    valid_lines = []
    current = ""
    for line in raw_lines:
        if not current:
            current = line
            continue
        tab_count = current.count("\t")
        if tab_count < 9:
            current += "\n" + line
        else:
            valid_lines.append(current)
            current = line
    if current:
        valid_lines.append(current)
    if not valid_lines:
        return []
    results = []
    for line in valid_lines[1:]:
        cols = line.split("\t")
        if len(cols) < 6:
            continue
        text_raw = cols[8] if len(cols) > 8 else ""
        file_raw = cols[9] if len(cols) > 9 else ""
        text_attachments = decode_text_attachments(text_raw)
        file_attachments = [x.strip() for x in file_raw.split("|") if x.strip()] if file_raw.strip() else []
        results.append({
            "send_type": cols[0].strip(),
            "date": cols[1].strip(),
            "sender": cols[2].strip(),
            "receiver": cols[3].strip(),
            "letter_type": cols[4].strip(),
            "directory": cols[5].strip(),
            "text_attachments": text_attachments,
            "file_attachments": file_attachments,
        })
    return results


def upsert_attachment_report(report_path, items, remove_empty=True):
    """按信件路径键合并写入 attachment_report.tsv（不触碰其他信件行）。

    items: 一条或多条附件记录。若某条无文字/文件附件且 remove_empty，则删除该路径旧行。
    """
    report_path = Path(report_path)
    report_path.parent.mkdir(parents=True, exist_ok=True)
    by_key = {normalize_letter_path_key(r["directory"]): r for r in load_attachment_report(report_path)}
    if not isinstance(items, (list, tuple)):
        items = [items]
    for item in items:
        key = normalize_letter_path_key(item.get("directory"))
        has = bool(item.get("text_attachments") or item.get("file_attachments"))
        if has:
            by_key[key] = item
        elif remove_empty and key in by_key:
            del by_key[key]
    save_attachment_report(list(by_key.values()), report_path)
    return report_path


def attachment_paths_index(report_path):
    """directory_key -> 是否有附件。"""
    idx = {}
    for r in load_attachment_report(report_path):
        key = normalize_letter_path_key(r.get("directory"))
        idx[key] = bool(r.get("text_attachments") or r.get("file_attachments"))
    return idx


def save_attachment_report(results, output_path):
    """保存附件报告为 TSV 文件"""
    with open(output_path, 'w', encoding='utf-8') as f:
        # 写入表头
        f.write("收/发类型\t发信日期\t发件人\t收件人\t信件类型\t信件下载位置\t文字附件数\t文件附件数\t文字附件详情\t文件附件路径\n")
        
        for item in results:
            # 编码后单行写入，避免真实换行撑破 TSV；可视化侧再还原
            text_att_str = encode_text_attachments(item.get('text_attachments') or [])
            file_att_str = ' | '.join(item.get('file_attachments') or [])
            
            f.write(f"{item['send_type']}\t"
                   f"{item['date']}\t"
                   f"{item['sender']}\t"
                   f"{item['receiver']}\t"
                   f"{item['letter_type']}\t"
                   f"{item['directory']}\t"
                   f"{len(item.get('text_attachments') or [])}\t"
                   f"{len(item.get('file_attachments') or [])}\t"
                   f"{text_att_str}\t"
                   f"{file_att_str}\n")
    
    print(f"\n附件报告已保存到: {output_path}")


# ========== 测试代码 ==========
if __name__ == "__main__":
    # 设置路径
    tsv_path = Path(__file__).parent / "email" / "global_result.tsv"
    output_path = Path(__file__).parent / "email" / "attachment_report.tsv"
    
    # 提取附件
    print("正在提取附件信息...")
    results = extract_attachments_from_tsv(tsv_path)
    
    # 打印摘要
    print_attachment_summary(results)
    
    # 保存报告
    save_attachment_report(results, output_path)
    
    # 额外测试：打印前3封邮件的详细信息
    print(f"\n{'='*80}")
    print("前3封带附件邮件的详细信息")
    print(f"{'='*80}\n")
    
    for item in results[:3]:
        print(f"发信日期: {item['date']}")
        print(f"发件人: {item['sender']}")
        print(f"收件人: {item['receiver']}")
        print(f"信件类型: {item['letter_type']}")
        print(f"文字附件行:")
        for i, line in enumerate(item['text_attachments'], 1):
            print(f"  [{i}] {line}")
        print(f"文件附件: {item['file_attachments']}")
        print(f"{'-'*80}\n")