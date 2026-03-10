import re
from pathlib import Path

def extract_attachments_from_tsv(tsv_path):
    """
    从 global_result.tsv 中提取所有带附件的邮件信息
    
    Args:
        tsv_path: global_result.tsv 的路径
        
    Returns:
        list: 包含附件信息的字典列表
    """
    
    # 读取 TSV 文件
    with open(tsv_path, 'r', encoding='utf-8') as f:
        lines = f.readlines()
    
    # 跳过前两行（OC注册信息 + 表头）
    data_lines = lines[2:]
    
    results = []
    
    # 获取项目根目录（用于构建绝对路径）
    project_root = Path(tsv_path).parent.parent
    
    for line in data_lines:
        cols = line.strip().split('\t')
        if len(cols) < 6:
            continue
        
        send_type = cols[0]
        date = cols[1]
        sender = cols[2]
        receiver = cols[3]
        letter_type = cols[4]
        directory = cols[5]
        
        # 构建邮件文件夹的绝对路径
        if directory.startswith('.\\'):
            rel_path = directory[2:]
        else:
            rel_path = directory
        
        email_dir = project_root / rel_path.replace('\\', '/')
        
        # 检查文件夹是否存在
        if not email_dir.exists():
            continue
        
        # 提取附件信息
        text_attachments = []
        file_attachments = []
        
        # 1. 检查 content.txt 中包含"附件"的行
        content_file = email_dir / "content.txt"
        if content_file.exists():
            try:
                with open(content_file, 'r', encoding='utf-8') as f:
                    for line_content in f:
                        line_stripped = line_content.strip()
                        # 如果这一行包含"附件"，就保存这一整行
                        if '附件' in line_stripped and line_stripped:
                            text_attachments.append(line_stripped)
            
            except Exception as e:
                print(f"读取 {content_file} 失败: {e}")
        
        # 2. 检查文件夹中的其他附件文件
        try:
            for file in email_dir.iterdir():
                if file.is_file() and file.name != 'content.txt':
                    # 保存相对于项目根目录的路径
                    try:
                        rel_file_path = file.relative_to(project_root)
                        file_attachments.append(str(rel_file_path).replace('\\', '/'))
                    except ValueError:
                        # 如果无法获取相对路径，使用绝对路径
                        file_attachments.append(str(file))
        except Exception as e:
            print(f"读取文件夹 {email_dir} 失败: {e}")
        
        # 如果有任何附件（文字或文件），记录这封邮件
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


def save_attachment_report(results, output_path):
    """保存附件报告为 TSV 文件"""
    with open(output_path, 'w', encoding='utf-8') as f:
        # 写入表头
        f.write("收/发类型\t发信日期\t发件人\t收件人\t信件类型\t信件下载位置\t文字附件数\t文件附件数\t文字附件详情\t文件附件路径\n")
        
        for item in results:
            # 用换行符分隔多个文字附件
            text_att_str = '\n'.join(item['text_attachments']) if item['text_attachments'] else ''
            file_att_str = ' | '.join(item['file_attachments']) if item['file_attachments'] else ''
            
            f.write(f"{item['send_type']}\t"
                   f"{item['date']}\t"
                   f"{item['sender']}\t"
                   f"{item['receiver']}\t"
                   f"{item['letter_type']}\t"
                   f"{item['directory']}\t"
                   f"{len(item['text_attachments'])}\t"
                   f"{len(item['file_attachments'])}\t"
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