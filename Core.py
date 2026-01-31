#!/usr/bin/env python3
"""
Macro Cleaner - 清除Office文档中的所有VBA宏
支持格式：.docm, .dotm, .xlsm, .xltm, .pptm, .potm
"""

import html as html_module
import hashlib
import os
import shutil
import sys
import zipfile
import tempfile
import re
from pathlib import Path
from datetime import datetime
from typing import List, Dict

# 打包无控制台时 stdout/stderr 可能为 None，print/flush 会报错，重定向到 devnull
if sys.stdout is None:
    sys.stdout = open(os.devnull, "w")
if sys.stderr is None:
    sys.stderr = open(os.devnull, "w")

# 报告与界面文案（中/英）
TEXTS_ZH = {
    "status_success": "成功",
    "status_fail": "失败",
    "msg_file_not_found": "文件不存在",
    "msg_unsupported_format": "不支持的格式: {suffix}",
    "msg_invalid_format": "无效的文件格式",
    "msg_unknown_doc_type": "无法识别文档类型",
    "msg_no_macro_skip": "文件未包含宏，无需处理",
    "msg_no_macro_format": "该格式不支持宏，视为无宏",
    "msg_save_failed": "文件保存失败",
    "msg_macro_cleared_replace": "宏已清除，原文件已替换，备份已创建",
    "msg_macro_cleared": "宏已清除({size} bytes)",
    "msg_no_macro": "文件未包含宏",
    "msg_exception": "处理异常: {e}",
    "report_title": "宏清除报告",
    "generated_at": "生成时间",
    "stat_ok": "处理成功",
    "stat_err": "处理失败",
    "stat_mac": "发现宏",
    "stat_sz": "清除大小",
    "section_title": "处理详情",
    "filter_status": "状态",
    "filter_macro": "宏",
    "filter_name": "文件名",
    "filter_all": "全部",
    "filter_success": "成功",
    "filter_fail": "失败",
    "filter_has_macro": "有宏",
    "filter_no_macro": "无宏",
    "placeholder_keyword": "输入关键词",
    "th_filename": "文件名",
    "th_orig_size": "原始大小",
    "th_macro_size": "宏大小",
    "th_status": "状态",
    "th_result": "处理结果",
    "th_time": "时间",
    "th_action": "操作",
    "btn_detail": "详情",
    "detail_orig_file": "原文件",
    "detail_output_file": "去除后文件",
    "detail_hash": "hash值",
    "detail_created": "创建日期",
    "detail_modified": "修改日期",
    "detail_macro": "宏",
    "footer": "Macro Cleaner 报告",
}
TEXTS_EN = {
    "status_success": "Success",
    "status_fail": "Failed",
    "msg_file_not_found": "File not found",
    "msg_unsupported_format": "Unsupported format: {suffix}",
    "msg_invalid_format": "Invalid file format",
    "msg_unknown_doc_type": "Unknown document type",
    "msg_no_macro_skip": "No macro in file, skipped",
    "msg_no_macro_format": "Format does not support macros, treated as no macro",
    "msg_save_failed": "Failed to save file",
    "msg_macro_cleared_replace": "Macro cleared, original replaced, backup created",
    "msg_macro_cleared": "Macro cleared ({size} bytes)",
    "msg_no_macro": "No macro in file",
    "msg_exception": "Error: {e}",
    "report_title": "Macro Cleaner Report",
    "generated_at": "Generated",
    "stat_ok": "Success",
    "stat_err": "Failed",
    "stat_mac": "Macro found",
    "stat_sz": "Cleared size",
    "section_title": "Details",
    "filter_status": "Status",
    "filter_macro": "Macro",
    "filter_name": "File name",
    "filter_all": "All",
    "filter_success": "Success",
    "filter_fail": "Failed",
    "filter_has_macro": "Has macro",
    "filter_no_macro": "No macro",
    "placeholder_keyword": "Keyword",
    "th_filename": "File name",
    "th_orig_size": "Original size",
    "th_macro_size": "Macro size",
    "th_status": "Status",
    "th_result": "Result",
    "th_time": "Time",
    "th_action": "Action",
    "btn_detail": "Details",
    "detail_orig_file": "Original file",
    "detail_output_file": "Output file",
    "detail_hash": "Hash",
    "detail_created": "Created",
    "detail_modified": "Modified",
    "detail_macro": "Macro",
    "footer": "Macro Cleaner Report",
}


# OpenXML 中 vbaProject.bin 在 zip 内的路径
_VBA_BIN_PATHS = ("word/vbaProject.bin", "xl/vbaProject.bin", "ppt/vbaProject.bin")
def _extract_vba_from_parser(parser) -> list:
    """从已打开的 VBA_Parser 中收集所有宏代码，返回字符串列表。"""
    out = []
    for (_, _, _, vba_code) in parser.extract_macros():
        if vba_code:
            if isinstance(vba_code, bytes):
                out.append(vba_code.decode("latin-1", errors="replace"))
            else:
                out.append(str(vba_code))
    if not out and hasattr(parser, "get_vba_code_all_modules"):
        try:
            all_code = parser.get_vba_code_all_modules()
            if all_code:
                out.append(all_code if isinstance(all_code, str) else all_code.decode("latin-1", errors="replace"))
        except Exception:
            pass
    if not out and getattr(parser, "vba_projects", None):
        for project in parser.vba_projects:
            for module in getattr(project, "modules", []):
                code = getattr(module, "code_str", None) or getattr(module, "code", None)
                if code:
                    out.append(code if isinstance(code, str) else code.decode("latin-1", errors="replace"))
    return out
def _extract_vba_code(file_path: str) -> str:
    """从 Office 文件中提取 VBA 宏源代码。依赖 oletools，未安装时返回提示。"""
    try:
        from oletools.olevba import VBA_Parser
    except ImportError:
        return "(请安装 oletools 以查看宏代码: pip install oletools)"
    out = []
    parser = None
    temp_bin = None
    try:
        # 1) 先对完整 Office 文件用 VBA_Parser
        parser = VBA_Parser(file_path)
        if hasattr(parser, "detect_vba_macros"):
            try:
                parser.detect_vba_macros()
            except Exception:
                pass
        out = _extract_vba_from_parser(parser)
        if parser is not None and hasattr(parser, "close"):
            try:
                parser.close()
            except Exception:
                pass
        parser = None

        # 2) 若仍无结果，从 OpenXML 中解出 vbaProject.bin 再解析（OLE 流）
        if not out:
            path = Path(file_path).resolve()
            if not path.exists():
                return "\n\n---\n\n".join(out) if out else "(无可见宏代码)"
            try:
                with zipfile.ZipFile(path, "r") as zf:
                    vba_bin_path = None
                    for name in _VBA_BIN_PATHS:
                        if name in zf.namelist():
                            vba_bin_path = name
                            break
                    if vba_bin_path is None:
                        return "\n\n---\n\n".join(out) if out else "(无可见宏代码)"
                    fd, temp_bin = tempfile.mkstemp(suffix=".bin")
                    try:
                        os.close(fd)
                        with open(temp_bin, "wb") as f:
                            f.write(zf.read(vba_bin_path))
                        parser = VBA_Parser(temp_bin)
                        out = _extract_vba_from_parser(parser)
                    finally:
                        if parser and hasattr(parser, "close"):
                            try:
                                parser.close()
                            except Exception:
                                pass
                        if temp_bin and os.path.exists(temp_bin):
                            try:
                                os.unlink(temp_bin)
                            except Exception:
                                pass
            except zipfile.BadZipFile:
                pass
            except Exception:
                pass
    except Exception as e:
        return f"(提取失败: {e})"
    finally:
        if parser is not None and hasattr(parser, "close"):
            try:
                parser.close()
            except Exception:
                pass
        if temp_bin and os.path.exists(temp_bin):
            try:
                os.unlink(temp_bin)
            except Exception:
                pass
    return "\n\n---\n\n".join(out) if out else "(无可见宏代码)"


_last_vba_size: int = 0


def get_last_vba_size() -> int:
    """返回最近一次 clean_vba_macro 清除的宏大小（字节），未清除则为 0。"""
    return _last_vba_size


def clean_vba_macro(file_path: str, replace_original: bool = False, generate_report: bool = False, is_english: bool = False) -> int:
    global _last_vba_size
    _last_vba_size = 0
    t = TEXTS_EN if is_english else TEXTS_ZH
    input_path = Path(file_path).resolve()
    temp_dir = None
    # 支持的扩展名（含宏格式）；无宏格式单独处理
    SUPPORTED_EXTS = ['.docm', '.dotm', '.xlsm', '.xltm', '.pptm', '.potm']
    # 不支持宏的文档格式，自动视为无宏
    NO_MACRO_EXTS = [
        # Microsoft Excel
        '.xlsx',  # Excel 工作簿
        '.xlsb',  # Excel 二进制工作簿（无宏）
        '.xltx',  # Excel 模板（无宏）
        '.xls',  # Excel 97-2003 工作簿（旧格式，无 VBA 存储结构）

        # Microsoft Word
        '.docx',  # Word 文档
        '.dotx',  # Word 模板（无宏）
        '.doc',  # Word 97-2003 文档（旧格式）

        # Microsoft PowerPoint
        '.pptx',  # PowerPoint 演示文稿
        '.potx',  # PowerPoint 模板（无宏）
        '.ppsx',  # PowerPoint 放映（无宏）
        '.ppt',  # PowerPoint 97-2003 演示文稿（旧格式）

        # 其他 Office 格式
        '.odt',  # OpenDocument 文本
        '.ods',  # OpenDocument 电子表格
        '.odp',  # OpenDocument 演示文稿
        '.rtf',  # 富文本格式
        '.pdf',  # PDF 文档
        '.txt',  # 纯文本
        '.csv',  # 逗号分隔值
        '.xml',  # XML 数据
    ]
    # 报告数据收集
    report_data = {
        'file_name': input_path.name,
        'file_path': str(input_path),
        'file_size': 0,
        'status': t['status_fail'],
        'vba_found': False,
        'vba_size': 0,
        'vba_code': '',
        'output_path': '',
        'message': '',
        'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        'original_md5': '',
        'original_ctime': '',
        'original_mtime': '',
        'output_md5': '',
        'output_ctime': '',
        'output_mtime': '',
        'is_english': is_english,
    }
    try:
        # 前置检查
        if not input_path.exists():
            report_data['message'] = t['msg_file_not_found']
            _save_report([report_data], generate_report)
            return 0

        suffix_lower = input_path.suffix.lower()
        # .xlsx / .docx / .pptx 等不支持宏的格式：直接视为无宏，不报错
        if suffix_lower in NO_MACRO_EXTS:
            report_data['file_size'] = input_path.stat().st_size
            orig_meta = _get_file_meta(str(input_path))
            report_data['original_md5'] = orig_meta['md5']
            report_data['original_ctime'] = orig_meta['ctime']
            report_data['original_mtime'] = orig_meta['mtime']
            report_data['status'] = t['status_success']
            report_data['message'] = t['msg_no_macro_format']
            report_data['output_path'] = str(input_path)
            report_data['output_md5'] = orig_meta['md5']
            report_data['output_ctime'] = orig_meta['ctime']
            report_data['output_mtime'] = orig_meta['mtime']
            _save_report([report_data], generate_report)
            return 0

        if suffix_lower not in SUPPORTED_EXTS:
            report_data['message'] = t['msg_unsupported_format'].format(suffix=input_path.suffix)
            _save_report([report_data], generate_report)
            return 0

        report_data['file_size'] = input_path.stat().st_size
        orig_meta = _get_file_meta(str(input_path))
        report_data['original_md5'] = orig_meta['md5']
        report_data['original_ctime'] = orig_meta['ctime']
        report_data['original_mtime'] = orig_meta['mtime']

        # 创建临时目录
        temp_dir = tempfile.mkdtemp(prefix="macro_clean_")
        extract_path = Path(temp_dir) / "extracted"

        # 1. 解压Office文档
        try:
            with zipfile.ZipFile(input_path, 'r') as zip_ref:
                zip_ref.extractall(extract_path)
        except zipfile.BadZipFile:
            report_data['message'] = t['msg_invalid_format']
            _save_report([report_data], generate_report)
            return 0

        # 2. 检测文档类型和VBA路径
        vba_paths = {
            'word': 'word/vbaProject.bin',
            'excel': 'xl/vbaProject.bin',
            'ppt': 'ppt/vbaProject.bin'
        }

        doc_type = None
        if (extract_path / "word").exists():
            doc_type = "word"
        elif (extract_path / "xl").exists():
            doc_type = "excel"
        elif (extract_path / "ppt").exists():
            doc_type = "ppt"

        if not doc_type:
            report_data['message'] = t['msg_unknown_doc_type']
            _save_report([report_data], generate_report)
            return 0

        # 3. 删除VBA组件
        vba_path = extract_path / vba_paths[doc_type]
        vba_found = False

        if vba_path.exists():
            vba_found = True
            vba_size = vba_path.stat().st_size
            report_data['vba_code'] = _extract_vba_code(str(input_path))
            vba_path.unlink()

            # 删除关系文件
            rels_path = vba_path.parent / (vba_path.name + ".rels")
            if rels_path.exists():
                rels_path.unlink()

            # if is_english:
            #     print(f"Macro cleared: {vba_paths[doc_type]} from: {input_path.name}")
            # else:
            #     print(f"已清除宏组件: {vba_paths[doc_type]}  来自:{input_path.name}文件")
            report_data['vba_found'] = True
            report_data['vba_size'] = vba_size
            _last_vba_size = vba_size

        # 4. 无宏则无需处理，不生成副本、不覆盖原文件
        if not vba_found:
            report_data['status'] = t['status_success']
            report_data['message'] = t['msg_no_macro_skip']
            report_data['output_path'] = str(input_path)
            out_meta = _get_file_meta(str(input_path))
            report_data['output_md5'] = out_meta['md5']
            report_data['output_ctime'] = out_meta['ctime']
            report_data['output_mtime'] = out_meta['mtime']
            _save_report([report_data], generate_report)
            return 0

        # 5. 清理[Content_Types].xml（仅在有宏时）
        content_types_path = extract_path / "[Content_Types].xml"
        if content_types_path.exists():
            _clean_content_types(content_types_path)

        if replace_original:
            # 替换原文件：先备份，再覆盖
            backup_path = input_path.parent / f"{input_path.stem}_backup{input_path.suffix}"
            shutil.copy2(input_path, backup_path)
        output_path = input_path

        # 6. 重新打包
        _repack_docx(extract_path, output_path)

        # 7. 验证输出
        if not output_path.exists():
            report_data['message'] = t['msg_save_failed']
            _save_report([report_data], generate_report)
            return 0

        # 8. 更新报告数据
        report_data['status'] = t['status_success']
        report_data['output_path'] = str(output_path)
        out_meta = _get_file_meta(str(output_path))
        report_data['output_md5'] = out_meta['md5']
        report_data['output_ctime'] = out_meta['ctime']
        report_data['output_mtime'] = out_meta['mtime']

        if replace_original and vba_found:
            report_data['message'] = t['msg_macro_cleared_replace']
            _save_report([report_data], generate_report)
            return 1
        elif vba_found:
            report_data['message'] = t['msg_macro_cleared'].format(size=report_data['vba_size'])
            _save_report([report_data], generate_report)
            return 1
        else:
            report_data['message'] = t['msg_no_macro']
            _save_report([report_data], generate_report)
            return 0
    except Exception as e:
        report_data['message'] = t['msg_exception'].format(e=str(e))
        _save_report([report_data], generate_report)
        return 0

    finally:
        # 清理临时目录
        if temp_dir and os.path.exists(temp_dir):
            shutil.rmtree(temp_dir, ignore_errors=True)


def _clean_content_types(content_types_path: Path):
    """清理Content_Types.xml中的宏声明"""
    try:
        content = content_types_path.read_text(encoding='utf-8')

        # 移除vbaProject相关内容类型
        patterns = [
            r'<Override PartName="/word/vbaProject\.bin".*?/>',
            r'<Override PartName="/xl/vbaProject\.bin".*?/>',
            r'<Override PartName="/ppt/vbaProject\.bin".*?/>',
            r'<Default Extension="bin".*?/>',
        ]

        for pattern in patterns:
            content = re.sub(pattern, '', content, flags=re.DOTALL)

        # 清理空行
        content = re.sub(r'\n\s*\n', '\n', content)

        content_types_path.write_text(content, encoding='utf-8')

    except Exception as e:
        pass
def _repack_docx(source_dir: Path, output_path: Path):
    """重新打包为Office文档"""
    output_path.parent.mkdir(parents=True, exist_ok=True)

    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for file_path in source_dir.rglob('*'):
            if file_path.is_file():
                arcname = file_path.relative_to(source_dir)
                zipf.write(file_path, arcname)


# ============ 报告生成功能 ============
def _get_file_meta(file_path: str) -> Dict[str, str]:
    """获取文件 MD5、创建时间、修改时间。文件不存在或出错时返回空字符串。"""
    out = {'md5': '', 'ctime': '', 'mtime': ''}
    p = Path(file_path)
    if not p.exists() or not p.is_file():
        return out
    try:
        h = hashlib.md5()
        with open(p, 'rb') as f:
            for chunk in iter(lambda: f.read(65536), b''):
                h.update(chunk)
        out['md5'] = h.hexdigest()
        st = p.stat()
        out['ctime'] = datetime.fromtimestamp(st.st_ctime).strftime('%Y-%m-%d %H:%M:%S')
        out['mtime'] = datetime.fromtimestamp(st.st_mtime).strftime('%Y-%m-%d %H:%M:%S')
    except Exception:
        pass
    return out


_report_history: List[Dict] = []
def _save_report(current_data: List[Dict], generate_report: bool):
    """保存报告数据"""
    global _report_history
    _report_history.extend(current_data)

    if generate_report:
        _generate_html_report(_report_history)
def _format_size(size_bytes: int) -> str:
    """格式化文件大小"""
    if size_bytes < 1024:
        return f"{size_bytes} B"
    elif size_bytes < 1024 * 1024:
        return f"{size_bytes / 1024:.2f} KB"
    else:
        return f"{size_bytes / (1024 * 1024):.2f} MB"
def _generate_html_report(data: List[Dict]):
    """生成HTML可视化报告"""
    if not data:
        return

    is_english = data[0].get("is_english", False)
    t = TEXTS_EN if is_english else TEXTS_ZH
    success_count = sum(1 for item in data if item["status"] == t["status_success"])
    failed_count = len(data) - success_count
    macro_found_count = sum(1 for item in data if item["vba_found"])

    total_vba_size = sum(item["vba_size"] for item in data)
    lang_attr = "en" if is_english else "zh-CN"

    html_content = f"""<!DOCTYPE html>
<html lang="{lang_attr}">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{t['report_title']}</title>
    <style>
        * {{ margin: 0; padding: 0; box-sizing: border-box; }}
        body {{
            font: 14px/1.5 "Segoe UI", "Microsoft YaHei", sans-serif;
            background: #f5f5f5;
            color: #333;
            padding: 24px;
        }}
        .container {{ max-width: 960px; margin: 0 auto; }}
        .header {{
            margin-bottom: 24px;
            padding-bottom: 16px;
            border-bottom: 1px solid #ddd;
        }}
        .header h1 {{ font-size: 18px; font-weight: 600; color: #111; }}
        .header .meta {{ font-size: 12px; color: #666; margin-top: 4px; }}
        .stats {{
            display: flex;
            flex-wrap: wrap;
            gap: 16px;
            margin-bottom: 24px;
        }}
        .stat {{
            background: #fff;
            border: 1px solid #e0e0e0;
            padding: 12px 16px;
            min-width: 120px;
        }}
        .stat .n {{ font-size: 20px; font-weight: 600; color: #111; }}
        .stat .l {{ font-size: 12px; color: #666; margin-top: 2px; }}
        .stat.ok .n {{ color: #0a6b2c; }}
        .stat.err .n {{ color: #b91c1c; }}
        .stat.mac .n {{ color: #92400e; }}
        .stat.sz .n {{ color: #1e3a5f; }}
        .section {{
            background: #fff;
            border: 1px solid #e0e0e0;
            margin-bottom: 24px;
        }}
        .section-title {{
            font-size: 13px; font-weight: 600;
            padding: 10px 12px;
            border-bottom: 1px solid #e0e0e0;
            background: #fafafa;
        }}
        table {{ width: 100%; border-collapse: collapse; }}
        th {{
            font-size: 12px; font-weight: 600; color: #555;
            text-align: left;
            padding: 8px 12px;
            border-bottom: 1px solid #e0e0e0;
            background: #fafafa;
        }}
        td {{
            font-size: 13px;
            padding: 8px 12px;
            border-bottom: 1px solid #eee;
        }}
        tr:hover td {{ background: #fafafa; }}
        .badge {{
            display: inline-block;
            padding: 2px 8px;
            font-size: 12px;
            border: 1px solid transparent;
        }}
        .badge-success {{ background: #e8f5e9; color: #1b5e20; border-color: #c8e6c9; }}
        .badge-failed {{ background: #ffebee; color: #b71c1c; border-color: #ffcdd2; }}
        .badge-macro {{ background: #fff3e0; color: #e65100; border-color: #ffe0b2; }}
        .badge-clean {{ background: #f5f5f5; color: #616161; }}
        tr.row-has-macro td {{ color: #b71c1c; }}
        tr.row-has-macro .file-name-cell {{ font-weight: 600; color: #b71c1c; }}
        .size-tag {{ font-size: 12px; color: #666; }}
        .file-path {{
            font-size: 12px; color: #666;
            max-width: 280px;
            overflow: hidden;
            text-overflow: ellipsis;
            white-space: nowrap;
        }}
        .btn-detail {{
            font-size: 12px; color: #555; cursor: pointer;
            background: #f0f0f0; border: 1px solid #ccc; padding: 4px 10px;
            border-radius: 4px;
        }}
        .btn-detail:hover {{ background: #e5e5e5; }}
        tr.detail-row {{ background: #fafafa; }}
        tr.detail-row-hide {{ display: none; }}
        tr.detail-row > td {{
            padding: 12px;
            border-bottom: 1px solid #e0e0e0;
            vertical-align: top;
        }}
        .detail-panel {{
            padding: 0;
            font-size: 12px;
        }}
        .detail-block {{ margin-bottom: 12px; }}
        .detail-block:last-child {{ margin-bottom: 0; }}
        .detail-block .title {{ font-weight: 600; color: #333; margin-bottom: 6px; }}
        .detail-block .line {{ color: #555; margin: 2px 0; }}
        .detail-macro {{ margin-top: 12px; padding-top: 12px; border-top: 1px solid #e0e0e0; }}
        .detail-panel pre {{
            margin: 0;
            padding: 10px;
            background: #f5f5f5;
            border: 1px solid #e0e0e0;
            border-radius: 4px;
            font: 12px/1.5 Consolas, monospace;
            white-space: pre-wrap;
            word-break: break-all;
            max-height: 240px;
            overflow: auto;
        }}
        .detail-panel pre code {{ font: inherit; }}
        .filter-bar {{
            display: flex;
            flex-wrap: wrap;
            align-items: center;
            gap: 12px;
            padding: 10px 12px;
            border-bottom: 1px solid #e0e0e0;
            background: #fafafa;
        }}
        .filter-bar label {{ font-size: 12px; color: #555; margin-right: 4px; }}
        .filter-bar select {{
            font-size: 12px;
            padding: 4px 8px;
            border: 1px solid #ccc;
            background: #fff;
        }}
        .filter-bar input[type="text"] {{
            font-size: 12px;
            padding: 4px 8px;
            border: 1px solid #ccc;
            width: 180px;
        }}
        .filter-bar .filter-item {{ display: flex; align-items: center; }}
        .footer {{
            font-size: 12px;
            color: #888;
            text-align: center;
            margin-top: 24px;
        }}
        tr.filter-hide {{ display: none; }}
        @media (max-width: 640px) {{
            .file-path {{ max-width: 140px; }}
        }}
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>{t['report_title']}</h1>
            <p class="meta">{t['generated_at']} {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>
        </div>

        <div class="stats">
            <div class="stat ok"><div class="n">{success_count}</div><div class="l">{t['stat_ok']}</div></div>
            <div class="stat err"><div class="n">{failed_count}</div><div class="l">{t['stat_err']}</div></div>
            <div class="stat mac"><div class="n">{macro_found_count}</div><div class="l">{t['stat_mac']}</div></div>
            <div class="stat sz"><div class="n">{_format_size(total_vba_size)}</div><div class="l">{t['stat_sz']}</div></div>
        </div>

        <div class="section">
            <div class="section-title">{t['section_title']}</div>
            <div class="filter-bar">
                <div class="filter-item">
                    <label>{t['filter_status']}</label>
                    <select id="filter-status" onchange="filterTable()">
                        <option value="">{t['filter_all']}</option>
                        <option value="{t['filter_success']}">{t['filter_success']}</option>
                        <option value="{t['filter_fail']}">{t['filter_fail']}</option>
                    </select>
                </div>
                <div class="filter-item">
                    <label>{t['filter_macro']}</label>
                    <select id="filter-macro" onchange="filterTable()">
                        <option value="">{t['filter_all']}</option>
                        <option value="1">{t['filter_has_macro']}</option>
                        <option value="0">{t['filter_no_macro']}</option>
                    </select>
                </div>
                <div class="filter-item">
                    <label>{t['filter_name']}</label>
                    <input type="text" id="filter-name" placeholder="{t['placeholder_keyword']}" oninput="filterTable()">
                </div>
            </div>
            <table>
                <thead>
                    <tr>
                        <th>{t['th_filename']}</th>
                        <th>{t['th_orig_size']}</th>
                        <th>{t['th_macro_size']}</th>
                        <th>{t['th_status']}</th>
                        <th>{t['th_result']}</th>
                        <th>{t['th_time']}</th>
                        <th>{t['th_action']}</th>
                    </tr>
                </thead>
                <tbody>
                    {_generate_table_rows(data, t)}
                </tbody>
            </table>
        </div>

        <div class="footer">{t['footer']}</div>
    </div>

    <script>
        function toggleDetail(rowIndex) {{
            var tr = document.getElementById('detail-row-' + rowIndex);
            if (!tr) return;
            var wasHidden = tr.classList.contains('detail-row-hide');
            var all = document.querySelectorAll('tr.detail-row');
            for (var i = 0; i < all.length; i++) all[i].classList.add('detail-row-hide');
            if (wasHidden) tr.classList.remove('detail-row-hide');
        }}
        function filterTable() {{
            var status = (document.getElementById('filter-status').value || '').trim();
            var macro = (document.getElementById('filter-macro').value || '').trim();
            var name = (document.getElementById('filter-name').value || '').trim().toLowerCase();
            var rows = document.querySelectorAll('.section tbody tr');
            for (var i = 0; i < rows.length; i++) {{
                var tr = rows[i];
                if (tr.classList.contains('detail-row')) {{
                    var prev = tr.previousElementSibling;
                    tr.classList.toggle('filter-hide', prev ? prev.classList.contains('filter-hide') : true);
                    continue;
                }}
                var show = true;
                if (status && tr.getAttribute('data-status') !== status) show = false;
                if (macro !== '' && tr.getAttribute('data-macro') !== macro) show = false;
                if (name) {{
                    var text = (tr.querySelector('.file-name-cell') || tr.querySelector('td')).textContent || '';
                    if (text.toLowerCase().indexOf(name) === -1) show = false;
                }}
                tr.classList.toggle('filter-hide', !show);
            }}
        }}
    </script>
</body>
</html>"""

    # 保存报告到当前工作目录，文件名：日期时间_macro_cleaner_report.html
    report_name = datetime.now().strftime("%Y-%m-%d_%H-%M-%S") + "_macro_cleaner_report.html"
    report_path = Path.cwd() / report_name
    report_path.write_text(html_content, encoding='utf-8')
def _generate_table_rows(data: List[Dict], t: Dict[str, str]) -> str:
    """生成表格行HTML：有宏的行显示红色，详情展开后直接显示宏代码。t 为 TEXTS_ZH 或 TEXTS_EN。"""
    status_ok = t["status_success"]
    rows = []
    for i, item in enumerate(data):
        status_class = "badge-success" if item["status"] == status_ok else "badge-failed"
        macro_class = "badge-macro" if item["vba_found"] else "badge-clean"
        macro_text = _format_size(item["vba_size"]) if item["vba_found"] else t["filter_no_macro"]
        row_class = " class=\"row-has-macro\"" if item["vba_found"] else ""
        name_cell_class = "file-name-cell" if item["vba_found"] else ""
        vba_code_escaped = html_module.escape(item.get("vba_code", ""))

        # 操作列：仅「详情」按钮；详情内容在下方独立 tr 中展开
        orig_md5 = html_module.escape(item.get("original_md5", "") or "—")
        orig_ctime = html_module.escape(item.get("original_ctime", "") or "—")
        orig_mtime = html_module.escape(item.get("original_mtime", "") or "—")
        out_md5 = html_module.escape(item.get("output_md5", "") or "—")
        out_ctime = html_module.escape(item.get("output_ctime", "") or "—")
        out_mtime = html_module.escape(item.get("output_mtime", "") or "—")
        detail_content = f"""<div class="detail-panel">
            <div class="detail-block">
                <div class="title">{t['detail_orig_file']}</div>
                <div class="line">{t['detail_hash']}: {orig_md5}</div>
                <div class="line">{t['detail_created']}: {orig_ctime}</div>
                <div class="line">{t['detail_modified']}: {orig_mtime}</div>
            </div>"""
        if item["vba_found"]:
            detail_content += f"""<div class="detail-block">
                <div class="title">{t['detail_output_file']}</div>
                <div class="line">{t['detail_hash']}: {out_md5}</div>
                <div class="line">{t['detail_created']}: {out_ctime}</div>
                <div class="line">{t['detail_modified']}: {out_mtime}</div>
            </div>
            <div class="detail-block detail-macro">
                <div class="title">{t['detail_macro']}</div>
                <pre><code>{vba_code_escaped}</code></pre>
            </div>"""
        detail_content += "</div>"
        action_cell = f'<td><button type="button" class="btn-detail" onclick="toggleDetail({i})">{t["btn_detail"]}</button></td>'

        data_status = html_module.escape(item["status"], quote=True)
        data_macro = "1" if item["vba_found"] else "0"
        data_filename = html_module.escape(item["file_name"], quote=True)
        data_row = f"""
        <tr{row_class} data-status="{data_status}" data-macro="{data_macro}" data-filename="{data_filename}">
            <td>
                <div class="{name_cell_class}" style="font-weight: 600; color: #1f2937;">{html_module.escape(item["file_name"])}</div>
                <div class="file-path" title="{html_module.escape(item["file_path"], quote=True)}">{html_module.escape(item["file_path"])}</div>
            </td>
            <td><span class="size-tag">{_format_size(item["file_size"])}</span></td>
            <td><span class="badge {macro_class}">{macro_text}</span></td>
            <td><span class="badge {status_class}">{item["status"]}</span></td>
            <td>{html_module.escape(item["message"])}</td>
            <td>{item["timestamp"]}</td>
            {action_cell}
        </tr>
        <tr id="detail-row-{i}" class="detail-row detail-row-hide">
            <td colspan="7">{detail_content}</td>
        </tr>"""
        rows.append(data_row)

    return "".join(rows)