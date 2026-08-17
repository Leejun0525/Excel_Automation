# Excel 自动清理工具（Windows 可打包为单文件可执行）
# 用法：双击打包后的 exe，或用 Python 运行此脚本

import os
import sys
import traceback
from datetime import datetime
import argparse
import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk

APP_TITLE = "Excel 自动清理工具"


def detect_and_format_dates(df):
    """
    更健壮的日期检测与格式化：
    - 支持 Excel 序列日期（数字形式）
    - 尝试解析不同格式的字符串
    - 如果解析成功则格式化为 YYYY-MM-DD 字符串；解析失败则保留原始值（转为字符串）
    """
    df = df.copy()

    def _parse_cell(val):
        # 保留空值
        if pd.isna(val):
            return pd.NaT
        # 尝试识别 Excel 序列日期（通常为数字且在合理范围）
        if isinstance(val, (int, float)):
            # 排除非常大的数或非常小的数
            try:
                # Excel 的序列日期以 1899-12-30 为起点（兼容 Excel 闰年 bug）
                parsed = pd.to_datetime(val, unit='d', origin='1899-12-30')
                return parsed
            except Exception:
                pass
        # 常规字符串/对象解析
        try:
            return pd.to_datetime(val, errors='coerce')
        except Exception:
            return pd.NaT

    for col in df.columns:
        try:
            parsed = df[col].apply(_parse_cell)
        except Exception:
            continue

        non_null_ratio = parsed.notna().sum() / max(1, len(parsed))

        # 如果超过 40% 能解析为日期，则认为这是日期列
        if non_null_ratio >= 0.4:
            # 格式化为 YYYY-MM-DD 字符串；对于解析失败的单元格，保留原始值（转为字符串）
            try:
                formatted = parsed.dt.strftime('%Y-%m-%d')
                df[col] = formatted.where(parsed.notna(), df[col].astype(str))
            except Exception:
                # 如果格式化失败则只将能解析的部分格式化，其他保留原值
                df[col] = [d.strftime('%Y-%m-%d') if not pd.isna(d) else str(orig) for d, orig in zip(parsed, df[col])]

    return df


def strip_whitespace(df):
    # 去除字符串类型前后空白（保留 NaN）
    df = df.copy()
    for col in df.columns:
        # 只处理对象/字符串列
        if df[col].dtype == object or str(df[col].dtype).startswith('string'):
            mask = df[col].notna()
            if mask.any():
                df.loc[mask, col] = df.loc[mask, col].astype(str).str.strip()
    return df


def auto_sort_column(df):
    # 找到第一个非全空列用于排序
    for col in df.columns:
        if df[col].notna().any():
            try:
                return df.sort_values(by=col, kind='mergesort')
            except Exception:
                try:
                    return df.sort_values(by=col, kind='mergesort', key=lambda x: x.astype(str))
                except Exception:
                    return df
    return df


def process_file(path, output_dir=None, dedupe_by=None):
    """
    处理单个文件：
    - path: 输入文件路径
    - output_dir: 可选输出目录
    - dedupe_by: 可选的去重列列表（示例：['ID','Name']），若为 None 则按全部列去重
    返回 (bool success, str log)
    """
    log_lines = []
    start_time = datetime.now()
    log_lines.append(f"开始处理: {path}")

    try:
        # 读取第一张表（sheet）
        df = pd.read_excel(path, sheet_name=0, engine='openpyxl')
    except Exception as e:
        log_lines.append(f"读取 Excel 失败: {e}")
        return False, '\n'.join(log_lines)

    original_rows = len(df)
    log_lines.append(f"原始行数: {original_rows}")

    # 去除全空行
    df = df.dropna(how='all')
    after_dropna_rows = len(df)
    log_lines.append(f"去除全空行后行数: {after_dropna_rows}")

    # 去除字符串前后空白
    df = strip_whitespace(df)

    # 标准化日期列
    df = detect_and_format_dates(df)

    # 去重（按用户传入列或默认所有列完全相等）
    df_before_dup = len(df)
    try:
        if dedupe_by:
            # 确保去重列存在
            dedupe_existing = [c for c in dedupe_by if c in df.columns]
            if dedupe_existing:
                df = df.drop_duplicates(subset=dedupe_existing)
            else:
                # 如果指定列不存在，则回退到全部列去重
                df = df.drop_duplicates()
        else:
            df = df.drop_duplicates()
    except Exception as e:
        log_lines.append(f"去重失败，回退到默认去重: {e}")
        df = df.drop_duplicates()

    df_after_dup = len(df)
    log_lines.append(f"去重前行数: {df_before_dup}, 去重后行数: {df_after_dup}")

    # 排序（默认按第一个非空列）
    df = auto_sort_column(df)

    # 汇总信息
    summary = {
        '原始行数': original_rows,
        '去除全空行后行数': after_dropna_rows,
        '去重后行数': df_after_dup,
        '处理时间': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    }

    # 每列缺失值计数
    missing_counts = df.isna().sum().astype(int)
    missing_df = pd.DataFrame({'列名': missing_counts.index, '缺失值数量': missing_counts.values})

    # 输出文件路径
    base = os.path.splitext(os.path.basename(path))[0]
    if output_dir is None:
        output_dir = os.path.dirname(path)
    os.makedirs(output_dir, exist_ok=True)

    cleaned_name = os.path.join(output_dir, f"{base}_cleaned.xlsx")
    log_name = os.path.join(output_dir, f"{base}_cleaned.log.txt")

    try:
        with pd.ExcelWriter(cleaned_name, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Cleaned')
            # 写入汇总表
            summary_df = pd.DataFrame(list(summary.items()), columns=['项', '值'])
            summary_df.to_excel(writer, index=False, sheet_name='Summary')
            missing_df.to_excel(writer, index=False, sheet_name='MissingValues')
    except Exception as e:
        log_lines.append(f"写入清理文件失败: {e}")
        return False, '\n'.join(log_lines)

    # 写日志
    log_lines.append(f"清理后文件: {cleaned_name}")
    log_lines.append(f"写入日志: {log_name}")
    log_lines.append(f"总共耗时: {(datetime.now()-start_time).total_seconds():.2f} 秒")

    try:
        with open(log_name, 'w', encoding='utf-8') as f:
            f.write('\n'.join(log_lines))
    except Exception as e:
        # 日志写入失败不影响主流程
        log_lines.append(f"日志写入失败: {e}")

    return True, '\n'.join(log_lines)


# 简单的 GUI
class App:
    def __init__(self, root):
        self.root = root
        root.title(APP_TITLE)
        root.geometry('640x360')
        root.resizable(False, False)

        self.label = tk.Label(root, text='选择需要清理的 Excel 文件（.xlsx）。可选择输出目录并开始批量处理。', wraplength=600, justify='left')
        self.label.pack(padx=10, pady=(14, 6))

        btn_frame = tk.Frame(root)
        btn_frame.pack(pady=6)

        self.select_btn = tk.Button(btn_frame, text='选择文件', width=20, command=self.select_file)
        self.select_btn.grid(row=0, column=0, padx=6)

        self.output_btn = tk.Button(btn_frame, text='选择输出目录', width=20, command=self.select_output_dir)
        self.output_btn.grid(row=0, column=1, padx=6)

        self.process_btn = tk.Button(btn_frame, text='开始处理选中文件', width=20, state='disabled', command=self.process_selected)
        self.process_btn.grid(row=0, column=2, padx=6)

        # 进度条与日志
        self.progressbar = ttk.Progressbar(root, mode='indeterminate')
        self.progressbar.pack(fill='x', padx=8, pady=(6, 4))

        self.progress = tk.Text(root, height=10, width=80)
        self.progress.pack(padx=8, pady=8)
        self.progress.configure(state='disabled')

        self.selected_path = None
        self.output_dir = None

    def select_file(self):
        path = filedialog.askopenfilename(title='选择 Excel 文件', filetypes=[('Excel 文件', '*.xlsx;*.xls')])
        if path:
            self.selected_path = path
            self.append_log(f'已选: {path}')
            self.process_btn.configure(state='normal')

    def select_output_dir(self):
        folder = filedialog.askdirectory(title='选择输出目录（可选）')
        if folder:
            self.output_dir = folder
            self.append_log(f'已选输出目录: {folder}')

    def process_selected(self):
        if not self.selected_path:
            messagebox.showwarning(APP_TITLE, '请先选择一个 Excel 文件')
            return

        self.append_log('开始处理...')
        self.progressbar.start(10)
        self.root.update()

        try:
            ok, log = process_file(self.selected_path, output_dir=self.output_dir)
            self.append_log(log)
            if ok:
                messagebox.showinfo(APP_TITLE, '处理完成！清理后的文件已生成。')
            else:
                messagebox.showerror(APP_TITLE, '处理失败，详情见日志输出')
        except Exception as e:
            tb = traceback.format_exc()
            self.append_log(str(e) + '\n' + tb)
            messagebox.showerror(APP_TITLE, '发生未处理的异常，已在日志中记录')
        finally:
            self.progressbar.stop()

    def append_log(self, text):
        self.progress.configure(state='normal')
        self.progress.insert('end', text + '\n')
        self.progress.see('end')
        self.progress.configure(state='disabled')


def cli_main(argv=None):
    parser = argparse.ArgumentParser(description='Excel 自动清理工具（命令行模式）')
    parser.add_argument('-i', '--input', help='单个 Excel 文件或目录（配合 --batch）')
    parser.add_argument('-o', '--output-dir', help='输出目录，默认为输入文件所在目录')
    parser.add_argument('-d', '--dedupe-cols', help='用于去重的列，多个用逗号分隔，例如 ID,Name')
    parser.add_argument('--batch', action='store_true', help='如果输入是目录，则批量处理目录下所有 .xlsx/.xls 文件')

    args = parser.parse_args(argv)

    if args.input:
        paths = []
        if os.path.isdir(args.input):
            if args.batch:
                for fn in os.listdir(args.input):
                    if fn.lower().endswith(('.xlsx', '.xls')):
                        paths.append(os.path.join(args.input, fn))
            else:
                print('输入是目录，请加 --batch 以批量处理，或指定单个文件路径')
                sys.exit(1)
        elif os.path.isfile(args.input):
            paths = [args.input]
        else:
            print('输入路径无效')
            sys.exit(1)

        dedupe_by = args.dedupe_cols.split(',') if args.dedupe_cols else None

        for p in paths:
            ok, log = process_file(p, output_dir=args.output_dir, dedupe_by=dedupe_by)
            print(log)
        return


def main():
    root = tk.Tk()
    app = App(root)
    root.mainloop()


if __name__ == '__main__':
    # 如果通过命令行传入参数则使用 CLI 模式，否则启动 GUI
    if len(sys.argv) > 1:
        cli_main()
    else:
        main()
