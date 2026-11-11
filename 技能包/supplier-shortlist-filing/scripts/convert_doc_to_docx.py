# -*- coding: utf-8 -*-
"""将老式 .doc 转换为 .docx（供 python-docx 流水线使用）。

用法:
    python convert_doc_to_docx.py <src.doc> <dst.docx>

原理:
    优先 WPS (KWPS.Application)，其次 MS Word (Word.Application) 的 COM 接口
    以 FileFormat=12 另存为 .docx。转换后建议人工在 WPS/Word 中打开确认版式。
"""
import os
import sys


def convert(src, dst):
    try:
        import win32com.client
    except ImportError:
        sys.exit("缺少 pywin32，请先: pip install pywin32")

    if not os.path.exists(src):
        sys.exit(f"源文件不存在: {src}")
    dst = os.path.abspath(dst)
    src = os.path.abspath(src)

    for appname in ("KWPS.Application", "Word.Application"):
        try:
            app = win32com.client.Dispatch(appname)
        except Exception:
            continue
        try:
            app.Visible = False
            try:
                app.DisplayAlerts = 0
            except Exception:
                pass
            doc = app.Documents.Open(src, ReadOnly=False)
            doc.SaveAs(dst, FileFormat=12)  # 12 = wdFormatXMLDocument (.docx)
            doc.Close(False)
            app.Quit()
            if os.path.exists(dst):
                print(f"OK: {src} -> {dst} (via {appname})")
                return True
        except Exception as e:
            try:
                app.Quit()
            except Exception:
                pass
            print(f"FAIL {appname}: {e}")
    sys.exit("WPS 与 MS Word COM 均不可用或转换失败，请人工在 WPS/Word 中另存为 .docx")


if __name__ == "__main__":
    if len(sys.argv) != 3:
        sys.exit(__doc__)
    convert(sys.argv[1], sys.argv[2])
