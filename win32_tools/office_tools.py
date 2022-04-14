#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import win32com.client as win32

def dispatch_office_app(app: str) -> win32.gencache.EnsureDispatch:
    """Clears win32 gen_py cache in attempt to dispatch office app.
    Example input: "Word", "Excel", etc."""
    # https://gist.github.com/rdapaz/63590adb94a46039ca4a10994dff9dbe
    print(f"Opening Microsoft {app} in background...", end=" ")
    try:
        office_app = win32.gencache.EnsureDispatch(app + ".Application")
        print("Opened.")
    except AttributeError:
        # Corner case dependencies.
        import os
        import re
        import sys
        import shutil
        # Remove cache and try again.
        MODULE_LIST = [m.__name__ for m in sys.modules.values()]
        for module in MODULE_LIST:
            if re.match(r"win32com\.gen_py\..+", module):
                del sys.modules[module]
        shutil.rmtree(os.path.join(os.environ.get("LOCALAPPDATA"), "Temp", "gen_py"))
        print("Opened (cache cleared).")
        office_app = win32.gencache.EnsureDispatch(app + ".Application")
    return office_app

def mark_index_entries(
    filename: str="CommentResponseSection.docx",
    automark: str="AutoMark.docx",
    add_index: bool=False
) -> None:
    word = dispatch_office_app("Word")
    cwd = os.getcwd()
    doc_filepath = os.path.join(cwd,filename)
    automark_filepath = os.path.join(cwd,automark)
    doc = word.Documents.Open(doc_filepath, Visible=False)
    
    def add_index_entries(doc,automark_filepath):
        index = doc.Indexes
        index.AutoMarkEntries(automark_filepath)
        print("Index entries marked.")
        return None
    
    def clean_index_entries(doc):
        docrng = doc.Content
        m = {r"zyx(*)xyz": r"\1"}
        for key, value in m.items():
            find = docrng.Find
            find.Text = key
            find.Replacement.Text = value
            find.Wrap = 1
            find.Forward = True
            find.MatchWildcards = True # wildcard search
            find.Execute(Replace=2)
        print("Index entries cleaned.")
        remove_line_breaks(doc)
        return None
    
    def append_index(doc):
        index = doc.Indexes
        docrng = doc.Content
        docrng.Collapse(0)
        docrng.InsertBreak(7)
        docrng.Collapse(0)
        docrng.Style = -2
        docrng.InsertAfter("Commenter Index\r")
        docrng.Collapse(0)
        docrng.Style = -1
        index.Add(Range=docrng,NumberOfColumns=2)
        index.Format = 4
        print("Index appended to end of document.")
        return None

    add_index_entries(doc,automark_filepath)
    clean_index_entries(doc)
    if add_index: append_index(doc)
    doc.Save()
    doc = None
    word.Application.Quit()
    return None

def remove_line_breaks(doc):
    docrng = doc.Content
    m = {r"^l^l": r"^l", r"^l": r"^p", r"^p^p": r"^p"}
    for key, value in m.items():
        find = docrng.Find
        find.Text = key
        find.Replacement.Text = value
        find.Wrap = 1
        find.Forward = True
        find.MatchWildcards = False
        find.Execute(Replace=2)
    print("Line breaks replaced with paragraph breaks.")
    return None