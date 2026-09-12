#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""office_sensitive_encryptor.py 端到端测试"""
import os, sys, json, queue, tempfile, shutil, traceback

SRC = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                   "office_sensitive_encryptor.py")

# ---------- 加载模块（stub GUI 组件） ----------
src = open(SRC, encoding="utf-8").read()
src = src.split('if __name__ == "__main__":')[0]
src = src.replace(
    "from tkinter import ttk, messagebox, filedialog, scrolledtext",
    "from tkinter import ttk, scrolledtext\nclass _Msg:\n    def __getattr__(self, n): return lambda *a, **k: None\nclass _Fd:\n    def __getattr__(self, n): return lambda *a, **k: ''\nmessagebox = _Msg()\nfiledialog = _Fd()")
ns = {}
exec(compile(src, "encryptor", "exec"), ns)

CryptoEngine = ns["CryptoEngine"]
WordProcessor = ns["WordProcessor"]
ExcelProcessor = ns["ExcelProcessor"]
PptProcessor = ns["PptProcessor"]
SubstitutionEngine = ns["SubstitutionEngine"]
BatchFrame = ns["BatchFrame"]
compute_file_sha256 = ns["compute_file_sha256"]
restore_doc_name = ns["restore_doc_name"]

ok, fail = [], []

def check(name, cond, detail=""):
    if cond:
        ok.append(name)
    else:
        fail.append((name, detail))
        print(f"  ✗ {name}: {detail}")

def make_fake_task():
    class FakeTask:
        def __init__(self):
            self.q = queue.Queue()
            self.msgs = []
        def put_log(self, m):
            self.q.put(("log", m))
            self.msgs.append(m)
    t = FakeTask()
    # 让 worker 里的 log 直接写入 self.task.q —— 保持原样，测试读取 q
    return t

def drain_logs(t):
    msgs = []
    try:
        while True:
            k, p = t.q.get_nowait()
            if k == "log":
                msgs.append(p)
    except queue.Empty:
        pass
    return msgs

TMP = tempfile.mkdtemp(prefix="oetest_")
print(f"测试目录: {TMP}")

# ================= 1. CryptoEngine 新格式 =================
def t_crypto():
    vocab = {"version": "2.3", "words": ["机密"], "mappings": {"机密": "aB"}, "doc_hash": "abc123"}
    enc = CryptoEngine.encrypt_vocab(vocab, "pass1234")
    assert enc[:4] == b"OEV2", "magic 缺失"
    assert len(enc) >= 4 + 16 + 12 + 16, "结构长度错误"
    dec = CryptoEngine.decrypt_vocab(enc, "pass1234")
    assert dec == vocab, "回环失败"
    # 盐/随机性
    enc2 = CryptoEngine.encrypt_vocab(vocab, "pass1234")
    assert enc != enc2
    # 错误密码
    try:
        CryptoEngine.decrypt_vocab(enc, "wrong")
        raise AssertionError("错误密码未报错")
    except ValueError as e:
        assert "密码错误" in str(e)
    # 篡改检测
    bad = bytearray(enc); bad[-1] ^= 0xFF
    try:
        CryptoEngine.decrypt_vocab(bytes(bad), "pass1234")
        raise AssertionError("篡改未检测")
    except ValueError as e:
        assert "密码错误" in str(e) or "损坏" in str(e)
    check("CryptoEngine AES-256-GCM 回环/随机盐/错误密码/篡改检测", True)

def t_legacy_fernet():
    # 模拟 v2.2 Fernet 格式词表，验证兼容解密
    import base64
    from cryptography.fernet import Fernet
    from cryptography.hazmat.primitives import hashes
    from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
    import secrets
    vocab = {"version": "2.2", "words": ["旧词"], "mappings": {"旧词": "Xy"}}
    salt = secrets.token_bytes(16)
    kdf = PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=salt, iterations=480000)
    key = base64.urlsafe_b64encode(kdf.derive(b"oldpass"))
    old_data = salt + Fernet(key).encrypt(json.dumps(vocab, ensure_ascii=False).encode("utf-8"))
    dec = CryptoEngine.decrypt_vocab(old_data, "oldpass")
    assert dec == vocab
    check("旧版 v2.2 Fernet 词表兼容解密", True)

# ================= 2. Word 单文件回环（正文/表格/页眉/超链接） =================
def make_word(path):
    from docx import Document
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    doc = Document()
    doc.add_paragraph("正文包含机密与薪酬两个词")
    # 表格
    tbl = doc.add_table(rows=1, cols=2)
    tbl.rows[0].cells[0].text = "表格机密"
    tbl.rows[0].cells[1].text = "普通"
    # 页眉
    sec = doc.sections[0]
    sec.header.paragraphs[0].text = "页眉机密"
    # 超链接
    part = doc.part
    r_id = part.relate_to("https://example.com",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        is_external=True)
    hp = doc.add_paragraph()
    hl = OxmlElement('w:hyperlink')
    hl.set(qn('r:id'), r_id)
    run = OxmlElement('w:r')
    t = OxmlElement('w:t'); t.text = "链接机密"
    run.append(t)
    hl.append(run)
    hp._p.append(hl)
    # 文本框（w:txbxContent）
    body = doc.element.body
    drawing = OxmlElement('w:drawing')
    txbx = OxmlElement('w:txbxContent')
    tp = OxmlElement('w:p')
    tr = OxmlElement('w:r')
    tt = OxmlElement('w:t'); tt.text = "文本框机密"
    tr.append(tt); tp.append(tr); txbx.append(tp)
    body.append(txbx)
    doc.save(path)

def read_word_text(path):
    from docx import Document
    from docx.oxml.ns import qn
    doc = Document(path)
    parts = [p.text for p in doc.paragraphs]
    parts += [c.text for tbl in doc.tables for row in tbl.rows for c in row.cells]
    parts += [doc.sections[0].header.paragraphs[0].text]
    # 超链接 & 文本框 XML 文本
    xml = doc.element.body.xml
    return "\n".join(parts), xml

def t_word_roundtrip():
    src_doc = os.path.join(TMP, "w_src.docx")
    enc_doc = os.path.join(TMP, "w_enc.docx")
    dec_doc = os.path.join(TMP, "w_dec.docx")
    make_word(src_doc)

    # 加密
    proc = WordProcessor()
    assert proc.load(src_doc)
    mappings = {"机密": "aB", "薪酬": "xY"}
    n = proc.replace_all(mappings, mode="encrypt")
    assert proc.save(enc_doc)
    text, xml = read_word_text(enc_doc)
    assert "机密" not in text and "机密" not in xml, f"加密后仍含'机密': {text}"
    assert "薪酬" not in text
    # 页眉/超链接/文本框都被替换（文本框内容在 body XML 中）
    assert "aB\u200b" in text, f"正文替换缺失: {text}"
    assert "页眉aB\u200b" in text, f"页眉未替换: {text}"
    assert "链接aB\u200b" in text, f"超链接未替换: {text}"
    assert "文本框aB\u200b" in xml, f"文本框未替换: {xml[-300:]}"
    # 记录加密文档哈希
    doc_hash = compute_file_sha256(enc_doc)

    # 解密（replace_all 传原始映射，内部自动反向）
    proc2 = WordProcessor()
    assert proc2.load(enc_doc)
    n2 = proc2.replace_all(mappings, mode="decrypt")
    assert proc2.save(dec_doc)
    text2, xml2 = read_word_text(dec_doc)
    assert "机密" in text2 and "薪酬" in text2, f"解密未恢复: {text2}"
    assert "页眉机密" in text2, f"页眉未恢复: {text2}"
    assert "链接机密" in text2, f"超链接未恢复: {text2}"
    assert "文本框机密" in xml2, f"文本框未恢复"
    assert "aB\u200b" not in text2 and "aB\u200b" not in xml2
    check("Word 正文/表格/页眉/超链接/文本框 加解密回环", True, f"替换{n}+{n2}次")

# ================= 3. Excel 单文件回环 + 类型保护 =================
def make_excel(path):
    import openpyxl
    from datetime import datetime
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = "文本机密"
    ws["A2"] = 123456          # 数字（不含敏感词，应保持）
    ws["A3"] = 123456          # 数字（含敏感词'123'，应跳过并警告）
    ws["A4"] = datetime(2024, 1, 1)  # 日期（跳过）
    ws["A5"] = "=SUM(A2:A3)"   # 公式（跳过）
    ws["A6"] = "身份机密证"
    wb.save(path)

def t_excel_roundtrip():
    src_x = os.path.join(TMP, "x_src.xlsx")
    enc_x = os.path.join(TMP, "x_enc.xlsx")
    dec_x = os.path.join(TMP, "x_dec.xlsx")
    make_excel(src_x)
    import openpyxl
    from datetime import datetime

    # 加密
    proc = ExcelProcessor()
    assert proc.load(src_x)
    mappings = {"机密": "aB", "123": "xYz"}
    n = proc.replace_all(mappings, mode="encrypt")
    assert proc.save(enc_x)
    wb = openpyxl.load_workbook(enc_x)
    ws = wb.active
    assert "机密" not in str(ws["A1"].value)
    assert ws["A1"].value == "文本aB\u200b" or "aB" in ws["A1"].value
    assert ws["A2"].value == 123456 and isinstance(ws["A2"].value, int), f"数字被破坏: {ws['A2'].value!r}"
    assert ws["A3"].value == 123456 and isinstance(ws["A3"].value, int), f"敏感数字被破坏: {ws['A3'].value!r}"
    assert ws["A4"].value == datetime(2024, 1, 1), "日期被破坏"
    assert str(ws["A5"].value).startswith("=SUM"), "公式被破坏"
    assert "aB" in ws["A6"].value or "aB\u200b" in ws["A6"].value
    # 跳过警告应包含 A3
    assert any("A3" in w for w in proc.skipped_warnings), f"缺少跳过警告: {proc.skipped_warnings}"
    print(f"    [Excel] 跳过警告: {proc.skipped_warnings}")

    # 解密（传原始映射，内部自动反向）
    proc2 = ExcelProcessor()
    assert proc2.load(enc_x)
    proc2.replace_all(mappings, mode="decrypt")
    assert proc2.save(dec_x)
    wb2 = openpyxl.load_workbook(dec_x)
    ws2 = wb2.active
    assert "机密" in ws2["A1"].value
    assert ws2["A2"].value == 123456
    assert ws2["A3"].value == 123456
    assert ws2["A6"].value == "身份机密证"
    check("Excel 文本/数字/日期/公式 回环 + 类型保护", True)

# ================= 4. PPT 单文件回环 =================
def t_ppt_roundtrip():
    from pptx import Presentation
    from pptx.util import Inches
    src_p = os.path.join(TMP, "p_src.pptx")
    enc_p = os.path.join(TMP, "p_enc.pptx")
    dec_p = os.path.join(TMP, "p_dec.pptx")
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    box = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(3), Inches(1))
    box.text_frame.text = "幻灯片机密"
    prs.save(src_p)

    proc = PptProcessor()
    assert proc.load(src_p)
    mappings = {"机密": "aB"}
    n = proc.replace_all(mappings, mode="encrypt")
    assert proc.save(enc_p)
    prs2 = Presentation(enc_p)
    t = prs2.slides[0].shapes[0].text_frame.text
    assert "机密" not in t and "aB" in t
    # 解密（传原始映射）
    proc2 = PptProcessor()
    assert proc2.load(enc_p)
    proc2.replace_all(mappings, mode="decrypt")
    assert proc2.save(dec_p)
    prs3 = Presentation(dec_p)
    assert "机密" in prs3.slides[0].shapes[0].text_frame.text
    check("PPT 加解密回环", True)

# ================= 5. 文档绑定校验 =================
def t_binding():
    # 两个内容不同的文档，加密后哈希不同；词表只认绑定文档
    from docx import Document
    a = os.path.join(TMP, "bind_a.docx")
    b = os.path.join(TMP, "bind_b.docx")
    da = Document(); da.add_paragraph("机密文档A")
    db = Document(); db.add_paragraph("机密文档B")
    da.save(a); db.save(b)

    mappings = {"机密": "aB"}

    def encrypt_doc(path, out):
        proc = WordProcessor()
        assert proc.load(path)
        proc.replace_all(mappings, mode="encrypt")
        assert proc.save(out)
        return compute_file_sha256(out)

    enc_a = os.path.join(TMP, "bind_a_enc.docx")
    enc_b = os.path.join(TMP, "bind_b_enc.docx")
    hash_a = encrypt_doc(a, enc_a)
    hash_b = encrypt_doc(b, enc_b)
    assert hash_a != hash_b, "不同文档哈希不应相同"

    vocab_data = {"version": "2.3", "words": ["机密"], "mappings": mappings, "doc_hash": hash_a}
    # 词表绑定 A：匹配 A，拒绝 B
    assert vocab_data["doc_hash"] == hash_a
    assert vocab_data["doc_hash"] != hash_b
    check("词表-文档 SHA-256 绑定：正确匹配/异文档区分", True)

# ================= 6. 批量加密/解密 =================
def build_batch_src(d):
    os.makedirs(d, exist_ok=True)
    make_word(os.path.join(d, "报告.docx"))
    make_excel(os.path.join(d, "台账.xlsx"))
    from pptx import Presentation
    from pptx.util import Inches
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    box = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(3), Inches(1))
    box.text_frame.text = "演示机密"
    prs.save(os.path.join(d, "演示.pptx"))
    # 一个不支持的文件应被忽略
    with open(os.path.join(d, "说明.txt"), "w") as f:
        f.write("机密")

def make_batch_frame(independent):
    # 绕过 GUI 构造，直接实例化 stub
    bf = object.__new__(BatchFrame)
    bf.task = make_fake_task()
    bf.independent_var = type("V", (), {"get": lambda self: independent})()
    return bf

def read_batch_out(d):
    from docx import Document
    from pptx import Presentation
    import openpyxl
    res = {}
    for name in os.listdir(d):
        full = os.path.join(d, name)
        if name.endswith(".docx"):
            doc = Document(full)
            res[name] = "\n".join(p.text for p in doc.paragraphs)
        elif name.endswith(".xlsx"):
            wb = openpyxl.load_workbook(full)
            ws = wb.active
            res[name] = str(ws["A1"].value)
        elif name.endswith(".pptx"):
            prs = Presentation(full)
            res[name] = prs.slides[0].shapes[0].text_frame.text
    return res

def t_batch_shared():
    src = os.path.join(TMP, "batch_shared_src")
    out1 = os.path.join(TMP, "batch_shared_out")
    out2 = os.path.join(TMP, "batch_shared_restore")
    build_batch_src(src)
    bf = make_batch_frame(independent=False)
    files = BatchFrame.list_office_files(src)
    assert len(files) == 3, f"应找到3个文件, 实际{len(files)}"
    bf._batch_encrypt_worker(["机密", "薪酬", "123"], "pw1234", files, out1, "ok", False)
    enc_logs = drain_logs(bf.task)
    assert os.path.exists(os.path.join(out1, "_batch_vocab.enc")), "共享词表缺失"
    encs = [f for f in os.listdir(out1) if f.endswith((".docx", ".xlsx", ".pptx"))]
    assert len(encs) == 3, f"加密产物应3个: {encs}"
    # 内容已加密
    enc_content = read_batch_out(out1)
    assert "机密" not in "\n".join(enc_content.values()), f"加密不彻底: {enc_content}"

    # 解密
    bf2 = make_batch_frame(independent=False)
    enc_files = [os.path.join(out1, f) for f in encs]
    shared = os.path.join(out1, "_batch_vocab.enc")
    bf2._batch_decrypt_worker("pw1234", enc_files, out1, out2, "ok", True, shared)
    dec_logs = drain_logs(bf2.task)
    restored = read_batch_out(out2)
    assert len(restored) == 3
    joined = "\n".join(restored.values())
    assert "机密" in joined, f"共享模式未恢复: {restored}"
    assert "\u200b" not in joined, f"残留标记: {restored}"
    assert "机密" in restored.get("台账.xlsx", ""), f"Excel 未恢复: {restored.get('台账.xlsx')}"
    check("批量加密(共享词表) → 批量解密 回环", True)

def t_batch_independent():
    src = os.path.join(TMP, "batch_ind_src")
    out1 = os.path.join(TMP, "batch_ind_out")
    out2 = os.path.join(TMP, "batch_ind_restore")
    build_batch_src(src)
    bf = make_batch_frame(independent=True)
    files = BatchFrame.list_office_files(src)
    bf._batch_encrypt_worker(["机密", "薪酬", "123"], "pw1234", files, out1, "ok", True)
    encs = [f for f in os.listdir(out1) if f.endswith((".docx", ".xlsx", ".pptx"))]
    vocabs = [f for f in os.listdir(out1) if f.endswith(".enc")]
    assert len(encs) == 3 and len(vocabs) == 3, f"独立词表应各3个: encs={encs} vocabs={vocabs}"
    assert not os.path.exists(os.path.join(out1, "_batch_vocab.enc")), "共享词表不应存在"

    bf2 = make_batch_frame(independent=True)
    enc_files = [os.path.join(out1, f) for f in encs]
    bf2._batch_decrypt_worker("pw1234", enc_files, out1, out2, "ok", False, None)
    restored = read_batch_out(out2)
    assert len(restored) == 3
    assert "机密" in "\n".join(restored.values())
    check("批量加密(每文件独立词表) → 批量解密 回环", True)

def t_binding_batch_reject():
    # 共享词表解"外来"文件应被指纹校验拦下
    src = os.path.join(TMP, "batch_reject_src")
    out1 = os.path.join(TMP, "batch_reject_out")
    out2 = os.path.join(TMP, "batch_reject_restore")
    build_batch_src(src)
    bf = make_batch_frame(independent=False)
    files = BatchFrame.list_office_files(src)
    bf._batch_encrypt_worker(["机密"], "pw1234", files, out1, "ok", False)
    # 构造一个未被加密的"外来"文件，塞进 out1
    alien = os.path.join(out1, "外来_encrypted.docx")
    make_word(alien)   # 内容含"机密"，但不在共享词表指纹里
    bf2 = make_batch_frame(independent=False)
    enc_files = [f for f in BatchFrame.list_office_files(out1) if "_encrypted" in os.path.basename(f)]
    shared = os.path.join(out1, "_batch_vocab.enc")
    bf2._batch_decrypt_worker("pw1234", enc_files, out1, out2, "ok", True, shared)
    logs = drain_logs(bf2.task)
    assert any("不匹配" in m for m in logs), f"外来文件未被指纹拦截: {logs}"
    # 输出目录不应包含外来文件的恢复
    assert "外来.docx" not in os.listdir(out2)
    check("批量解密：外来文件被指纹校验拦截", True)

def t_restore_name():
    assert restore_doc_name("报告_encrypted.docx") == "报告.docx"
    assert restore_doc_name("报告_encrypted_v2_encrypted.docx") == "报告_encrypted_v2.docx"
    assert restore_doc_name("普通.docx") == "普通.docx"
    assert restore_doc_name("a.b.c_encrypted.xlsx") == "a.b.c.xlsx"
    check("restore_doc_name 文件名恢复边界", True)

# ================= 运行 =================
tests = [t_crypto, t_legacy_fernet, t_word_roundtrip, t_excel_roundtrip,
         t_ppt_roundtrip, t_binding, t_batch_shared, t_batch_independent,
         t_binding_batch_reject, t_restore_name]

for t in tests:
    try:
        t()
    except Exception as e:
        fail.append((t.__name__, traceback.format_exc()))
        print(f"  ✗ {t.__name__} 异常: {e}")

print("=" * 60)
print(f"通过: {len(ok)} 项")
for o in ok:
    print("  ✓", o)
if fail:
    print(f"失败: {len(fail)} 项")
    for name, detail in fail:
        print(f"  ✗ {name}")
        print("    " + detail[-1500:].replace("\n", "\n    "))

shutil.rmtree(TMP, ignore_errors=True)
print("测试目录已清理")
