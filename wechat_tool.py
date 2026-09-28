import os
import sys
import json
import time
import traceback
from multiprocessing import freeze_support


def app_dir():
    """exe 所在目录(源码运行时为脚本目录),所有输出都放这里"""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def resource_path(rel):
    """打包资源所在位置"""
    base = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, rel)


# 让 exporter 能找到随包携带的 ffmpeg.exe
os.environ['PATH'] = resource_path('exporter') + os.pathsep + os.environ.get('PATH', '')


def is_admin():
    try:
        import ctypes
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        return False


# ---------------- 1. 解密 ----------------
def do_decrypt():
    from wxManager import Me
    from wxManager.decrypt import get_info_v4, get_info_v3
    from wxManager.decrypt.decrypt_dat import get_decode_code_v4
    from wxManager.decrypt import decrypt_v4, decrypt_v3

    if not is_admin():
        print('提示:当前不是管理员权限,可能读取不到微信密钥。')
    ver = input('你的微信版本是 3.x 还是 4.0?(输入 3 或 4):').strip()
    root = os.path.join(app_dir(), 'decrypted')

    if ver == '3':
        with open(resource_path('wxManager/decrypt/version_list.json'), 'r', encoding='utf-8') as f:
            version_list = json.loads(f.read())
        infos = get_info_v3(version_list)
    else:
        infos = get_info_v4()

    found = False
    for wx_info in infos:
        found = True
        print(wx_info)
        me = Me()
        me.wx_dir = wx_info.wx_dir
        me.wxid = wx_info.wxid
        me.name = wx_info.nick_name
        if ver != '3':
            me.xor_key = get_decode_code_v4(wx_info.wx_dir)
        info_data = me.to_json()
        key = wx_info.key
        if not key:
            print('错误:未找到 key,请重启微信并登录后再试')
            continue
        out = os.path.join(root, wx_info.wxid)
        os.makedirs(out, exist_ok=True)
        if ver == '3':
            decrypt_v3.decrypt_db_files(key, src_dir=wx_info.wx_dir, dest_dir=out)
            sub = 'Msg'
        else:
            decrypt_v4.decrypt_db_files(key, src_dir=wx_info.wx_dir, dest_dir=out)
            sub = 'db_storage'
        os.makedirs(os.path.join(out, sub), exist_ok=True)
        with open(os.path.join(out, sub, 'info.json'), 'w', encoding='utf-8') as f:
            json.dump(info_data, f, ensure_ascii=False, indent=4)
        print(f'解密完成:{os.path.join(out, sub)}')
    if not found:
        print('没有找到正在运行的微信,请先登录微信再试。')


# ---------------- 选择已解密的数据库 ----------------
def choose_db():
    root = os.path.join(app_dir(), 'decrypted')
    dbs = []
    if os.path.isdir(root):
        for name in sorted(os.listdir(root)):
            p4 = os.path.join(root, name, 'db_storage')
            p3 = os.path.join(root, name, 'Msg')
            if os.path.isdir(p4):
                dbs.append((name, p4, 4))
            elif os.path.isdir(p3):
                dbs.append((name, p3, 3))
    if not dbs:
        print('没有找到已解密的数据库,请先执行"1 解密"。')
        return None
    if len(dbs) == 1:
        print(f'使用数据库:{dbs[0][0]}')
        return dbs[0]
    for i, d in enumerate(dbs):
        print(f'  {i}: {d[0]} (版本 {d[2]})')
    idx = int(input('选择序号:').strip())
    return dbs[idx]


# ---------------- 2. 联系人 ----------------
def do_contacts():
    from wxManager import DatabaseConnection
    sel = choose_db()
    if not sel:
        return
    _, db_dir, ver = sel
    database = DatabaseConnection(db_dir, ver).get_interface()
    cnt = 0
    for contact in database.get_contacts():
        print(contact)
        cnt += 1
    print(f'联系人/群总数:{cnt}(其中 wxid 用于第 3 步导出)')


# ---------------- 3. 导出 ----------------
def do_export():
    from exporter.config import FileType
    from exporter import (HtmlExporter, TxtExporter, AiTxtExporter,
                          DocxExporter, MarkdownExporter, ExcelExporter)
    from wxManager import DatabaseConnection

    sel = choose_db()
    if not sel:
        return
    _, db_dir, ver = sel
    database = DatabaseConnection(db_dir, ver).get_interface()

    formats = {
        '1': ('HTML', FileType.HTML, HtmlExporter),
        '2': ('TXT', FileType.TXT, TxtExporter),
        '3': ('Markdown', FileType.MARKDOWN, MarkdownExporter),
        '4': ('Excel', FileType.XLSX, ExcelExporter),
        '5': ('Word', FileType.DOCX, DocxExporter),
        '6': ('AI用TXT', FileType.AI_TXT, AiTxtExporter),
    }
    for k, v in formats.items():
        print(f'  {k}: {v[0]}')
    fmt = formats.get(input('选择导出格式:').strip())
    if not fmt:
        print('无效选择')
        return

    target = input('输入要导出的 wxid(输入 all 导出全部联系人):').strip()
    out_dir = os.path.join(app_dir(), 'export') + os.sep
    os.makedirs(out_dir, exist_ok=True)

    if target.lower() == 'all':
        contacts = database.get_contacts()
    else:
        c = database.get_contact_by_username(target)
        if c is None:
            print('没有找到这个 wxid')
            return
        contacts = [c]

    st = time.time()
    for contact in contacts:
        try:
            ex = fmt[2](
                database, contact,
                output_dir=out_dir,
                type_=fmt[1],
                message_types=None,
                time_range=['2000-01-01 00:00:00', '2035-01-01 00:00:00'],
                group_members=None,
            )
            ex.start()
        except Exception:
            print(f'导出失败:{contact}')
            traceback.print_exc()
    print(f'完成,耗时 {time.time() - st:.1f}s,输出目录:{out_dir}')


def main():
    while True:
        print('\n===== WeChatMsg 工具 =====')
        print('1 解密数据库(需微信已登录)')
        print('2 查看联系人')
        print('3 导出聊天记录')
        print('0 退出')
        c = input('请选择:').strip()
        try:
            if c == '1':
                do_decrypt()
            elif c == '2':
                do_contacts()
            elif c == '3':
                do_export()
            elif c == '0':
                break
        except Exception:
            traceback.print_exc()


if __name__ == '__main__':
    freeze_support()  # 多进程打包必须
    main()
