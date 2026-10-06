"""學生程式碼回傳。預設只檢查，寄信必須明確選擇 test 或 send。"""
import argparse
import base64
import csv
import hashlib
import io
import json
import os
from pathlib import Path
import re
from email.message import EmailMessage
from datetime import datetime, timezone
import zipfile

ROOT = Path(__file__).resolve().parent
CONFIG = {}
SENDER = ''
SCOPES = ['https://www.googleapis.com/auth/gmail.send']

def load_plan(excel, archive):
    import openpyxl
    workbook = openpyxl.load_workbook(excel, data_only=True)
    rows = list((workbook[CONFIG['worksheet']] if CONFIG.get('worksheet') else workbook.active).values)
    if not rows:
        raise ValueError('Excel 為空')
    headers = [str(v).strip() for v in rows[0]]
    indexes = [headers.index(h) for h in ('學號', '姓名', 'Email')]
    students, ids, emails = [], set(), set()
    for row in rows[1:]:
        if not any(v is not None for v in row):
            continue
        sid, name, email = [str(row[i] or '').strip() for i in indexes]
        sid = sid.upper()
        if not re.fullmatch(r'S\d{8}', sid) or not name:
            raise ValueError(f'學號或姓名不完整：{sid}')
        if email.lower() != f"{sid}@{CONFIG['student_email_domain']}".lower():
            raise ValueError(f'Email 與學號不符：{sid} / {email}')
        if sid in ids or email.lower() in emails:
            raise ValueError(f'重複學號或 Email：{sid}')
        ids.add(sid)
        emails.add(email.lower())
        students.append({'學號': sid, '姓名': name, 'Email': email})
    with zipfile.ZipFile(archive) as outer:
        matches = {}
        prefix = CONFIG.get('archive_prefix', '').strip('/')
        pattern = r'^' + (re.escape(prefix) + '/' if prefix else '') + r'(S\d{8})(?:_|\.zip)'
        for info in outer.infolist():
            if info.is_dir() or not info.filename.lower().endswith('.zip'):
                continue

            match = re.match(
                pattern,
                info.filename,
                re.I,
            )
            if match:
                matches.setdefault(match[1].upper(), []).append(info)
        if set(matches) != ids:
            raise ValueError(f'附件名單不符。缺少：{ids-set(matches)}；多出：{set(matches)-ids}')
        for student in students:
            sid = student['學號']
            if len(matches[sid]) != 1:
                raise ValueError(f'{sid} 有多份 ZIP，請先確認')
            data = outer.read(matches[sid][0])
            if len(data) > 15 * 1024 * 1024:
                raise ValueError(f'{sid} 附件超過本工具的 15 MiB 上限')
            with zipfile.ZipFile(io.BytesIO(data)) as inner:
                if inner.testzip() or not inner.namelist():
                    raise ValueError(f'{sid} 附件損壞或為空')
            student['data'] = data
            student['附件'] = f'{sid}.zip'
            student['SHA256'] = hashlib.sha256(data).hexdigest()
    if not students:
        raise ValueError('學生名單為空')
    return students

def make_message(student, test=False):
    message = EmailMessage()
    message['From'] = SENDER
    message['To'] = SENDER if test else student['Email']
    message['Subject'] = ('【測試】' if test else '') + CONFIG['subject']
    text = CONFIG['body'].format(name=student['姓名'], student_id=student['學號'], assignment_name=CONFIG['assignment_name'])
    if test:
        text = f"這是測試信，未寄給學生。\n原收件人：{student['Email']}\n學號：{student['學號']}\n\n" + text
    message.set_content(text)
    message.add_attachment(student['data'], maintype='application', subtype='zip', filename=student['附件'])
    return message

def gmail_service():
    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
    from googleapiclient.discovery import build
    token = resolve_path(CONFIG.get('token_file', 'token.json'))
    creds = Credentials.from_authorized_user_file(str(token), SCOPES) if token.exists() else None
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            path = resolve_path(CONFIG.get('credentials_file', 'credentials.json'))
            if not path.exists():
                raise ValueError('請先放入 Google Cloud 的 Desktop app credentials.json，參考使用說明。')
            flow = InstalledAppFlow.from_client_secrets_file(str(path), SCOPES)
            creds = flow.run_local_server(port=0, login_hint=SENDER, prompt='select_account')
        token.parent.mkdir(parents=True, exist_ok=True)
        token.write_text(creds.to_json(), encoding='utf-8')
    return build('gmail', 'v1', credentials=creds, cache_discovery=False)

def resolve_path(value):
    path = Path(value).expanduser()
    return path if path.is_absolute() else ROOT / path

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('mode', nargs='?', choices=['check', 'test', 'send'], default='check')
    parser.add_argument('--config', default='config.json', help='設定檔路徑，相對程式資料夾')
    parser.add_argument('--student', help='測試使用的學號，預設名單第一位')
    args = parser.parse_args()
    global CONFIG, SENDER
    CONFIG = json.loads(resolve_path(args.config).read_text(encoding='utf-8-sig'))
    required = ['assignment_id', 'assignment_name', 'sender', 'excel_file', 'archive_file', 'student_email_domain', 'subject', 'body']
    if any(not CONFIG.get(k) for k in required):
        raise ValueError('設定檔缺少必填內容')
    if not re.fullmatch(r'[A-Za-z0-9_-]+', CONFIG['assignment_id']):
        raise ValueError('assignment_id 只能使用英文字母、數字、底線與減號')
    SENDER = CONFIG['sender'].strip()
    if not re.fullmatch(r'[^\s@]+@[^\s@]+\.[^\s@]+', SENDER):
        raise ValueError('寄件信箱格式錯誤')
    output = ROOT / 'outputs' / CONFIG['assignment_id']
    output.mkdir(parents=True, exist_ok=True)
    # 排他鎖避免兩個程式同時寄送；異常終止時需人工確認後刪除。
    lock = output / 'running.lock'
    try:
        fd = os.open(lock, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
    except FileExistsError:
        raise ValueError('偵測到 running.lock。確認沒有另一個程式執行後，才可刪除這個鎖檔。')
    os.close(fd)
    try:
        plan = load_plan(resolve_path(CONFIG['excel_file']), resolve_path(CONFIG['archive_file']))
        with (output / '配對清單.csv').open('w', encoding='utf-8-sig', newline='') as f:
            writer = csv.DictWriter(f, fieldnames=['學號', '姓名', 'Email', '附件', 'SHA256'], extrasaction='ignore')
            writer.writeheader()
            writer.writerows(plan)
        print(f'配對成功：{len(plan)} 位。清單位於 {output / '配對清單.csv'}')
        fingerprint = hashlib.sha256(json.dumps({'sender': SENDER, 'subject': CONFIG['subject'], 'body': CONFIG['body'], 'assignment_name': CONFIG['assignment_name'], 'plan': [{k: v for k, v in st.items() if k != 'data'} for st in plan]}, sort_keys=True, ensure_ascii=False).encode()).hexdigest()
        if args.mode == 'check':
            print('僅檢查，沒有寄信。')
            return
        ledger_path = output / '寄送紀錄.json'
        ledger = json.loads(ledger_path.read_text(encoding='utf-8')) if ledger_path.exists() else {}
        legacy = ROOT / 'outputs' / '寄送紀錄.json'
        if args.mode == 'send' and legacy.exists() and not ledger:
            raise ValueError('找到舊版寄送紀錄，請先依 README 遷移，避免重寄。')
        for st in plan:
            previous = ledger.get('send:' + st['學號'])
            if previous and (previous['to'].lower() != st['Email'].lower() or previous['sha256'] != st['SHA256']):
                raise ValueError(f"{st['學號']} 的資料與既有紀錄不同，請人工確認；不自動重寄")
        def save():
            temp = ledger_path.with_suffix('.tmp')
            temp.write_text(json.dumps(ledger, ensure_ascii=False, indent=2), encoding='utf-8')
            os.replace(temp, ledger_path)
        if args.mode == 'test':
            selected = [s for s in plan if not args.student or s['學號'] == args.student.upper()][:1]
            if not selected:
                raise ValueError('測試學號不在名單內')
            key = 'test:' + fingerprint + ':' + selected[0]['學號']
            print(f'測試信只會寄到 {SENDER}，附件為 {selected[0]["附件"]}')
        else:
            if not any(k.startswith('test:' + fingerprint + ':') and v['status'] == 'sent' for k, v in ledger.items()):
                raise ValueError('請先成功寄送測試信，並檢查收到的附件。')
            selected = [s for s in plan if 'send:' + s['學號'] not in ledger]
            print(f'待寄送 {len(selected)} 位，已嘗試寄送者會略過（包括結果不確定者）。')
            if not selected:
                return
            if input('確認測試信與配對清單正確後，輸入 SEND 開始寄送：').strip() != 'SEND':
                print('已取消。')
                return
        service = gmail_service()
        for student in selected:
            key = ('test:' + fingerprint + ':' + student['學號']) if args.mode == 'test' else ('send:' + student['學號'])
            if key in ledger:
                print('此信已嘗試寄送，請查閱紀錄與 Gmail 寄件備份，不自動重寄。')
                continue
            message = make_message(student, args.mode == 'test')
            raw = base64.urlsafe_b64encode(message.as_bytes()).decode('ascii')
            ledger[key] = {'status': 'pending', 'to': str(message['To']), 'student': student['學號'], 'sha256': student['SHA256'], 'time': datetime.now(timezone.utc).isoformat()}
            save()
            try:
                result = service.users().messages().send(userId=SENDER, body={'raw': raw}).execute(num_retries=0)
                ledger[key].update(status='sent', message_id=result['id'])
                save()
                print(f'已寄送：{student["學號"]} → {message["To"]}')
            except Exception:
                ledger[key]['status'] = 'uncertain'
                save()
                print('寄送中斷，結果未確認。請先查 Gmail 寄件備份，再處理紀錄；不自動重試。')
                raise
    finally:
        lock.unlink(missing_ok=True)

if __name__ == '__main__':
    try:
        main()
    except Exception as exc:
        print(f'停止：{exc}')
        raise SystemExit(1)
