# TA Mail Automation

使用 Python 與 Gmail API，依學號配對 Excel 名單和個人 ZIP，將學生程式碼各別寄回本人。支援寄前檢查、寄給自己測試及逐封寄送紀錄。

## 安裝與範例檢查

需要 Python 3.10 以上。

```powershell
py -m pip install -r requirements.txt
Copy-Item config.example.json config.json
py main.py check
```

內附兩位虛構學生，檢查模式不會登入或寄信。命令在 main.py 所在資料夾執行；設定內相對路徑也以這個資料夾為基準。

## 每次作業設定

編輯本機 config.json：

| 設定 | 用途 |
|---|---|
| assignment_id | 該次作業唯一編號，例如 2026-oop-exam-01；續寄必須沿用同一編號 |
| assignment_name | 信件中的作業名稱 |
| sender | 實際寄件 Gmail / Google Workspace 帳號 |
| excel_file | 學生 Excel 路徑 |
| archive_file | 包含個人 ZIP 的總壓縮檔 |
| archive_prefix | 總 ZIP 中個人附件所在的前綴，例如 0929；個人 ZIP 在最外層就填空字串 |
| student_email_domain | 學生信箱網域，例如 stumail.nutn.edu.tw |
| worksheet | 工作表名稱；空字串表示作用中工作表 |
| subject | 信件主旨 |
| body | 信件內文，支援 {name}、{student_id}、{assignment_name} |
| credentials_file / token_file | 憑證與授權檔路徑；更換寄件帳號需指定另一份 token 檔重新授權 |

學生欄位必須為「學號」「姓名」「Email」。學號採 S 加 8 位數字，Email 必須為該學號加設定網域。這是目前的校內使用規則，尚未支援任意學號及信箱格式。

附件名稱為 Sxxxxxxxx.zip 或 Sxxxxxxxx_姓名.zip。程式直接讀取個人 ZIP，不重新執行或改寫學生程式碼。姓名編碼即使含錯誤斜線，仍從設定前綴後的學號辨識；同一學號找到多個 ZIP 時停止，不任選一份。

真實資料請放在 data/、submissions/，避免提交 GitHub。

## Gmail 授權與寄送

Google 官方設定說明：
https://developers.google.com/workspace/gmail/api/quickstart/python

建立 Google Cloud 專案、啟用 Gmail API、設定 OAuth 同意畫面。External 且 Testing 時加入實際寄件帳號為測試使用者；建立 Desktop app OAuth client 並下載 credentials.json。程式只要求 gmail.send 權限。

```powershell
py main.py check
py main.py test
# 也可指定測試附件：
py main.py test --student S00000002
# 確認自己收到的信與附件正確後：
py main.py send
```

正式寄送需輸入 SEND。test 永遠寄到設定的 sender；正式寄送每位學生只有自己的附件。API 接受寄件不等於對方已收件，另行查看退信。

另一份設定可使用：

```powershell
py main.py check --config config.exam02.local.json
```

## 寄送紀錄與續寄

紀錄位於 outputs/<assignment_id>/寄送紀錄.json，配對清單也在同一資料夾。

- pending：呼叫 API 前已保存。
- sent：API 回傳成功及信件 ID。
- uncertain：呼叫期間失敗，寄件結果不確定。

同一作業重跑時，所有已嘗試的學生都會略過。pending / uncertain 需先人工查看 Gmail 寄件備份，確認未寄出才調整該筆紀錄；不要刪整份紀錄重跑。中斷留下 running.lock 時，先確認沒有另一個程式執行，再刪鎖檔。

測試紀錄綁定完整配對清單、寄件帳號及信件內容；修改以上內容需重新測試。若已寄學生的 Email 或附件變更，程式會停止並要求人工確認。

## 從舊版遷移（已寄過第一次上機考）

不要再次寄送已完成的考試，也不要刪舊版紀錄。

1. 先備份整個舊資料夾。
2. 給舊考試固定 assignment_id，例如 2026-oop-exam-01。
3. 把舊 outputs/寄送紀錄.json **複製**到新 outputs/2026-oop-exam-01/寄送紀錄.json。
4. 把真實 Excel 與總 ZIP 放到本機 data/，設定 archive_prefix 為 0929、student_email_domain 為 stumail.nutn.edu.tw，worksheet 依 Excel 填寫。
5. 原 credentials.json、token.json 留在本機；設定指向原檔即可。
6. 執行 check。已寄資料仍保留 send:<學號>，新版會略過；舊測試紀錄不會當作新版本的測試。

下一次作業用新的 assignment_id，使用新的檔案與內文。若程式在根目錄找到舊紀錄且該次作業尚無紀錄，會阻止正式寄送；確認舊紀錄已備份遷移後，把根目錄舊紀錄移到備份位置。

## GitHub 上傳

本包不包含真實名單、学生附件、個人信箱或登入憑證。勿把舊專案整個複製進新專案。

```powershell
git init
git add main.py requirements.txt config.example.json README.md .gitignore examples tests
git diff --cached --stat
git commit -m "Add reusable TA mail automation tool"
```

建立 GitHub 空 repository（不要先建立 README），取得網址後：

```powershell
git branch -M main
git remote add origin https://github.com/YOUR_USERNAME/ta-mail-automation.git
git push -u origin main
```

若敏感檔案曾經被 git 追蹤，.gitignore 不會清除歷史。本指令適用於這份新的乾淨專案。

## 測試與限制

```powershell
py -m unittest discover -s tests -v
```

測試不登入或寄送真實信件。每份個人 ZIP 限 15 MiB；全部附件目前載入記憶體，適合小班級。未實作退信追蹤、排程、GUI 或自動重試。OAuth 授權仍受學校管理政策影響。
