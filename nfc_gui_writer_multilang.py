
import nfc
import tkinter as tk
from tkinter import filedialog, messagebox
from openpyxl import load_workbook
import threading

languages = {
    'zh': {
        'title': 'NFC 名片批次寫入工具',
        'select_file': '請選擇 Excel 檔案：',
        'choose': '選擇檔案',
        'start': '開始批次寫入 NFC',
        'success': '✅ 寫入完成，請更換下一張標籤，按 Enter 繼續...',
        'fail': '⚠️ 寫入失敗',
        'scan': '📇 請將 NFC 標籤靠近：',
        'error': '錯誤',
        'lang_switch': '切換語言'
    },
    'en': {
        'title': 'NFC vCard Batch Writer',
        'select_file': 'Select Excel file:',
        'choose': 'Browse',
        'start': 'Start Writing NFC',
        'success': '✅ Write successful. Replace tag and press Enter to continue...',
        'fail': '⚠️ Write failed',
        'scan': '📇 Hold NFC tag near reader for: ',
        'error': 'Error',
        'lang_switch': 'Switch Language'
    }
}

current_lang = 'zh'

def generate_vcard(name, title, phone, email, company, address, website):
    return f"""BEGIN:VCARD
VERSION:3.0
FN:{name}
ORG:{company}
TITLE:{title}
TEL:{phone}
EMAIL:{email}
ADR:{address}
URL:{website}
END:VCARD"""

def write_to_tag(vcard_text, log_text):
    try:
        clf = nfc.ContactlessFrontend('usb')
        def connected(tag):
            if tag.ndef:
                record = nfc.ndef.TextRecord(vcard_text)
                tag.ndef.records = [record]
                return True
            return False
        clf.connect(rdwr={'on-connect': connected})
        clf.close()
        return True
    except Exception as e:
        log_text.insert(tk.END, f"❌ {str(e)}\n")
        log_text.update()
        return False

def start_writing(file_path, log_text, strings):
    try:
        wb = load_workbook(filename=file_path)
        sheet = wb.active
        for idx, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=2):
            name, title, phone, email, company, address, website = row
            vcard = generate_vcard(name, title, phone, email, company, address, website)
            log_text.insert(tk.END, f"{strings['scan']}{name} (row {idx})...\n")
            log_text.update()
            result = write_to_tag(vcard, log_text)
            if result:
                log_text.insert(tk.END, f"{strings['success']}\n\n")
            else:
                log_text.insert(tk.END, f"{strings['fail']}\n\n")
            log_text.update()
            input(">> Replace tag then press Enter to continue.")
    except Exception as e:
        messagebox.showerror(strings['error'], str(e))

def browse_file(entry):
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    entry.delete(0, tk.END)
    entry.insert(0, file_path)

def run_gui():
    global current_lang
    window = tk.Tk()
    window.geometry("700x460")

    def refresh_ui():
        window.title(languages[current_lang]['title'])
        label.config(text=languages[current_lang]['select_file'])
        button_browse.config(text=languages[current_lang]['choose'])
        button_start.config(text=languages[current_lang]['start'])
        button_lang.config(text=languages[current_lang]['lang_switch'])

    def switch_language():
        global current_lang
        current_lang = 'en' if current_lang == 'zh' else 'zh'
        refresh_ui()

    label = tk.Label(window)
    label.pack(pady=5)

    entry = tk.Entry(window, width=60)
    entry.pack(pady=5)

    button_browse = tk.Button(window, command=lambda: browse_file(entry))
    button_browse.pack(pady=5)

    log_text = tk.Text(window, height=15)
    log_text.pack(pady=10)

    def start_thread():
        log_text.delete(1.0, tk.END)
        strings = languages[current_lang]
        threading.Thread(target=start_writing, args=(entry.get(), log_text, strings), daemon=True).start()

    button_start = tk.Button(window, bg="green", fg="white", command=start_thread)
    button_start.pack(pady=5)

    button_lang = tk.Button(window, command=switch_language)
    button_lang.pack(pady=5)

    refresh_ui()
    window.mainloop()

if __name__ == "__main__":
    run_gui()
