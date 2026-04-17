import pythoncom
import win32com.client as win32
from pathlib import Path
from dagster import ConfigurableResource
import time

class OutlookEmailResource(ConfigurableResource):
    sender_email: str = "hec.comex.1@dksh.com"

    def _load_template(self, config_dir: str):
        """Đọc file Template_mail.txt và tách Header/Body"""
        fp_txt = Path(config_dir)
        if not fp_txt.exists():
            raise FileNotFoundError(f"Không tìm thấy file template tại: {fp_txt}")

        with open(fp_txt, 'r', encoding='utf-8') as f:
            lines = f.read().splitlines()

        headers = {}
        body_lines = []
        is_body = False

        for line in lines:
            if line.strip() == '---':
                is_body = True
                continue

            if not is_body and ':' in line:
                key, value = line.split(':', 1)
                headers[key.strip().lower()] = value.strip()
            elif is_body:
                body_lines.append(line)

        return headers, '\n'.join(body_lines)

    def _replace_placeholder(self, text: str, email_name: str):
        """Thay thế {email_name} trong template bằng giá trị thực tế"""
        if not text:
            return ""
        if email_name is None:
            # Nếu gửi mail tổng, xóa placeholder hoặc thay bằng từ chung chung
            return text.replace("{email_name}", "Team/Sếp")
        return text.replace("{email_name}", str(email_name))

    def _create_mail(self, outlook, headers, body, email_name):
        """Khởi tạo object Mail và điền thông tin"""
        mail = outlook.CreateItem(0)

        # Chọn tài khoản gửi
        account_to_use = None
        for acc in outlook.Session.Accounts:
            if acc.SmtpAddress.lower() == self.sender_email.lower():
                account_to_use = acc
                break
        if account_to_use:
            mail._oleobj_.Invoke(*(64209, 0, 8, 0, account_to_use))

        # Điền Header & Body (đã replace placeholder)
        mail.Subject = self._replace_placeholder(headers.get('subject', ''), email_name)
        mail.To = self._replace_placeholder(headers.get('to', ''), email_name)
        mail.CC = self._replace_placeholder(headers.get('cc', ''), email_name)
        mail.Body = self._replace_placeholder(body, email_name)

        return mail

    def send_report(self, dfs: dict = None, email_list: list = None, config_dir: str = "", bline: str = "", attach_file: str = None):
        headers, body = self._load_template(config_dir)
        pythoncom.CoInitialize()
        
        try:
            # Dùng DispatchEx để tạo instance mới sạch sẽ
            outlook = win32.DispatchEx('Outlook.Application')
            ns = outlook.GetNamespace("MAPI") # Lấy namespace để refresh

            if email_list is None:
                # --- TRƯỜNG HỢP 1: GỬI MAIL TỔNG ---
                mail = self._create_mail(outlook, headers, body, None)
                # Đính kèm file có sẵn (nếu có)
                if attach_file:
                    fp = Path(attach_file)
                    if fp.exists():
                        mail.Attachments.Add(str(fp.absolute()))
                        print(f"[ATTACH] Đã đính kèm file: {fp.name}")
 
                # Đính kèm file từ DataFrames (nếu có)
                if dfs:
                    for key, df_item in dfs.items():
                        file_path = Path(f"{bline}_{key}.xlsm")
                        df_item.to_excel(file_path, index=False)
                        mail.Attachments.Add(str(file_path.absolute()))
                mail.Send()
                del mail # Giải phóng ngay
                ns.SendAndReceive(True) # Ép Outlook đẩy mail đi
                print(f"[SENT] Đã gửi mail tổng thành công.")

            else:
                # --- TRƯỜNG HỢP 2: GỬI THEO DANH SÁCH ---
                for count, email in enumerate(email_list, 1):
                    try:
                        mail = self._create_mail(outlook, headers, body, email)
                        if attach_file:
                            mail.Attachments.Add(str(Path(attach_file).absolute()))
 
                        # Mỗi người nhận file Excel riêng từ DataFrame (nếu có trong dict dfs)
                        if dfs and email in dfs:
                            df_sub = dfs[email]
                            safe_name = email.replace("@", "_").replace(".", "_")
                            file_path = Path(f"{bline}_{safe_name}.xlsm")
                            df_sub.to_excel(file_path, index=False)
                            mail.Attachments.Add(str(file_path.absolute()))
                        
                        mail.Send()
                        print(f"[SENT] {count}. Đã gửi tới: {email}")
                                                
                        del mail
                        if count % 5 == 0:
                            ns.SendAndReceive(True) 
                            time.sleep(2) 
                            
                    except Exception as e:
                        print(f"[ERROR] Lỗi khi gửi cho {email}: {e}")
                ns.SendAndReceive(True)

        finally:
            outlook = None
            pythoncom.CoUninitialize()