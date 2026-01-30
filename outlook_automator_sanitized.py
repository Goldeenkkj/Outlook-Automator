import os
import win32com.client
import re
from PyPDF2 import PdfWriter, PdfReader
import time
import pythoncom
from weasyprint import HTML
from bs4 import BeautifulSoup
from PIL import Image

class OutlookAutomator:
    def __init__(self):
        """Inicializa o automator do Outlook, conectando-se à aplicação Outlook."""
        pythoncom.CoInitialize()
        self.outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

        # CONFIGURAÇÃO: Substitua pelo nome da sua caixa compartilhada
        shared_mailbox_name = os.getenv("SHARED_MAILBOX", "example@company.com")
        inbox_folder_name = os.getenv("INBOX_FOLDER", "Caixa de Entrada")

        shared_mailbox = self.outlook.Folders.Item(shared_mailbox_name)
        self.inbox = shared_mailbox.Folders.Item(inbox_folder_name)

    def get_unread_emails(self):
        """Filtra e retorna apenas os emails não lidos da caixa de entrada."""
        start_time = time.time()
        unread_filter = "[Unread] = True"
        messages = self.inbox.Items.Restrict(unread_filter)
        end_time = time.time()
        duration = end_time - start_time
        print(f"[LOG] Encontrados {len(messages)} emails não lidos em {duration:.2f} segundos.")
        return messages

    def _extract_and_save_embedded_images(self, message, email_folder, html_body):
        """Extrai imagens incorporadas (cid:) do HTML e as salva localmente."""
        soup = BeautifulSoup(html_body, 'html.parser')

        for img_tag in soup.find_all('img', src=re.compile(r'^cid:')):
            cid = img_tag["src"][4:]
            print(f"[LOG] CID encontrado: {cid}")

            for attachment in message.Attachments:
                attachment_cid = None
                try:
                    attachment_cid = attachment.PropertyAccessor.GetProperty(
                        "http://schemas.microsoft.com/mapi/proptag/0x3712001E"
                    )
                except Exception:
                    pass

                if attachment_cid == cid:
                    image_filename = attachment.FileName
                    image_path = os.path.join(email_folder, image_filename)
                    try:
                        attachment.SaveAsFile(image_path)
                        print(f"[LOG] Imagem incorporada salva: {image_path}")
                        img_tag["src"] = f"file:///{image_path}"
                        break
                    except Exception as e:
                        print(f"[LOG] Erro ao salvar imagem: {e}")

        return str(soup)

    def _convert_image_to_pdf(self, image_path, pdf_path):
        """Converte um arquivo de imagem para PDF."""
        try:
            img = Image.open(image_path)
            if img.mode == 'RGBA':
                img = img.convert('RGB')
            img.save(pdf_path, "PDF", resolution=100.0)
            print(f"[LOG] Imagem convertida para PDF: {pdf_path}")
            return True
        except Exception as e:
            print(f"[LOG] Erro ao converter imagem: {e}")
            return False

    def _merge_pdfs(self, pdf_list, output_path):
        """Mescla múltiplos PDFs em um único arquivo."""
        try:
            writer = PdfWriter()
            for pdf_path in pdf_list:
                try:
                    reader = PdfReader(pdf_path)
                    for page in reader.pages:
                        writer.add_page(page)
                except Exception as e:
                    print(f"[LOG] Erro ao ler PDF {pdf_path}: {e}")

            with open(output_path, 'wb') as output_file:
                writer.write(output_file)
            print(f"[LOG] PDF consolidado criado: {output_path}")
        except Exception as e:
            print(f"[LOG] Erro ao mesclar PDFs: {e}")

    def process_emails(self, output_folder):
        """Processa todos os emails não lidos e cria PDFs consolidados."""
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)

        messages = self.get_unread_emails()

        for message in messages:
            try:
                subject = message.Subject
                cleaned_subject = re.sub(r'[<>:"/\\|?*]', '_', subject)[:100]
                timestamp = message.SentOn.strftime("%Y-%m-%d_%H-%M")
                email_folder = os.path.join(output_folder, f"{timestamp}_{cleaned_subject}")

                if not os.path.exists(email_folder):
                    os.makedirs(email_folder)

                pdfs_to_merge = []

                # Processar corpo do email
                try:
                    html_body = message.HTMLBody
                    html_body = self._extract_and_save_embedded_images(message, email_folder, html_body)
                    email_pdf_path = os.path.join(email_folder, f"{cleaned_subject}_email.pdf")
                    HTML(string=html_body).write_pdf(email_pdf_path)
                    pdfs_to_merge.append(email_pdf_path)
                except Exception as e:
                    print(f"[LOG] Erro ao processar corpo HTML: {e}")
                    try:
                        text_body = message.Body
                        email_pdf_path = os.path.join(email_folder, f"{cleaned_subject}_email.pdf")
                        HTML(string=f"<pre>{text_body}</pre>").write_pdf(email_pdf_path)
                        pdfs_to_merge.append(email_pdf_path)
                    except Exception as e2:
                        print(f"[LOG] Erro ao processar texto: {e2}")

                # Processar anexos
                if message.Attachments.Count > 0:
                    for attachment in message.Attachments:
                        attachment_filename = attachment.FileName.strip()
                        attachment_lower = attachment_filename.lower()

                        if attachment_lower.startswith("image00"):
                            continue

                        from unicodedata import normalize
                        safe_filename = normalize("NFKD", attachment_filename).encode("ascii", "ignore").decode("ascii")
                        attachment_path = os.path.join(email_folder, safe_filename)

                        try:
                            time.sleep(1)
                            attachment.SaveAsFile(attachment_path)

                            if not os.path.exists(attachment_path) or os.path.getsize(attachment_path) == 0:
                                print(f"[LOG] Anexo salvo com 0 bytes")
                                continue

                            if attachment_lower.endswith(".pdf"):
                                try:
                                    PdfReader(attachment_path)
                                    pdfs_to_merge.append(attachment_path)
                                except Exception:
                                    print(f"[LOG] PDF inválido")

                            elif attachment_lower.endswith((".jpg", ".jpeg", ".png", ".gif", ".bmp")):
                                image_pdf_path = os.path.join(email_folder, f"{os.path.splitext(safe_filename)[0]}.pdf")
                                if self._convert_image_to_pdf(attachment_path, image_pdf_path):
                                    pdfs_to_merge.append(image_pdf_path)

                        except Exception as e:
                            print(f"[LOG] Erro ao processar anexo: {e}")

                # Consolidar PDFs
                consolidated_pdf_path = os.path.join(email_folder, f"{cleaned_subject}_consolidado.pdf")

                if pdfs_to_merge:
                    self._merge_pdfs(pdfs_to_merge, consolidated_pdf_path)

                    for p_file in pdfs_to_merge:
                        try:
                            os.remove(p_file)
                        except Exception:
                            pass

                message.UnRead = False
                print(f"[LOG] Email processado: {message.Subject}\n")

            except Exception as e:
                print(f"[LOG] Erro ao processar email: {e}\n")

if __name__ == "__main__":
    output_directory = os.getenv("OUTPUT_DIRECTORY", r"C:\Temp\Outlook_PDFs")
    automator = OutlookAutomator()
    automator.process_emails(output_directory)
