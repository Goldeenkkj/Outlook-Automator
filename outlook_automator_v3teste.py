import os
import win32com.client
import re
from PyPDF2 import PdfWriter, PdfReader
import time
import pythoncom
from weasyprint import HTML
from bs4 import BeautifulSoup
import base64
import mimetypes
from PIL import Image
import email
from email import policy
from email.header import decode_header

class OutlookAutomator:
    def __init__(self):
        """Inicializa o automator do Outlook, conectando-se à aplicação Outlook."""
        pythoncom.CoInitialize()
        self.outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        shared_mailbox = self.outlook.Folders.Item("pagamentos@farmausa.com")
        self.inbox = shared_mailbox.Folders.Item("Caixa de Entrada")


    def get_unread_emails(self):
        """Filtra e retorna apenas os emails não lidos da caixa de entrada."""
        start_time = time.time()
        unread_filter = "[Unread] = True"
        messages = self.inbox.Items.Restrict(unread_filter)
        end_time = time.time()
        duration = end_time - start_time
        print(f"[LOG] Encontrados {len(messages)} emails não lidos em {duration:.2f} segundos.")
        return messages

    def save_attachments(self, message, folder_path):
        """Salva todos os anexos de um email em uma pasta especificada."""
        print(f"[LOG] Tentando salvar anexos para o email \'{message.Subject}\' na pasta: {folder_path}")
        if not os.path.exists(folder_path):
            print(f"[LOG] Erro: Pasta de anexos não existe ou não pôde ser criada: {folder_path}")
            return

        if message.Attachments.Count > 0:
            for attachment in message.Attachments:
                attachment_path = os.path.join(folder_path, attachment.FileName)
                print(f"[LOG] Tentando salvar anexo \'{attachment.FileName}\' em: {attachment_path}")
                try:
                    attachment.SaveAsFile(attachment_path)
                    print(f"[LOG] Anexo \'{attachment.FileName}\' salvo com sucesso em: {attachment_path}")
                except Exception as e:
                    print(f"[LOG] Erro ao salvar anexo {attachment.FileName} em {attachment_path}: {e}")
        else:
            print("[LOG] Nenhum anexo encontrado para este email.")

    def _extract_and_save_embedded_images(self, message, email_folder, html_body):
        """Extrai imagens incorporadas (cid:) do HTML e as salva localmente, atualizando o HTML para referenciá-las."""
        soup = BeautifulSoup(html_body, 'html.parser')
        for img_tag in soup.find_all('img', src=re.compile(r'^cid:')):
            cid = img_tag["src"][4:]  # Remove 'cid:'
            print(f"[LOG] CID encontrado: {cid}")
            for attachment in message.Attachments:
                attachment_cid = None
                try:
                    # Tenta obter o Content-ID do anexo
                    attachment_cid = attachment.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E" )
                    print(f"[LOG] Anexo \'{attachment.FileName}\' tem CID: {attachment_cid}")
                except Exception:
                    # Se não tiver Content-ID, não é uma imagem incorporada por CID
                    pass # Ignora anexos sem Content-ID ou com erro ao acessá-lo

                if attachment_cid == cid:
                    print(f"[LOG] Correspondência encontrada para CID: {cid} com anexo: {attachment.FileName}")
                    image_filename = attachment.FileName
                    image_path = os.path.join(email_folder, image_filename)
                    try:
                        attachment.SaveAsFile(image_path)
                        print(f"[LOG] Imagem incorporada salva: {image_path}")
                        # Atualizar o src da tag <img> para o caminho local
                        img_tag["src"] = f"file:///{image_path}"
                        break  # Sai do loop de anexos assim que a imagem for encontrada e salva
                    except Exception as e:
                        print(f"[LOG] Erro ao salvar imagem incorporada {image_filename}: {e}")
        return str(soup)

    def _convert_image_to_pdf(self, image_path, pdf_path):
        """Converte um arquivo de imagem para PDF."""
        try:
            img = Image.open(image_path)
            if img.mode == 'RGBA':
                img = img.convert('RGB')
            img.save(pdf_path, "PDF", resolution=100.0)
            print(f"[LOG] Imagem \'{image_path}\' convertida para PDF: {pdf_path}")
            return True
        except Exception as e:
            print(f"[LOG] Erro ao converter imagem \'{image_path}\' para PDF: {e}")
            return False

    def _process_eml_attachment(self, eml_path, output_folder):
        """Processa um anexo .eml, extrai seu conteúdo e anexos, e os converte para PDF.
        Retorna o caminho do PDF consolidado do .eml, se gerado."""
        print(f"[LOG] Processando anexo .eml: {eml_path}")
        try:
            with open(eml_path, 'rb') as f:
                msg = email.message_from_binary_file(f, policy=policy.default)

            # Extrair informações do cabeçalho do .eml
            eml_subject_decoded = decode_header(msg['Subject'])[0][0]
            eml_subject = eml_subject_decoded.decode() if isinstance(eml_subject_decoded, bytes) else eml_subject_decoded
            eml_from_decoded = decode_header(msg['From'])[0][0]
            eml_from = eml_from_decoded.decode() if isinstance(eml_from_decoded, bytes) else eml_from_decoded
            eml_to_decoded = decode_header(msg['To'])[0][0]
            eml_to = eml_to_decoded.decode() if isinstance(eml_to_decoded, bytes) else eml_to_decoded
            eml_date_decoded = decode_header(msg['Date'])[0][0]
            eml_date = eml_date_decoded.decode() if isinstance(eml_date_decoded, bytes) else eml_date_decoded

            eml_info = f"""
                <p><b>De:</b> {eml_from}</p>
                <p><b>Enviado em:</b> {eml_date}</p>
                <p><b>Para:</b> {eml_to}</p>
                <p><b>Assunto:</b> {eml_subject}</p>
            """
            eml_attachments_names = []

            eml_html_body = ""
            eml_text_body = ""
            eml_parts_pdfs = [] # PDFs gerados a partir das partes do .eml

            for part in msg.walk():
                content_type = part.get_content_type()
                disposition = part.get("Content-Disposition")

                if content_type == "text/html" and disposition is None:
                    eml_html_body = part.get_payload(decode=True).decode()
                elif content_type == "text/plain" and disposition is None and not eml_html_body:
                    eml_text_body = part.get_payload(decode=True).decode()
                elif disposition is not None and part.get_filename():
                    filename = part.get_filename()
                    eml_attachments_names.append(filename)
                    eml_attachment_path = os.path.join(output_folder, filename) # Salva anexo do .eml na pasta do email principal
                    try:
                        with open(eml_attachment_path, 'wb') as apf:
                            apf.write(part.get_payload(decode=True))
                        print(f"[LOG] Anexo de .eml \'{filename}\' salvo em: {eml_attachment_path}")

                        # Converter anexos do .eml para PDF se forem imagens ou PDFs
                        if filename.lower().endswith('.pdf'):
                            eml_parts_pdfs.append(eml_attachment_path)
                        elif filename.lower().endswith(('.jpg', '.jpeg', '.png', '.gif', '.bmp')):
                            eml_image_pdf_path = os.path.join(output_folder, f'{os.path.splitext(filename)[0]}.pdf')
                            if self._convert_image_to_pdf(eml_attachment_path, eml_image_pdf_path):
                                eml_parts_pdfs.append(eml_image_pdf_path)

                    except Exception as e:
                        print(f"[LOG] Erro ao salvar anexo de .eml \'{filename}\' em {eml_attachment_path}: {e}")
            
            if eml_attachments_names:
                eml_info += f"<p><b>Anexos do EML:</b> {'; '.join(eml_attachments_names)}</p>"
            eml_info += "<hr>"

            final_eml_html = f"<div>{eml_info}{eml_html_body if eml_html_body else f'<pre>{eml_text_body}</pre>'}</div>"
            eml_pdf_path = os.path.join(output_folder, f"{eml_subject}_eml.pdf")
            self.convert_html_to_pdf(final_eml_html, eml_pdf_path)
            eml_parts_pdfs.insert(0, eml_pdf_path) # Adiciona o PDF do corpo do .eml primeiro

            if eml_parts_pdfs:
                consolidated_eml_pdf_path = os.path.join(output_folder, f"{eml_subject}_eml_consolidado.pdf")
                self._merge_pdfs(eml_parts_pdfs, consolidated_eml_pdf_path)
                # Remover PDFs individuais do .eml após a mesclagem
                for p_file in eml_parts_pdfs:
                    try:
                        os.remove(p_file)
                    except Exception as e:
                        print(f"[LOG] Erro ao remover PDF individual de .eml {p_file}: {e}")
                return consolidated_eml_pdf_path
            return None
        except Exception as e:
            print(f"[LOG] Erro ao processar anexo .eml {eml_path}: {e}")
            return None

    def _process_msg_attachment(self, msg_path, output_folder):
        """Processa um anexo .msg, extrai seu conteúdo e anexos, e os converte para PDF.
        Retorna o caminho do PDF consolidado do .msg, se gerado."""
        print(f"[LOG] Processando anexo .msg: {msg_path}")
        try:
            outlook_app = win32com.client.Dispatch("Outlook.Application")
            msg_item = outlook_app.CreateItemFromTemplate(msg_path)

            msg_subject = msg_item.Subject
            msg_from = msg_item.SenderName
            msg_to = msg_item.To
            msg_sent_on = msg_item.SentOn.strftime("%A, %d de %B de %Y %H:%M")

            msg_info = f"""
                <p><b>De:</b> {msg_from}</p>
                <p><b>Enviado em:</b> {msg_sent_on}</p>
                <p><b>Para:</b> {msg_to}</p>
                <p><b>Assunto:</b> {msg_subject}</p>
            """
            msg_attachments_names = []
            msg_html_body = msg_item.HTMLBody
            msg_parts_pdfs = []

            if msg_item.Attachments.Count > 0:
                for attachment in msg_item.Attachments:
                    filename = attachment.FileName
                    msg_attachments_names.append(filename)
                    msg_attachment_path = os.path.join(output_folder, filename)
                    try:
                        attachment.SaveAsFile(msg_attachment_path)
                        print(f"[LOG] Anexo de .msg \'{filename}\' salvo em: {msg_attachment_path}")

                        if filename.lower().endswith(".pdf"):
                            msg_parts_pdfs.append(msg_attachment_path)
                        elif filename.lower().endswith(('.jpg', '.jpeg', '.png', '.gif', '.bmp')):
                            msg_image_pdf_path = os.path.join(output_folder, f'{os.path.splitext(filename)[0]}.pdf')
                            if self._convert_image_to_pdf(msg_attachment_path, msg_image_pdf_path):
                                msg_parts_pdfs.append(msg_image_pdf_path)
                        elif filename.lower().endswith(".eml"):
                            eml_consolidated_pdf = self._process_eml_attachment(msg_attachment_path, output_folder)
                            if eml_consolidated_pdf:
                                msg_parts_pdfs.append(eml_consolidated_pdf)

                    except Exception as e:
                        print(f"[LOG] Erro ao salvar anexo de .msg \'{filename}\' em {msg_attachment_path}: {e}")
            
            if msg_attachments_names:
                msg_info += f"<p><b>Anexos do MSG:</b> {'; '.join(msg_attachments_names)}</p>"
            msg_info += "<hr>"

            final_msg_html = f"<div>{msg_info}{msg_html_body}</div>"
            msg_pdf_path = os.path.join(output_folder, f"{msg_subject}_msg.pdf")
            self.convert_html_to_pdf(final_msg_html, msg_pdf_path)
            msg_parts_pdfs.insert(0, msg_pdf_path)

            if msg_parts_pdfs:
                consolidated_msg_pdf_path = os.path.join(output_folder, f"{msg_subject}_msg_consolidado.pdf")
                self._merge_pdfs(msg_parts_pdfs, consolidated_msg_pdf_path)
                for p_file in msg_parts_pdfs:
                    try:
                        os.remove(p_file)
                    except Exception as e:
                        print(f"[LOG] Erro ao remover PDF individual de .msg {p_file}: {e}")
                return consolidated_msg_pdf_path
            return None
        except Exception as e:
            print(f"[LOG] Erro ao processar anexo .msg {msg_path}: {e}")
            return None

    def convert_html_to_pdf(self, html_content, pdf_path):
        """Converte conteúdo HTML para PDF usando WeasyPrint."""
        print(f"[LOG] Tentando converter HTML para PDF em: {pdf_path}")
        try:
            # Garante que o diretório de destino existe
            os.makedirs(os.path.dirname(pdf_path), exist_ok=True)
            HTML(string=html_content, base_url=os.path.dirname(pdf_path)).write_pdf(pdf_path)
            print(f"[LOG] HTML convertido para PDF com sucesso: {pdf_path}")
        except Exception as e:
            print(f"[LOG] Erro ao converter HTML para PDF em {pdf_path}: {e}")

    def _merge_pdfs(self, pdf_list, output_path):
        """Mescla uma lista de arquivos PDF em um único arquivo PDF."""
        print(f"[LOG] Iniciando mesclagem de PDFs. Lista: {pdf_list}")
        writer = PdfWriter()
        for pdf_file in pdf_list:
            try:
                reader = PdfReader(pdf_file)
                for page in reader.pages:
                    writer.add_page(page)
                print(f"[LOG] Adicionado {pdf_file} à mesclagem.")
            except Exception as e:
                print(f"[LOG] Erro ao adicionar {pdf_file} à mesclagem: {e}")
        
        if not writer.pages:
            print(f"[LOG] Aviso: Nenhum PDF válido para mesclar. Nenhum arquivo será criado em {output_path}")
            return

        try:
            with open(output_path, "wb") as f:
                writer.write(f)
            print(f"[LOG] PDFs mesclados com sucesso em: {output_path}")
        except Exception as e:
            print(f"[LOG] Erro ao salvar PDF mesclado em {output_path}: {e}")

    def process_emails(self, output_folder):
        """Processa emails não lidos: extrai HTML, converte para PDF e salva anexos.
        Marca os emails como lidos após o processamento bem-sucedido.
        O script roda em loop contínuo, verificando novos emails a cada 20 segundos.
        """
        while True: # Loop contínuo para verificar novos emails
            print(f"[LOG] Verificando novos emails não lidos...")
            messages = self.get_unread_emails()
            if messages.Count == 0:
                print(f"[LOG] Nenhum email não lido encontrado. Aguardando 20 segundos...")
                time.sleep(20) # Espera 20 segundos antes de verificar novamente
                continue

            for message in list(messages): # Converte para lista para evitar problemas de coleção durante a iteração
                print(f"[LOG] Processando email com assunto: \'{message.Subject}\'")
                # Limpar o assunto do email para criar um nome de arquivo/pasta válido
                cleaned_subject = re.sub(r'[\\/:*?"<>|\t\n\r]', '', message.Subject) # Remove caracteres inválidos e quebras de linha/tabulações
                cleaned_subject = re.sub(r'\s+', ' ', cleaned_subject).strip() # Substitui múltiplos espaços por um único e remove espaços extras
                
                if not cleaned_subject: # Se o assunto ficar vazio após a limpeza, usar um nome padrão
                    cleaned_subject = "Email_Sem_Assunto_" + str(int(time.time()))
                    print(f"[LOG] Assunto limpo vazio, usando nome padrão: {cleaned_subject}")

                email_folder = os.path.join(output_folder, cleaned_subject)
                os.makedirs(email_folder, exist_ok=True)
                print(f"[LOG] Pasta de destino para o email: {email_folder}")

                pdfs_to_merge = []

                try:
                    html_body = message.HTMLBody
                    # Adicionar informações do cabeçalho ao HTML
                    header_info = f"""
                        <p><b>De:</b> {message.SenderName}</p>
                        <p><b>Enviado em:</b> {message.SentOn.strftime("%A, %d de %B de %Y %H:%M")}</p>
                        <p><b>Para:</b> {message.To}</p>
                        <p><b>Assunto:</b> {message.Subject}</p>
                        <p><b>Anexos:</b> {'; '.join([att.FileName for att in message.Attachments]) if message.Attachments.Count > 0 else 'Nenhum'}</p>
                        <hr>
                    """
                    html_body = header_info + html_body

                    # Extrair e salvar imagens incorporadas, atualizando o HTML
                    html_body = self._extract_and_save_embedded_images(message, email_folder, html_body)

                    # Salvar o corpo do email como PDF
                    email_pdf_path = os.path.join(email_folder, f"{cleaned_subject}.pdf")
                    self.convert_html_to_pdf(html_body, email_pdf_path)
                    pdfs_to_merge.append(email_pdf_path)
                except Exception as e:
                    print(f"[LOG] Erro ao processar corpo HTML do email \'{message.Subject}\' : {e}")
                    # Tentar processar corpo de texto se HTML falhar
                    try:
                        text_body = message.Body
                        header_info = f"""
                            De: {message.SenderName}\n
                            Enviado em: {message.SentOn.strftime("%A, %d de %B de %Y %H:%M")}\n
                            Para: {message.To}\n
                            Assunto: {message.Subject}\n
                            Anexos: {'; '.join([att.FileName for att in message.Attachments]) if message.Attachments.Count > 0 else 'Nenhum'}\n
                            --------------------------------------------------------------------------------\n
                        """
                        text_body = header_info + text_body
                        email_txt_path = os.path.join(email_folder, f"{cleaned_subject}.txt")
                        with open(email_txt_path, 'w', encoding='utf-8') as f:
                            f.write(text_body)
                        print(f"[LOG] Corpo de texto do email salvo em: {email_txt_path}")
                        # Converter texto para PDF (simples)
                        email_pdf_path = os.path.join(email_folder, f"{cleaned_subject}.pdf")
                        HTML(string=f"<pre>{text_body}</pre>").write_pdf(email_pdf_path)
                        pdfs_to_merge.append(email_pdf_path)
                    except Exception as e:
                        print(f"[LOG] Erro ao processar corpo de texto do email \'{message.Subject}\' : {e}")

                # Processar anexos
                if message.Attachments.Count > 0:
                    print(f"[LOG] Encontrados {message.Attachments.Count} anexos no email '{message.Subject}'")

                    for attachment in message.Attachments:
                        attachment_filename = attachment.FileName.strip()
                        attachment_lower = attachment_filename.lower()

        # Ignorar imagens inline comuns (geradas pelo Outlook)
                        if attachment_lower.startswith("image00") or attachment_lower.endswith((".gif", ".png", ".jpg", ".jpeg")):
                            print(f"[LOG] Ignorando imagem inline: {attachment_filename}")
                            continue

        # Normalizar nome (evita falhas com acentos ou espaços)
                        from unicodedata import normalize
                        safe_filename = normalize("NFKD", attachment_filename).encode("ascii", "ignore").decode("ascii")
                        attachment_path = os.path.join(email_folder, safe_filename)

                        try:
            # Aguardar um instante para o Outlook carregar totalmente o anexo
                            time.sleep(1)
                            attachment.SaveAsFile(attachment_path)
                            print(f"[LOG] Anexo '{attachment_filename}' salvo em: {attachment_path}")

            # Verificar se o arquivo foi realmente salvo e possui conteúdo
                            if not os.path.exists(attachment_path) or os.path.getsize(attachment_path) == 0:
                                print(f"[LOG] ⚠️ Anexo '{attachment_filename}' salvo com 0 bytes (pode estar corrompido).")
                                continue

            # Tratar anexos conforme o tipo
                            if attachment_lower.endswith(".pdf"):
                # Validar se é um PDF legível
                                try:
                                    from PyPDF2 import PdfReader
                                    PdfReader(attachment_path)
                                    print(f"[LOG] ✅ PDF válido: {attachment_filename}")
                                    pdfs_to_merge.append(attachment_path)
                                except Exception:
                                    print(f"[LOG] ⚠️ Arquivo '{attachment_filename}' não é um PDF válido — ignorado.")
                                    continue

                            elif attachment_lower.endswith((".jpg", ".jpeg", ".png", ".gif", ".bmp")):
                                image_pdf_path = os.path.join(email_folder, f"{os.path.splitext(safe_filename)[0]}.pdf")
                                if self._convert_image_to_pdf(attachment_path, image_pdf_path):
                                    pdfs_to_merge.append(image_pdf_path)

                            elif attachment_lower.endswith(".eml"):
                                eml_consolidated_pdf = self._process_eml_attachment(attachment_path, email_folder)
                                if eml_consolidated_pdf:
                                    pdfs_to_merge.append(eml_consolidated_pdf)

                            elif attachment_lower.endswith(".msg") or 'ms-outlook' in attachment.DisplayName.lower():
                                msg_consolidated_pdf = self._process_msg_attachment(attachment_path, email_folder)
                                if msg_consolidated_pdf:
                                    pdfs_to_merge.append(msg_consolidated_pdf)

                            else:
                                print(f"[LOG] Anexo '{attachment_filename}' ignorado (tipo desconhecido).")

                        except Exception as e:
                            print(f"[LOG] Erro ao salvar ou processar anexo '{attachment_filename}': {e}")
                else:
                    print("[LOG] Nenhum anexo encontrado para este email.")

                # Definir o caminho para o PDF consolidado
                consolidated_pdf_path = os.path.join(email_folder, f"{cleaned_subject}_consolidado.pdf")

                # Mesclar todos os PDFs (email + anexos PDF) em um único arquivo
                if pdfs_to_merge:
                    self._merge_pdfs(pdfs_to_merge, consolidated_pdf_path)
                    # Remover os PDFs individuais após a mesclagem para manter apenas o consolidado
                    for p_file in pdfs_to_merge:
                        try:
                            os.remove(p_file)
                            print(f"[LOG] PDF individual removido após mesclagem: {p_file}")
                        except Exception as e:
                            print(f"[LOG] Erro ao remover PDF individual {p_file}: {e}")
                else:
                    print("[LOG] Nenhum PDF para mesclar (nem email, nem anexos PDF).")

                message.UnRead = False # Marca o email como lido
                print(f"[LOG] Email \'{message.Subject}\' processado e marcado como lido.")


if __name__ == "__main__":
    output_directory = r"C:\Temp\Outlook_PDFs"
    automator = OutlookAutomator()
    automator.process_emails(output_directory)
