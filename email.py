import win32com.client as win32
import pandas as pd
import os
from pathlib import Path
from config import (
    EMAIL_CC,
    ASSUNTO_EMAIL,
    EXCEL_FILE,
    WORKSHEET_NAME,
    REQUIRED_COLUMNS
)
from logger import logger

class EmailSender:
    def __init__(self):
        self.excel = None
        self.wb = None
        self.ws = None
        self.outlook = None

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close_resources()

    def close_resources(self):
        """Fecha todos os recursos abertos de forma segura"""
        try:
            if self.wb:
                self.wb.Close(SaveChanges=False)
            if self.excel:
                self.excel.Quit()
            if self.outlook:
                self.outlook.Quit()
        except Exception as e:
            logger.error(f"Erro ao fechar recursos: {e}")

    def validate_excel_file(self):
        """Valida se o arquivo Excel existe e é acessível"""
        if not os.path.exists(EXCEL_FILE):
            raise FileNotFoundError(f"Arquivo Excel não encontrado: {EXCEL_FILE}")
        
        if not os.access(EXCEL_FILE, os.R_OK):
            raise PermissionError(f"Sem permissão para ler o arquivo: {EXCEL_FILE}")

    def open_excel(self):
        """Abre o arquivo Excel e a planilha especificada"""
        try:
            self.excel = win32.gencache.EnsureDispatch("Excel.Application")
            self.wb = self.excel.Workbooks.Open(EXCEL_FILE)
            self.ws = self.wb.Worksheets(WORKSHEET_NAME)
            logger.info("Arquivo Excel aberto com sucesso")
        except Exception as e:
            logger.error(f"Erro ao abrir Excel: {e}")
            raise

    def validate_data(self, df):
        """Valida se o DataFrame contém todas as colunas necessárias"""
        missing_columns = [col for col in REQUIRED_COLUMNS if col not in df.columns]
        if missing_columns:
            raise ValueError(f"Colunas faltantes no DataFrame: {missing_columns}")

    def build_email_body(self, group):
        """Constrói o corpo do email com validação de dados"""
        try:
            body = "<table>"
            for _, row in group.iterrows():
                # Validação de dados nulos
                for col in REQUIRED_COLUMNS:
                    if pd.isna(row[col]):
                        logger.warning(f"Dado nulo encontrado na coluna {col}")
                        row[col] = "N/A"

                body += (
                    f"<tr><td>Chave de Acesso:</td><td>{row['Chave de Acesso']}</td></tr>"
                    f"<tr><td>N° da Nota:</td><td>{row['N° da Nota']}</td></tr>"
                    f"<tr><td>Justificativa:</td><td>{row['Justificativa']}</td></tr>"
                    f"<tr><td>Justificativa2:</td><td>{row['Justificativa2']}</td></tr>"
                    f"<tr><td>Info Adicional:</td><td>{row['Info Adicional']}</td></tr>"
                )
            body += "</table>"
            return body
        except Exception as e:
            logger.error(f"Erro ao construir corpo do email: {e}")
            raise

    def send_email(self, to_email, subject, body, cc_emails=None):
        """Envia um email usando o Outlook"""
        try:
            if not self.outlook:
                self.outlook = win32.Dispatch('Outlook.Application')
            
            mail = self.outlook.CreateItem(0)
            mail.To = to_email
            mail.Subject = subject
            mail.HTMLBody = body
            
            if cc_emails:
                mail.CC = ';'.join(cc_emails)
            
            mail.Send()
            logger.info(f"Email enviado com sucesso para: {to_email}")
            
        except Exception as e:
            logger.error(f"Erro ao enviar email: {e}")
            raise

    def send_grouped_email(self):
        """Envia emails agrupados com tratamento de erros"""
        try:
            # Validação inicial
            self.validate_excel_file()
            
            # Abre o Excel
            self.open_excel()
            
            # Lê os dados da planilha
            df = pd.read_excel(EXCEL_FILE, sheet_name=WORKSHEET_NAME)
            
            # Valida os dados
            self.validate_data(df)
            
            # Agrupa os dados (implementar lógica de agrupamento conforme necessário)
            grouped_data = df.groupby('Chave de Acesso')
            
            # Envia emails para cada grupo
            for _, group in grouped_data:
                email_body = self.build_email_body(group)
                # Aqui você pode adicionar a lógica para determinar o destinatário
                # Por exemplo, usando uma coluna do DataFrame
                to_email = group['Email'].iloc[0] if 'Email' in group.columns else "destinatario@exemplo.com"
                self.send_email(to_email, ASSUNTO_EMAIL, email_body, EMAIL_CC)
            
            logger.info("Processo de envio de emails concluído com sucesso")
            
        except Exception as e:
            logger.error(f"Erro durante o processo de envio de emails: {e}")
            raise
        finally:
            self.close_resources()

def historico():
    """Registra o histórico de operações"""
    logger.info("Função histórico chamada")
    # Implementar lógica do histórico

def excluir():
    """Exclui registros conforme necessário"""
    logger.info("Função excluir chamada")
    # Implementar lógica de exclusão

if __name__ == "__main__":
    try:
        with EmailSender() as sender:
            sender.send_grouped_email()
    except Exception as e:
        logger.error(f"Erro fatal: {e}")
        raise