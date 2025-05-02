import os
from pathlib import Path

# Configurações de Email
EMAIL_CC = [
    "adicione o email CC aqui"
]
ASSUNTO_EMAIL = "ASSUNTO DO EMAIL"

# Configurações de Arquivo
BASE_DIR = Path(__file__).parent
EXCEL_FILE = ""  # Modifique este caminho
WORKSHEET_NAME = "nome da aba"

# Configurações de Log
LOG_FILE = os.path.join(BASE_DIR, "logs", "email_sender.log")
LOG_LEVEL = "INFO"

# Configurações de Validação
REQUIRED_COLUMNS = [
    "Chave de Acesso",
    "N° da Nota",
    "Justificativa",
    "Justificativa2",
    "Info Adicional"
] 