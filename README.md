DESCRIÇÃO DO CÓDIGO - SISTEMA DE ENVIO DE EMAIL

Este sistema automatiza o envio de emails usando dados de uma planilha Excel, integrando-se com o Microsoft Outlook. O código é organizado em três arquivos principais:

1. email.py
-----------
Classe principal que gerencia todo o processo de envio de emails:
- Abre e lê dados do Excel
- Valida as informações
- Constrói o corpo do email em HTML
- Envia os emails via Outlook
- Gerencia recursos (Excel e Outlook)
- Registra logs de operações

2. config.py
-----------
Arquivo de configuração que contém:
- Lista de emails para cópia (CC)
- Assunto padrão dos emails
- Caminho do arquivo Excel
- Nome da planilha
- Configurações de logging
- Colunas obrigatórias para validação

3. logger.py
-----------
Sistema de logging que:
- Cria arquivo de log automaticamente
- Registra operações e erros
- Mostra mensagens no console
- Mantém histórico de execuções

FUNCIONAMENTO:
1. O sistema lê uma planilha Excel
2. Agrupa os dados conforme necessário
3. Para cada grupo, constrói um email
4. Envia o email via Outlook
5. Registra todas as operações em log

REQUISITOS:
- Python 3.x
- Microsoft Outlook instalado
- Pacotes: pandas, pywin32
- Permissões para acessar Excel e Outlook

USO:
1. Configure o arquivo config.py
2. Coloque a planilha no local correto
3. Execute o script email.py
4. Verifique os logs em caso de erro 
