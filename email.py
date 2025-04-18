import win32com.client as win32
import pandas as pd
import os


# Constants (consider moving these to a config file)
EMAIL_CC = "jose.batista@bunge.com; sidnei.souza@bunge.com; flavia.leite@bunge.com"
ASSUNTO_EMAIL = "Solicitação de carta de correção"
EXCEL_FILE = "C:/Users/CT024486/OneDrive - BUNGE/Área de Trabalho/Nova pasta/teste.xlsm"  # Replace with your file path

def send_grouped_email():
    try:
        # Open the specified Excel file
        excel = win32.gencache.EnsureDispatch("Excel.Application")
        wb = excel.Workbooks.Open(EXCEL_FILE) # Open the specified workbook
        ws = wb.Worksheets("Controle cce")

        # ... (rest of your code remains the same)
       # ...

        # Close Outlook and Excel explicitly (Good Practice)
       # outlook.Quit()
        #wb.Close(SaveChanges=True)  # Save changes if needed
        excel.Quit()


    except Exception as e:
        print(f"An error occurred: {e}")
        # Ensure Excel is visible even if an error occurs
        excel.Visible = True #Keep excel visiable
        # Close Excel if it's open (Important for error handling)
        if excel: 
            wb.Close(SaveChanges=False)  # Or True to save changes even after an error
            excel.Quit()


def build_email_body(group):
      # Use HTML table for better formatting
    body = "<table>"
    for _, row in group.iterrows():
        body += (
            f"<tr><td>Chave de Acesso:</td><td>{row[0]}</td></tr>"
            f"<tr><td>N° da Nota:</td><td>{row[6]}</td></tr>"
            f"<tr><td>Justificativa:</td><td>{row[8]}</td></tr>"
            f"<tr><td>Justificativa2:</td><td>{row[9]}</td></tr>"
            f"<tr><td>Info Adicional:</td><td>{row[10]}</td></tr>" # Include label for Info Adicional
        )
    body += "</table>"
    return body




# Placeholder functions (replace with your actual implementations)
def historico():
    print("Placeholder for historico function")

def excluir():
    print("Placeholder for excluir function")


if __name__ == "__main__":
    send_grouped_email()