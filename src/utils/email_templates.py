import pandas as pd

def email_success_subject():
    return f"RPA Cadastro Cliente - {pd.Timestamp.now().strftime('%d/%m/%Y')} – {pd.Timestamp.now().strftime('%H:%M')}"

def email_success_body():
    return f"""
    O processo RPA Cadastro Cliente foi executado com sucesso na data {pd.Timestamp.now().strftime('%d/%m/%Y')} – {pd.Timestamp.now().strftime('%H:%M')}.
    """

def email_error_subject(process_name):
    return f"Erro - RPA {process_name} {pd.Timestamp.now().strftime('%d%m%Y')} – {pd.Timestamp.now().strftime('%H%M')}"

def email_error_body(process_name, task_name):
    return f"""
    Foi encontrado ERRO durante a execução do processo RPA {process_name}, na data {pd.Timestamp.now().strftime('%d/%m/%Y')} – {pd.Timestamp.now().strftime('%H:%M')}, na tarefa {task_name}.
    """
