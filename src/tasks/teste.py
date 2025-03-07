import smtplib

EMAIL_USER = "eduardo.rochavargas2@gmail.com" #os.getenv("EMAIL_USER")
EMAIL_PASS = "aplb ohnj ncpg epjb"

try:
    server = smtplib.SMTP("smtp.gmail.com", 587)
    server.starttls()
    server.login(EMAIL_USER, EMAIL_PASS)
    print("✅ Conexão bem-sucedida!")
    server.quit()
except Exception as e:
    print(f"❌ Erro: {e}")
