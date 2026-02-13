import win32com.client as win32
import os
import zipfile

arquivo_pbix = r"caminhoarquivo"
arquivo_zip = arquivo_pbix.replace(".pbix", ".zip")

with zipfile.ZipFile(arquivo_zip, 'w', zipfile.ZIP_DEFLATED) as zipf:
    zipf.write(arquivo_pbix, os.path.basename(arquivo_pbix))

print("📦 ZIP criado:", arquivo_zip)


outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)

mail.BCC = "emmailcopiaoculta"
mail.Subject = "Indicador de Desempenho Ambiental"
mail.Body = (
    "Prezado(a),\n\n"
    "Segue em anexo o indicador de desempenho ambiental da área atualizado "
    "Att."
)

if os.path.exists(arquivo_zip):
    mail.Attachments.Add(arquivo_zip)
else:
    print("Arquivo ZIP não encontrado:", arquivo_zip)

mail.Send()

print("✅ E-mail enviado com sucesso!")
