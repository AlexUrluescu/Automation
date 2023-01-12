import openpyxl 
from email.message import EmailMessage
import ssl
import smtplib

book = openpyxl.load_workbook('registru.xlsx', data_only=True)

foaie = book.active

casute = foaie['A4':'D5']

array = [] 

for fila in casute:
    array.append([celda.value for celda in fila])

print(array)

for lista in array:
    print('a inceput')
    nume = lista[0]
    prenume = lista[1]
    email = lista[3]
    nota = str(lista[2])
    print(f"Nume: {nume}")
    print(f"Prenume: {prenume}")
    print(f"Email: {email}")
    print(f"Nota: {nota}")

    email_emisor = 'alexurluescu23@gmail.com'
    email_parola = 'xecgxqiixecxaayw'
    email_receptor = email

    subiect = 'Nota fizica'
    cuerpo = 'Buna ziua ' + nume + " " + prenume + ' ai luat nota ' + nota

    em = EmailMessage()
    em['From'] = email_emisor
    em['To'] = email_receptor
    em['Subject'] = subiect
    em.set_content(cuerpo)

    contexto = ssl.create_default_context()

    with smtplib.SMTP_SSL('smtp.gmail.com', 465, context=contexto) as smtp:
        smtp.login(email_emisor, email_parola)
        smtp.sendmail(email_emisor, email_receptor, em.as_string())

    print('s-a terminat')