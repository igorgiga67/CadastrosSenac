from flask import Flask, render_template, request, redirect, url_for, flash
import pandas as pd
from flask_mail import Mail, Message
import openpyxl

app = Flask(__name__)
app.secret_key = 'some_secret_key'

cadastros = {}
max_pessoas_por_horario = 20
arquivo_excel = "reservas.xlsx"

app.config['MAIL_SERVER'] = 'smtp.office365.com'  # Substitua 'your-email-provider.com' pelo servidor SMTP do seu provedor de e-mail.
app.config['MAIL_PORT'] = 587  # Porta para envio de e-mails. 587 é comum para conexões não seguras, 465 para seguras.
app.config['MAIL_USERNAME'] = 'igor-67@hotmail.com'  # Seu endereço de e-mail
app.config['MAIL_PASSWORD'] = 'urvucrzdeegrkxuy'  # Sua senha de e-mail
app.config['MAIL_USE_TLS'] = True  # Use isso para conexões não seguras
app.config['MAIL_USE_SSL'] = False  # Se estiver usando a porta 465, defina isso como True

mail = Mail(app)

@app.route('/')
def index():
    return render_template('index.html', cadastros=cadastros)

@app.route('/reservar', methods=['POST'])
def reservar():
    horario_reserva = request.form.get('horario')
    nome = request.form.get('nome')
    email = request.form.get('email')
    empresa = request.form.get('empresa')
    telefone = request.form.get('telefone')
    

    if not all([nome, email, empresa, telefone, horario_reserva]):
        flash('Todos os campos são obrigatórios!')
        return redirect(url_for('index'))

    if horario_reserva not in cadastros:
        cadastros[horario_reserva] = []

    if len(cadastros[horario_reserva]) < max_pessoas_por_horario:
        cadastros[horario_reserva].append({
            'nome': nome,
            'email': email,
            'empresa': empresa,
            'telefone': telefone
        })
        flash(f"Reserva feita para {nome} às {horario_reserva}!")
    else:
        flash(f"O horário de {horario_reserva} já está cheio. Escolha outro horário.")

    return redirect(url_for('index'))

@app.route('/salvar')
def salvar():
    data = []

    for horario, reservas in cadastros.items():
        for reserva in reservas:
            reserva_data = {
                'horario': horario,
                **reserva
            }
            data.append(reserva_data)

    df = pd.DataFrame(data)
    df.to_excel(arquivo_excel, index=False)
    flash(f"Reservas salvas em {arquivo_excel}!")
    
    msg = Message('Reservas Salvas', sender='igor-67@hotmail.com', recipients=['igorgiga67@gmail.com'])
    msg.body = 'Olá, tudo bem? Meu nome é Igor, e sou Instrutor no Senac MT. Este email é um email automático referente a RQ - 060 que foi preenchida no evento que está ocorrendo neste momento.'
    
    caminho_arquivo = r'C:\Users\igor-\OneDrive\Área de Trabalho\cadastro\reservas.xlsx'
    
    with open(caminho_arquivo, 'rb') as fp:
        msg.attach("reservas.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fp.read())

    #with app.open_resource("C:\Users\igor-\OneDrive\Área de Trabalho\cadastro") as fp:
    #    msg.attach("reservas.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fp.read())
    mail.send(msg)

    flash(f"Reservas salvas em {arquivo_excel} e e-mail enviado!")
    return redirect(url_for('index'))

@app.route('/excluir/<nome>', methods=['GET'])
def excluir_reserva(nome):
    reservas = []
    with open("reservas.xlsx", "rb") as f:
        workbook = openpyxl.load_workbook(f)
        sheet = workbook.active
        for row in sheet.iter_rows(values_only=True):
            if row[0] == "Nome":
                continue  # Pula o cabeçalho
            reserva = {
                'horario': row[0],
                'nome': row[1],
                'email': row[2],
                'empresa': row[3],
                'telefone': row[4]
            }
            if reserva['nome'] != nome:
                reservas.append(reserva)
    
    # Sobrescreva o arquivo Excel sem a reserva excluída
    novo_workbook = openpyxl.Workbook()
    novo_sheet = novo_workbook.active
    
    if not reservas:
        novo_sheet.append(["horario", "nome", "email", "empresa", "telefone"])
    for reserva in reservas:
        novo_sheet.append([reserva['horario'], reserva['nome'], reserva['email'], reserva['empresa'], reserva['telefone']])
    novo_workbook.save("reservas.xlsx")
    
    cadastros.clear()  # Limpe o dicionário cadastros

    # Recarregue as reservas do arquivo Excel atualizado
    df = pd.read_excel(arquivo_excel)
    for _, row in df.iterrows():
        horario = row['horario']
        reserva = {
            'nome': row['nome'],
            'email': row['email'],
            'empresa': row['empresa'],
            'telefone': row['telefone']
        }

        if horario not in cadastros:
            cadastros[horario] = []

        cadastros[horario].append(reserva)

    
    flash('Reserva excluída com sucesso!')
    return redirect(url_for('index'))


if __name__ == '__main__':
    try:
        df = pd.read_excel(arquivo_excel)
        for _, row in df.iterrows():
            horario = row['horario']
            reserva = {
                'nome': row['nome'],
                'email': row['email'],
                'empresa': row['empresa'],
                'telefone': row['telefone']
            }

            if horario not in cadastros:
                cadastros[horario] = []

            cadastros[horario].append(reserva)
    except FileNotFoundError:
        print("Arquivo de reservas não encontrado. Começando com uma lista vazia.")

    app.run(debug=True)
