import fitz  
import pandas as pd
import yagmail
import os

def adicionar_texto_com_caixa(pdf_fundo_path, nome, curso, output_pdf_path):

    try:
        pdf_document = fitz.open(pdf_fundo_path)
    except FileNotFoundError:
        print(f"Arquivo {pdf_fundo_path} não encontrado.")
        return

    pagina = pdf_document[0]
    
    # Definir o texto a ser adicionado
    texto = (f"Certificamos que {nome},\n"
             f"participou do curso de {curso} realizado pela área de Robótica Educacional do\n"
             "Instituto Conceição Moura.\n"
             "A formação foi concluída em junho de 2024 com carga horária de 25h.")

    font = "helv"
    fontsize = 13

    largura, altura = pagina.rect.width, pagina.rect.height

    # Caixa delimitadora
    largura_caixa = 400  # Largura fixa da caixa
    altura_caixa = 120   # Altura fixa da caixa

    # Calcular a posição da caixa para centralizar
    posicao_caixa_x = (largura - largura_caixa) / 2
    posicao_caixa_y = (altura - altura_caixa) / 2 - 15

    # Verificar tamanho da caixa:
    #pagina.draw_rect(fitz.Rect(posicao_caixa_x, posicao_caixa_y, posicao_caixa_x + largura_caixa, posicao_caixa_y + altura_caixa), color=(0, 0, 0), width=1)

    rect = fitz.Rect(posicao_caixa_x, posicao_caixa_y, posicao_caixa_x + largura_caixa, posicao_caixa_y + altura_caixa)

    # Calcular altura da linha
    text_line_height = fontsize * 1.5  
    text_lines = texto.split('\n')

    y = posicao_caixa_y + (altura_caixa - len(text_lines) * text_line_height) / 2

    for line in text_lines:
        text_width = fitz.get_text_length(line, fontname=font, fontsize=fontsize)
        x = posicao_caixa_x + (largura_caixa - text_width) / 2  
        # Modificar cor do texto(apenas 0 ou 1 biblioteca não trabalha com RGB):
        pagina.insert_text((x, y), line, fontsize=fontsize, fontname=font, color=(1, 1, 1))
        y += text_line_height 

    pdf_document.save(output_pdf_path)
    pdf_document.close()


def enviar_certificado_por_email(nome, email, pdf_path, titulo, conteudo, yagmail_client):
  
    try:
        yagmail_client.send(
            to=email,
            subject=titulo,
            contents=conteudo,
            attachments=pdf_path
        )
        print(f"E-mail enviado para {email}.")
    except Exception as e:
        print(f"Falha ao enviar e-mail para {email}: {e}")

def gerar_certificados(planilha_path, pdf_fundo_path, pasta_saida, yagmail_user, yagmail_password):

    os.makedirs(pasta_saida, exist_ok=True)
    dados = pd.read_excel(planilha_path)

    yagmail_client = yagmail.SMTP(yagmail_user, yagmail_password)

    for i, linha in dados.iterrows():
        nome = linha['nome']
        curso = linha['curso']
        email = linha['email']  

        nome_arquivo = nome.replace(' ', '_').replace('/', '_').replace('\\', '_')
        output_pdf = f"{pasta_saida}/certificado_{nome_arquivo}.pdf"
        
        adicionar_texto_com_caixa(pdf_fundo_path, nome, curso, output_pdf)

        # Conteúdo do e-mail
        titulo = 'Certificado de Participação - Introdução ao Arduino'
        conteudo = f'Olá {nome},\nAqui está o seu certificado de participação.\nAtenciosamente,\nEquipe ICM'

        enviar_certificado_por_email(nome, email, output_pdf, titulo, conteudo, yagmail_client)

# Principais Configurações
planilha = "certificados.xlsx"  # Planilha com os dados
pdf_fundo = "certificado_modelo.pdf"  # PDF de fundo
pasta_saida = "certificados_gerados"  # Pasta para salvar os certificados

# Configurações do Email / Gerar senha de app no gmail
yagmail_user = "email@email.com"
yagmail_password = "senha"

gerar_certificados(planilha, pdf_fundo, pasta_saida, yagmail_user, yagmail_password)
