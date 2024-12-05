from flask import Flask, request, render_template, send_file
from docx import Document
from docx2pdf import convert
from datetime import datetime
import os
import pythoncom

app = Flask(__name__)

def substituir_campos_docx(caminho_docx_entrada, nome, email, curso):
    try:
        doc = Document(caminho_docx_entrada)
        
        for p in doc.paragraphs:
            for run in p.runs:
                if '{{name}}' in run.text:
                    run.text = run.text.replace('{{name}}', nome)
                if '{{email}}' in run.text:
                    run.text = run.text.replace('{{email}}', email)
                if '{{curso}}' in run.text:
                    run.text = run.text.replace('{{curso}}', curso)
                if '{{date}}' in run.text:
                    run.text = run.text.replace('{{date}}', datetime.now().strftime('%d/%m/%Y'))

        # Salva o documento modificado
        doc.save('temp_modificado.docx')
        return 'temp_modificado.docx'
    except Exception as e:
        print(f"Erro ao substituir campos no documento: {e}")
        raise

def gerar_pdf_de_docx(caminho_docx_entrada):
    pythoncom.CoInitialize()  # Inicializa o COM
    caminho_pdf_saida = 'temp_certificado.pdf'
    convert(caminho_docx_entrada, caminho_pdf_saida)
    return caminho_pdf_saida

def remover_arquivo(caminho_arquivo):
    try:
        os.remove(caminho_arquivo)
    except Exception as e:
        print(f"Erro ao remover o arquivo {caminho_arquivo}: {e}")

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        nome = request.form['nome']
        email = request.form['email']
        curso = request.form['curso']  # Você pode definir isso como fixo ou dinâmico conforme necessário

        caminho_docx_entrada = 'Dados_do_Curso.docx'
        
        # Substituir os campos no DOCX
        try:
            caminho_docx_modificado = substituir_campos_docx(caminho_docx_entrada, nome, email, curso)
            # Gerar PDF temporário
            caminho_pdf_saida = gerar_pdf_de_docx(caminho_docx_modificado)

            # Enviar o PDF gerado como download
            response = send_file(caminho_pdf_saida, as_attachment=True, download_name=f'certificado_{nome}.pdf')

            # Remover os arquivos temporários após o download
            remover_arquivo(caminho_docx_modificado)
            remover_arquivo(caminho_pdf_saida)

            return response
        except Exception as e:
            print(f"Erro durante o processo de geração do certificado: {e}")
            return "Houve um erro ao gerar o certificado. Por favor, tente novamente."

    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True)