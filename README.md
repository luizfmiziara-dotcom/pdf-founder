EXTRATOR DE PAGINAS COM DIRECOES NM

O que faz
- Permite selecionar um ou varios PDFs.
- Procura legendas como 0 grau NM, 30 graus NM e 90 graus NM.
- Gera um novo PDF por arquivo contendo somente as paginas encontradas.
- Se houver varios resultados, permite baixar todos em um ZIP.

Como instalar
1. Instale Python 3.10 ou superior.
2. Abra o Prompt de Comando nesta pasta.
3. Crie um ambiente virtual:
   python -m venv .venv
4. Ative o ambiente no Windows:
   .venv\Scripts\activate
5. Instale as dependencias:
   pip install -r requirements.txt

Como executar
streamlit run app_atc_pdf_v2.py

Gerar atalho para a aplicação: 
- novo atalho
- Destino: C:\Users\Luiz.Miziara\ProjetosPython\Streamlit\iniciar_app.py
- Iniciar em: C:\Users\Luiz.Miziara\ProjetosPython\Streamlit 

O navegador abrira a interface automaticamente.

Regra atual
0, 30, 60, 90, 120, 150, 180, 210, 240, 270, 300 e 330 graus, seguidos de NM.


