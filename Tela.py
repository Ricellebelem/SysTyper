from flask import Flask, render_template, request, redirect, url_for
import sqlite3
import db

app = Flask(__name__)

# Nome da empresa e do banco de dados
empresa = "Typer"
database_name = f"db_{empresa}.db"

def get_db_connection():
    conn = sqlite3.connect(database_name)
    conn.row_factory = sqlite3.Row
    return conn


# Rota principal para exibir o formulário e os resultados
@app.route('/')
def home():
    return render_template('index.html')

@app.route('/deletar_ordem', methods=['POST'])
def deletar_ordem():
    sob = request.form.get('sob')
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM processo WHERE sob = ?", (sob,))
    conn.commit()
    conn.close()
    return redirect(url_for('home'))

@app.route('/', methods=['GET', 'POST'])
def buscar_solicitante():
    solicitante = None  # Inicializa a variável do solicitante como vazia
    sob = ""  # Variável para manter o valor da SOB digitada

    if request.method == 'POST':
        # Obtendo a SOB digitada no formulário
        sob = request.form.get('sob')

        # Conectar ao banco de dados
        conn = get_db_connection()
        cursor = conn.cursor()

        # Consultar o campo "solicitante" baseado na SOB
        cursor.execute("SELECT solicitante FROM processo WHERE sob = ?", (sob,))
        result = cursor.fetchone()  # Retorna o primeiro resultado ou None

        if result:
            solicitante = result['solicitante']  # Pega o valor do solicitante

        # Fechar a conexão com o banco de dados
        conn.close()

    # Renderizar a página com o formulário e o resultado
    return render_template('index.html', solicitante=solicitante, sob=sob)
if __name__ == '__main__':
    app.run(debug=True)