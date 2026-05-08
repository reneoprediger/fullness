import psycopg2
from psycopg2.extras import DictCursor
from flask import Flask, render_template, request, redirect, url_for, session, flash, send_from_directory, jsonify, send_file
import os
import io
import traceback
import json
import hashlib
from werkzeug.utils import secure_filename
from werkzeug.security import generate_password_hash, check_password_hash
from datetime import datetime, timedelta, date
import calendar
import pandas as pd
from ofxparse import OfxParser
from collections import defaultdict
from apscheduler.schedulers.background import BackgroundScheduler
import atexit
from itertools import groupby

# --- CONFIGURAÇÕES ---
UPLOAD_FOLDER = 'uploads/planilhas'
ALLOWED_EXTENSIONS = {'xlsx', 'xls', 'csv', 'ofx'}

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024
app.secret_key = os.environ.get('SECRET_KEY', 'uma-chave-bem-secreta-para-desenvolvimento')

if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

# --- FUNÇÕES HELPER ---

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def get_db():
    db_url = os.environ.get('DATABASE_URL')
    if not db_url:
        raise ValueError("A variável de ambiente DATABASE_URL não foi definida!")
    conn = psycopg2.connect(db_url)
    conn.cursor_factory = DictCursor
    return conn

def check_and_add_column(conn, table_name, column_name, column_type):
    """Verifica se uma coluna existe e a adiciona se necessário."""
    cursor = conn.cursor()
    try:
        cursor.execute(f"""
            SELECT column_name 
            FROM information_schema.columns 
            WHERE table_name='{table_name}' AND column_name='{column_name}';
        """)
        if not cursor.fetchone():
            cursor.execute(f"ALTER TABLE {table_name} ADD COLUMN {column_name} {column_type};")
            conn.commit()
    except Exception as e:
        conn.rollback()
        print(f"Erro ao verificar/adicionar coluna {column_name}: {e}")
    finally:
        cursor.close()

def ler_ofx_seguro(file):
    """Lê arquivo OFX tratando codificações diversas."""
    file.seek(0)
    content_bytes = file.read()

    try:
        content_str = content_bytes.decode('utf-8')
    except UnicodeDecodeError:
        try:
            content_str = content_bytes.decode('latin-1')
        except:
            content_str = content_bytes.decode('utf-8', errors='ignore')

    content_str = content_str.replace('\x00', '')
    utf8_bytes = content_str.encode('utf-8')
    return OfxParser.parse(io.BytesIO(utf8_bytes))

def get_tipo_by_codigo(codigo, user_id):
    conn = get_db()
    cursor = conn.cursor()
    prefixo = codigo.split('.')[0]
    cursor.execute("SELECT nome FROM tipos_conta WHERE user_id = %s AND prefixo = %s", (user_id, prefixo))
    tipo_conta = cursor.fetchone()
    conn.close()
    return tipo_conta['nome'] if tipo_conta else 'Outros'

def sugerir_plano_contas_automatico(descricao, user_id):
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("SELECT texto_chave, plano_conta_id FROM mapeamento_regras WHERE user_id = %s", (user_id,))
    regras = cursor.fetchall()
    conn.close()
    descricao_upper = descricao.upper()
    for regra in regras:
        if regra['texto_chave'].upper() in descricao_upper:
            return regra['plano_conta_id']
    return None

# --- FILTROS TEMPLATE ---

def format_date_br(value):
    if isinstance(value, str):
        try:
            date_obj = datetime.strptime(value, '%Y-%m-%d')
            return date_obj.strftime('%d/%m/%Y')
        except ValueError:
            return value
    elif isinstance(value, date):
        return value.strftime('%d/%m/%Y')
    return value

app.jinja_env.filters['dt_format'] = format_date_br

@app.context_processor
def inject_today_date():
    return {'today_date': date.today}

def to_date_filter(s):
    try:
        return datetime.strptime(s, '%Y-%m-%d').date()
    except (ValueError, TypeError):
        return None
app.jinja_env.filters['to_date'] = to_date_filter

# --- BANCO DE DADOS E INIT ---

def create_tables():
    conn = get_db()
    cursor = conn.cursor()
    
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS users (
            id SERIAL PRIMARY KEY, username TEXT UNIQUE, password TEXT, role TEXT, email TEXT UNIQUE,
            telefone TEXT, tipo_pessoa TEXT DEFAULT 'fisica', cpf_cnpj TEXT,
            razao_social TEXT, nome_fantasia TEXT,
            endereco_rua TEXT, endereco_numero TEXT, endereco_bairro TEXT,
            endereco_cidade TEXT, endereco_cep TEXT
        );
    """)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS contas_bancarias (
            id SERIAL PRIMARY KEY, user_id INTEGER NOT NULL, nome_banco TEXT NOT NULL,
            numero_banco TEXT, agencia TEXT, numero_conta TEXT NOT NULL, apelido_conta TEXT,
            saldo_inicial REAL DEFAULT 0.0,
            FOREIGN KEY (user_id) REFERENCES users (id) ON DELETE CASCADE
        );
    """)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS plano_contas (
            id SERIAL PRIMARY KEY, user_id INTEGER NOT NULL, codigo TEXT NOT NULL, nome TEXT NOT NULL,
            tipo TEXT NOT NULL, aceita_lancamentos BOOLEAN NOT NULL,
            FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE CASCADE, UNIQUE(user_id, codigo)
        );
    """)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS fornecedores (
            id SERIAL PRIMARY KEY, user_id INTEGER, nome TEXT NOT NULL, 
            info_adicional TEXT, telefone TEXT, tipo_pessoa TEXT, cpf_cnpj TEXT, 
            endereco_rua TEXT, endereco_numero TEXT, endereco_bairro TEXT, 
            endereco_cidade TEXT, endereco_cep TEXT, conta_padrao_id INTEGER,
            FOREIGN KEY (user_id) REFERENCES users (id) ON DELETE CASCADE, FOREIGN KEY (conta_padrao_id) REFERENCES plano_contas(id) ON DELETE SET NULL
        );
    """)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS meus_clientes (
            id SERIAL PRIMARY KEY, user_id INTEGER, nome TEXT, email TEXT, telefone TEXT, 
            tipo_pessoa TEXT, cpf_cnpj TEXT, endereco_rua TEXT, endereco_numero TEXT, 
            endereco_bairro TEXT, endereco_cidade TEXT, endereco_cep TEXT, conta_padrao_id INTEGER,
            FOREIGN KEY (user_id) REFERENCES users (id) ON DELETE CASCADE, FOREIGN KEY (conta_padrao_id) REFERENCES plano_contas(id) ON DELETE SET NULL
        );
    """)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS contas_pagar_receber (
            id SERIAL PRIMARY KEY, user_id INTEGER, fornecedor_id INTEGER, meu_cliente_id INTEGER, 
            descricao TEXT, valor_previsto REAL, valor_real REAL, data_vencimento DATE, 
            data_pagamento DATE, status TEXT, tipo TEXT, recorrencia TEXT, plano_conta_id INTEGER,
            FOREIGN KEY (user_id) REFERENCES users (id) ON DELETE CASCADE, 
            FOREIGN KEY (fornecedor_id) REFERENCES fornecedores (id) ON DELETE SET NULL, 
            FOREIGN KEY (meu_cliente_id) REFERENCES meus_clientes(id) ON DELETE SET NULL, 
            FOREIGN KEY (plano_conta_id) REFERENCES plano_contas(id) ON DELETE SET NULL
        );
    """)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS planilhas (
            id SERIAL PRIMARY KEY, admin_id INTEGER, cliente_id INTEGER, nome_arquivo_original TEXT, 
            nome_arquivo_servidor TEXT, data_upload TIMESTAMP DEFAULT CURRENT_TIMESTAMP, 
            descricao TEXT, processada BOOLEAN DEFAULT FALSE, num_transacoes_importadas INTEGER DEFAULT 0, 
            prazo_entrega DATE, 
            FOREIGN KEY (admin_id) REFERENCES users (id) ON DELETE SET NULL, 
            FOREIGN KEY (cliente_id) REFERENCES users (id) ON DELETE CASCADE
        );
    """)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS transacoes_pendentes (
            id SERIAL PRIMARY KEY, user_id INTEGER NOT NULL, conta_bancaria_id INTEGER NOT NULL,
            data_transacao DATE NOT NULL, descricao TEXT NOT NULL, valor REAL NOT NULL,
            tipo TEXT NOT NULL, fitid TEXT, data_importacao TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE CASCADE,
            FOREIGN KEY (conta_bancaria_id) REFERENCES contas_bancarias(id) ON DELETE CASCADE,
            UNIQUE(user_id, fitid)
        );
    """)    
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS transacoes_financeiras (
            id SERIAL PRIMARY KEY, planilha_id INTEGER, cliente_id INTEGER, conta_id INTEGER,
            conta_bancaria_id INTEGER, plano_conta_id INTEGER,
            data_transacao DATE, descricao TEXT, tipo TEXT, valor REAL, categoria TEXT, fitid TEXT,
            FOREIGN KEY (planilha_id) REFERENCES planilhas (id) ON DELETE CASCADE, 
            FOREIGN KEY (cliente_id) REFERENCES users (id) ON DELETE CASCADE,
            FOREIGN KEY (conta_id) REFERENCES contas_pagar_receber(id) ON DELETE SET NULL, 
            FOREIGN KEY (conta_bancaria_id) REFERENCES contas_bancarias(id) ON DELETE SET NULL,
            FOREIGN KEY (plano_conta_id) REFERENCES plano_contas(id) ON DELETE SET NULL,
            UNIQUE(cliente_id, fitid)
        );
    """)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS assinaturas (
            id SERIAL PRIMARY KEY, cliente_id INTEGER UNIQUE, data_inicio DATE, data_fim DATE, 
            status TEXT, plano_nome TEXT, 
            FOREIGN KEY (cliente_id) REFERENCES users (id) ON DELETE CASCADE
        );
    """)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS categoria_alteracoes_log (
            id SERIAL PRIMARY KEY, transacao_id INTEGER, user_id INTEGER, categoria_antiga TEXT, 
            categoria_nova TEXT, data_alteracao TIMESTAMP DEFAULT CURRENT_TIMESTAMP, 
            FOREIGN KEY (transacao_id) REFERENCES transacoes_financeiras (id) ON DELETE CASCADE, 
            FOREIGN KEY (user_id) REFERENCES users (id) ON DELETE CASCADE
        );
    """)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS mapeamento_regras (
            id SERIAL PRIMARY KEY, user_id INTEGER NOT NULL, texto_chave TEXT NOT NULL, plano_conta_id INTEGER NOT NULL,
            FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE CASCADE, 
            FOREIGN KEY (plano_conta_id) REFERENCES plano_contas(id) ON DELETE CASCADE, 
            UNIQUE(user_id, texto_chave)
        );
    """)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS notificacoes (
            id SERIAL PRIMARY KEY, user_id INTEGER NOT NULL, mensagem TEXT NOT NULL, url TEXT, lida BOOLEAN DEFAULT FALSE,
            data_criacao TIMESTAMP DEFAULT CURRENT_TIMESTAMP, 
            FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE CASCADE
        );
    """)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS tipos_conta (
            id SERIAL PRIMARY KEY, user_id INTEGER NOT NULL, prefixo TEXT NOT NULL, nome TEXT NOT NULL, natureza TEXT NOT NULL,
            FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE CASCADE, UNIQUE(user_id, prefixo)
        );
    """)
    
    conn.commit()
    
    # Usuários Padrão
    try:
        hash_admin = generate_password_hash('admin123')
        hash_client = generate_password_hash('senha123')
        cursor.execute("INSERT INTO users (username, password, role, email) VALUES ('admin_user', %s, 'admin', 'admin@exemplo.com') ON CONFLICT (username) DO NOTHING", (hash_admin,))
        cursor.execute("INSERT INTO users (username, password, role, email) VALUES ('utilizador_teste', %s, 'cliente', 'cliente@exemplo.com') ON CONFLICT (username) DO NOTHING", (hash_client,))
        conn.commit()
    except psycopg2.IntegrityError:
        conn.rollback()
    
    conn.close()
    print("Tabelas verificadas/criadas.")

def criar_plano_contas_padrao(cursor, user_id):
    """Cria a estrutura inicial de plano de contas."""
    cursor.execute("DELETE FROM plano_contas WHERE user_id = %s", (user_id,))
    cursor.execute("DELETE FROM tipos_conta WHERE user_id = %s", (user_id,))

    tipos_padrao = [
        (user_id, '01', 'Receita Bruta', 'credora'),
        (user_id, '02', 'Custo Variável', 'devedora'),
        (user_id, '03', 'Despesa Fixa', 'devedora'),
        (user_id, '04', 'Investimento', 'devedora'),
        (user_id, '05', 'Receita Não Operacional', 'credora'),
        (user_id, '06', 'Despesa Não Operacional', 'devedora')
    ]
    cursor.executemany(
        "INSERT INTO tipos_conta (user_id, prefixo, nome, natureza) VALUES (%s, %s, %s, %s) ON CONFLICT DO NOTHING",
        tipos_padrao
    )

    plano_padrao = [
        (user_id, '01', 'Receita Bruta', 'Receita Bruta', False),
        (user_id, '02', 'Custo Variável', 'Custo Variável', False),
        (user_id, '03', 'Despesa Fixa', 'Despesa Fixa', False)
    ]
    cursor.executemany(
        "INSERT INTO plano_contas (user_id, codigo, nome, tipo, aceita_lancamentos) VALUES (%s, %s, %s, %s, %s) ON CONFLICT DO NOTHING",
        plano_padrao
    )

def calcular_dre_anual_refatorado(user_id, ano):
    """Calcula o DRE de forma hierárquica."""
    conn = get_db()
    cursor = conn.cursor()

    data_inicio, data_fim = f"{ano}-01-01", f"{ano}-12-31"

    sql_transacoes = """
        SELECT
            TO_CHAR(t.data_transacao, 'MM') as mes,
            t.plano_conta_id,
            CASE WHEN t.tipo = 'entrada' THEN t.valor ELSE -t.valor END as valor_calculado
        FROM transacoes_financeiras t
        JOIN plano_contas pc ON t.plano_conta_id = pc.id
        WHERE t.cliente_id = %s AND t.data_transacao BETWEEN %s AND %s AND pc.aceita_lancamentos = TRUE AND t.plano_conta_id IS NOT NULL
    """
    cursor.execute(sql_transacoes, (user_id, data_inicio, data_fim))
    transacoes = cursor.fetchall()

    totals_analiticos = defaultdict(lambda: [0.0] * 12)
    for t in transacoes:
        mes_idx = int(t['mes']) - 1
        totals_analiticos[t['plano_conta_id']][mes_idx] += t['valor_calculado']

    cursor.execute("SELECT * FROM plano_contas WHERE user_id = %s ORDER BY codigo", (user_id,))
    contas = [dict(c) for c in cursor.fetchall()]
    conn.close()

    contas_map = {c['codigo']: c for c in contas}
    for c in contas:
        c['children'] = []
        c['monthly_values'] = [0.0] * 12
        c['total'] = 0.0
        c['avg'] = 0.0

    root_nodes = []
    for conta in sorted(contas, key=lambda x: x['codigo']):
        parent_code = '.'.join(conta['codigo'].split('.')[:-1])
        if parent_code and parent_code in contas_map:
            contas_map[parent_code]['children'].append(conta)
        else:
            root_nodes.append(conta)

    def sum_tree_totals(node):
        if len(node['children']) > 0:
            child_totals = [[0.0] * 12]
            for child in node['children']:
                child_totals.append(sum_tree_totals(child))
            
            for i in range(12):
                node['monthly_values'][i] = sum(month_values[i] for month_values in child_totals)
        else:
            node['monthly_values'] = totals_analiticos.get(node['id'], [0.0] * 12)
        
        node['total'] = sum(node['monthly_values'])
        num_months = sum(1 for v in node['monthly_values'] if v != 0)
        node['avg'] = node['total'] / num_months if num_months > 0 else 0.0
        return node['monthly_values']

    for root in root_nodes:
        sum_tree_totals(root)

    receitas = next((r['monthly_values'] for r in root_nodes if r['codigo'].startswith('01')), [0.0] * 12)
    custos = next((c['monthly_values'] for c in root_nodes if c['codigo'].startswith('02')), [0.0] * 12)
    despesas_fixas = next((d['monthly_values'] for d in root_nodes if d['codigo'].startswith('03')), [0.0] * 12)
    investimentos = next((i['monthly_values'] for i in root_nodes if i['codigo'].startswith('04')), [0.0] * 12)
    receitas_nao_op = next((rno['monthly_values'] for rno in root_nodes if rno['codigo'].startswith('05')), [0.0] * 12)
    despesas_nao_op = next((dno['monthly_values'] for dno in root_nodes if dno['codigo'].startswith('06')), [0.0] * 12)

    margem_contribuicao = [(r + c) for r, c in zip(receitas, custos)]
    
    resultado_liquido = []
    for i in range(12):
        mes_total = (
            margem_contribuicao[i] +
            despesas_fixas[i] +
            investimentos[i] +
            receitas_nao_op[i] +
            despesas_nao_op[i]
        )
        resultado_liquido.append(mes_total)
    
    return {
        "dre_tree": root_nodes,
        "margem_contribuicao": margem_contribuicao,
        "resultado_operacional": resultado_liquido
    }

# --- TAREFAS AGENDADAS ---

def tarefa_criar_notificacoes_ofx():
    with app.app_context():
        hoje = date.today()
        if hoje.day == 28:
            conn = get_db()
            cursor = conn.cursor()
            cursor.execute("SELECT id FROM users WHERE role = 'cliente'")
            clientes = cursor.fetchall()
            mensagem = "Lembrete: Não se esqueça de importar seu extrato OFX para manter suas finanças em dia!"
            url_destino = url_for('cliente_importar_ofx')
            for cliente in clientes:
                cursor.execute("INSERT INTO notificacoes (user_id, mensagem, url) VALUES (%s, %s, %s)", (cliente['id'], mensagem, url_destino))
            conn.commit()
            conn.close()

scheduler = BackgroundScheduler(daemon=True)
scheduler.add_job(func=tarefa_criar_notificacoes_ofx, trigger="interval", days=1)
scheduler.start()
atexit.register(lambda: scheduler.shutdown())

# --- ROTAS PRINCIPAIS ---

@app.route('/')
def index():
    if not session.get('logged_in'):
        return redirect(url_for('login_page'))
    return redirect(url_for('admin_dashboard_home')) if session.get('role') == 'admin' else redirect(url_for('cliente_dashboard'))

@app.route('/login', methods=['GET', 'POST'])
def login_page():
    if session.get('logged_in'): 
        return redirect(url_for('index'))
        
    if request.method == 'POST':
        username = request.form['username']
        password = request.form['password']
        
        conn = get_db()
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM users WHERE username = %s", (username,))
        user = cursor.fetchone()
        conn.close()

        if user and check_password_hash(user['password'], password):
            session.update(
                logged_in=True, 
                user_id=user['id'], 
                username=user['username'], 
                role=user['role']
            )
            flash(f'Login como {user["role"]} bem-sucedido!', 'success')
            return redirect(url_for('index'))
        else: 
            flash('Nome de usuário ou senha incorretos.', 'danger')
            
    return render_template('login.html')

@app.route('/logout')
def logout():
    session.clear()
    flash('Sessão terminada com sucesso.', 'info')
    return redirect(url_for('login_page'))

# --- ROTAS DE CLIENTE ---

@app.route('/cliente/dashboard')
def cliente_dashboard():
    if not session.get('logged_in') or session.get('role') != 'cliente':
        return redirect(url_for('login_page'))
    return render_template('cliente_dashboard.html', active_page='dashboard')

@app.route('/cliente/setup-inicial')
def cliente_setup_inicial():
    if not session.get('logged_in') or session.get('role') != 'cliente':
        flash('Você precisa estar logado como cliente.', 'warning')
        return redirect(url_for('login_page'))
    
    user_id = session['user_id']
    conn = get_db()
    cursor = conn.cursor()
    criar_plano_contas_padrao(cursor, user_id)
    conn.commit()
    conn.close()
    
    flash('Plano de Contas Padrão foi criado/restaurado com sucesso!', 'success')
    return redirect(url_for('cliente_plano_de_contas'))

@app.route('/cliente/painel_indicadores')
def cliente_painel_indicadores():
    if not session.get('logged_in') or session.get('role') != 'cliente':
        return redirect(url_for('login_page'))

    user_id = session['user_id']
    hoje = date.today()
    ano_selecionado = int(request.args.get('ano', hoje.year))

    dre_data = calcular_dre_anual_refatorado(user_id, ano_selecionado)
    
    root_nodes = dre_data.get('dre_tree', [])
    receita_node = next((r for r in root_nodes if r['codigo'].startswith('01')), None)
    despesas_node = next((d for d in root_nodes if d['codigo'].startswith('03')), None)

    receita_bruta_total = receita_node['total'] if receita_node else 0.0
    despesas_fixas_total = abs(despesas_node['total']) if despesas_node else 0.0
    
    resultado_operacional_total = sum(dre_data.get('resultado_operacional', []))
    margem_contribuicao_total = sum(dre_data.get('margem_contribuicao', []))

    lucratividade = (resultado_operacional_total / receita_bruta_total * 100) if receita_bruta_total else 0
    margem_contribuicao_percentual = (margem_contribuicao_total / receita_bruta_total * 100) if receita_bruta_total else 0
    ponto_de_equilibrio = (despesas_fixas_total / (margem_contribuicao_percentual / 100)) if margem_contribuicao_percentual > 0 else 0
    
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("SELECT count(id) as num_vendas FROM transacoes_financeiras WHERE cliente_id = %s AND TO_CHAR(data_transacao, 'YYYY') = %s AND tipo = 'entrada' AND plano_conta_id IS NOT NULL", (user_id, str(ano_selecionado)))
    num_vendas_row = cursor.fetchone()
    num_vendas = num_vendas_row['num_vendas'] if num_vendas_row else 0
    conn.close()
    
    ticket_medio = receita_bruta_total / num_vendas if num_vendas > 0 else 0
    
    indicadores = {
        'lucratividade': lucratividade, 
        'ponto_de_equilibrio': ponto_de_equilibrio, 
        'margem_contribuicao': margem_contribuicao_percentual, 
        'ticket_medio': ticket_medio, 
        'receita_total': receita_bruta_total, 
        'resultado_total': resultado_operacional_total
    }
    
    anos_disponiveis = range(datetime.now().year + 1, datetime.now().year - 5, -1)
    
    return render_template(
        'cliente_painel_indicadores.html', 
        active_page='painel_indicadores', 
        ano_selecionado=ano_selecionado, 
        anos_disponiveis=anos_disponiveis, 
        indicadores=indicadores
    )

@app.route('/cliente/importar-ofx', methods=['GET', 'POST'])
def cliente_importar_ofx():
    if not session.get('logged_in') or session.get('role') != 'cliente':
        return redirect(url_for('login_page'))
    user_id = session['user_id']
    conn = get_db()
    cursor = conn.cursor()
    
    if request.method == 'POST':
        file = request.files.get('ofx_file')
        conta_bancaria_id = request.form.get('conta_bancaria_id')
        if not file or file.filename == '' or not conta_bancaria_id:
            flash('Por favor, selecione uma conta bancária e um ficheiro OFX.', 'warning')
            conn.close()
            return redirect(request.url)

        if allowed_file(file.filename):
            try:
                cursor.execute("SELECT fitid FROM transacoes_financeiras WHERE cliente_id = %s AND fitid IS NOT NULL", (user_id,))
                fitids_existentes = {row['fitid'] for row in cursor.fetchall()}
                cursor.execute("SELECT fitid FROM transacoes_pendentes WHERE user_id = %s AND fitid IS NOT NULL", (user_id,))
                fitids_existentes.update({row['fitid'] for row in cursor.fetchall()})

                ofx = ler_ofx_seguro(file)

                transacoes_temporarias = []
                all_accounts = ofx.accounts + getattr(ofx, 'credit_cards', [])
                
                for account in all_accounts:
                    for t in account.statement.transactions:
                        if t.id not in fitids_existentes:
                            valor = float(t.amount)
                            transacoes_temporarias.append({
                                'data': t.date.strftime('%Y-%m-%d'), 
                                'descricao': t.memo, 
                                'valor': abs(valor), 
                                'tipo': 'entrada' if valor >= 0 else 'saida', 
                                'fitid': t.id
                            })

                if not transacoes_temporarias:
                    flash('Nenhuma transação nova encontrada para importar (todas já existem no sistema).', 'info')
                    conn.close()
                    return redirect(request.url)

                session['transacoes_para_mapear_cliente'] = transacoes_temporarias
                session['conta_bancaria_id_importacao'] = conta_bancaria_id
                conn.close()
                return redirect(url_for('cliente_mapear_transacoes'))
                
            except Exception as e:
                flash(f'Erro ao processar o ficheiro OFX: {e}', 'danger')
        else:
            flash('Tipo de ficheiro não permitido. Apenas ficheiros .ofx são aceites.', 'danger')
        conn.close()
        return redirect(request.url)

    cursor.execute("SELECT * FROM contas_bancarias WHERE user_id = %s ORDER BY apelido_conta", (user_id,))
    contas_bancarias = cursor.fetchall()
    conn.close()
    return render_template('cliente_importar_ofx.html', active_page='importar_ofx', contas_bancarias=contas_bancarias)

@app.route('/cliente/mapear-transacoes')
def cliente_mapear_transacoes():
    if not session.get('logged_in') or session.get('role') != 'cliente':
        return redirect(url_for('login_page'))
    user_id = session['user_id']
    transacoes = session.get('transacoes_para_mapear_cliente', [])
    if not transacoes:
        return redirect(url_for('cliente_importar_ofx'))

    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("SELECT id, codigo, nome FROM plano_contas WHERE user_id = %s AND aceita_lancamentos = TRUE ORDER BY codigo", (user_id,))
    plano_contas = cursor.fetchall()
    conn.close()

    for t in transacoes:
        t['plano_sugerido_id'] = sugerir_plano_contas_automatico(t['descricao'], user_id)

    return render_template('cliente_mapear_transacoes.html', active_page='importar_ofx', transacoes=transacoes, plano_contas=plano_contas)

@app.route('/cliente/salvar_transacoes_importadas', methods=['POST'])
def cliente_salvar_transacoes_importadas():
    if not session.get('logged_in') or session.get('role') != 'cliente':
        return redirect(url_for('login_page'))
    user_id = session['user_id']
    conta_bancaria_id = session.get('conta_bancaria_id_importacao')
    
    datas = request.form.getlist('data')
    descricoes = request.form.getlist('descricao')
    valores = request.form.getlist('valor')
    tipos = request.form.getlist('tipo')
    fitids = request.form.getlist('fitid')
    planos_conta_ids = request.form.getlist('plano_conta_id')
    
    transacoes_para_inserir = []
    transacoes_para_pendencia = []
    
    for i in range(len(datas)):
        if planos_conta_ids[i]:
            transacoes_para_inserir.append((
                user_id, conta_bancaria_id, planos_conta_ids[i], 
                datas[i], descricoes[i], tipos[i], 
                float(valores[i]), fitids[i]
            ))
        else:
            transacoes_para_pendencia.append((
                user_id, conta_bancaria_id, datas[i], 
                descricoes[i], float(valores[i]), tipos[i], fitids[i]
            ))
            
    conn = get_db()
    cursor = conn.cursor()
    if transacoes_para_inserir:
        try:
            cursor.executemany("""
                INSERT INTO transacoes_financeiras 
                (cliente_id, conta_bancaria_id, plano_conta_id, data_transacao, descricao, tipo, valor, fitid) 
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s) ON CONFLICT (cliente_id, fitid) DO NOTHING
            """, transacoes_para_inserir)
        except psycopg2.Error as e:
            conn.rollback()
            flash(f"Erro ao salvar transações: {e}", "danger")
            conn.close()
            return redirect(url_for('cliente_importar_ofx'))
            
    if transacoes_para_pendencia:
        try:
            cursor.executemany("""
                INSERT INTO transacoes_pendentes 
                (user_id, conta_bancaria_id, data_transacao, descricao, valor, tipo, fitid) 
                VALUES (%s, %s, %s, %s, %s, %s, %s) ON CONFLICT (user_id, fitid) DO NOTHING
            """, transacoes_para_pendencia)
        except psycopg2.Error as e:
            conn.rollback()
            flash(f"Erro ao salvar pendências: {e}", "danger")
            conn.close()
            return redirect(url_for('cliente_importar_ofx'))
            
    conn.commit()
    conn.close()
    flash(f'{len(transacoes_para_inserir)} transações importadas e {len(transacoes_para_pendencia)} salvas como pendentes.', 'success')
    session.pop('transacoes_para_mapear_cliente', None)
    session.pop('conta_bancaria_id_importacao', None)
    return redirect(url_for('cliente_transacoes_pendentes'))

@app.route('/cliente/transacoes_pendentes', methods=['GET', 'POST'])
def cliente_transacoes_pendentes():
    if not session.get('logged_in') or session.get('role') != 'cliente':
        return redirect(url_for('login_page'))
    user_id = session['user_id']
    conn = get_db()
    cursor = conn.cursor()
    
    if request.method == 'POST':
        pendencia_id = request.form.get('pendencia_id')
        plano_conta_id = request.form.get('plano_conta_id')
        if pendencia_id and plano_conta_id:
            cursor.execute("SELECT * FROM transacoes_pendentes WHERE id = %s AND user_id = %s", (pendencia_id, user_id))
            pendencia = cursor.fetchone()
            if pendencia:
                cursor.execute("""
                    INSERT INTO transacoes_financeiras (cliente_id, conta_bancaria_id, plano_conta_id, data_transacao, descricao, tipo, valor, fitid)
                    VALUES (%s, %s, %s, %s, %s, %s, %s, %s) ON CONFLICT (cliente_id, fitid) DO NOTHING
                """, (user_id, pendencia['conta_bancaria_id'], plano_conta_id, pendencia['data_transacao'], pendencia['descricao'], pendencia['tipo'], pendencia['valor'], pendencia['fitid']))
                cursor.execute("DELETE FROM transacoes_pendentes WHERE id = %s", (pendencia_id,))
                conn.commit()
                flash('Transação categorizada com sucesso!', 'success')
            else:
                flash('Pendência não encontrada.', 'danger')
        else:
            flash('Você precisa selecionar um plano de contas.', 'warning')
        conn.close()
        return redirect(url_for('cliente_transacoes_pendentes'))

    cursor.execute("""
        SELECT tp.*, cb.apelido_conta 
        FROM transacoes_pendentes tp 
        LEFT JOIN contas_bancarias cb ON tp.conta_bancaria_id = cb.id 
        WHERE tp.user_id = %s 
        ORDER BY tp.data_transacao DESC
    """, (user_id,))
    pendencias = cursor.fetchall()
    
    cursor.execute("SELECT id, codigo, nome FROM plano_contas WHERE user_id = %s AND aceita_lancamentos = TRUE ORDER BY codigo", (user_id,))
    plano_contas = cursor.fetchall()
    conn.close()
    return render_template('cliente_transacoes_pendentes.html', active_page='pendentes', pendencias=pendencias, plano_contas=plano_contas)

@app.route('/cliente/demonstrativo')
def cliente_demonstrativo():
    if not session.get('logged_in') or session.get('role') != 'cliente':
        return redirect(url_for('login_page'))
    
    user_id = session['user_id']
    ano_selecionado = request.args.get('ano', default=date.today().year, type=int)
    anos_disponiveis = range(datetime.now().year + 1, datetime.now().year - 5, -1)
    meses_pt = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
    
    dre_data = calcular_dre_anual_refatorado(user_id=user_id, ano=ano_selecionado)

    return render_template(
        'cliente_demonstrativo.html',
        active_page='demonstrativo',
        ano_selecionado=ano_selecionado,
        anos_disponiveis=anos_disponiveis,
        meses=meses_pt,
        **dre_data
    )

@app.route('/cliente/plano_de_contas')
def cliente_plano_de_contas():
    if not session.get('logged_in') or session.get('role') != 'cliente':
        return redirect(url_for('login_page'))
    
    user_id = session['user_id']
    conn = get_db()
    cursor = conn.cursor()

    cursor.execute("SELECT * FROM plano_contas WHERE user_id = %s ORDER BY codigo", (user_id,))
    contas = [dict(c) for c in cursor.fetchall()]
    conn.close()
    
    contas_dict = {c['codigo']: c for c in contas}
    root_contas = []
    
    for conta in sorted(contas, key=lambda x: x['codigo']):
        conta['children'] = []
        parent_code = '.'.join(conta['codigo'].split('.')[:-1])
        
        if parent_code and parent_code in contas_dict:
            contas_dict[parent_code]['children'].append(conta)
        else:
            root_contas.append(conta)

    return render_template(
        'cliente_plano_de_contas.html', 
        active_page='plano_de_contas',
        plano_contas_tree=root_contas
    )

@app.route('/cliente/perfil')
def cliente_perfil():
    if not session.get('logged_in') or session.get('role') != 'cliente': return redirect(url_for('login_page'))
    conn = get_db(); cursor = conn.cursor(); cursor.execute("SELECT * FROM users WHERE id = %s", (session['user_id'],)); user = dict(cursor.fetchone())
    cursor.execute("SELECT * FROM assinaturas WHERE cliente_id = %s", (session['user_id'],)); assinatura_raw = cursor.fetchone(); conn.close()
    assinatura = None
    if assinatura_raw:
        assinatura = dict(assinatura_raw)
        for key in ['data_inicio', 'data_fim']:
            if assinatura.get(key) and isinstance(assinatura[key], str):
                try: assinatura[key] = datetime.strptime(assinatura[key], '%Y-%m-%d').date()
                except ValueError: assinatura[key] = None
    return render_template('cliente_perfil.html', active_page='perfil', user=user, assinatura=assinatura)

@app.route('/cliente/planilhas') 
def cliente_planilhas():
    if not session.get('logged_in') or session.get('role') != 'cliente': return redirect(url_for('login_page'))
    conn = get_db(); cursor = conn.cursor()
    cursor.execute("SELECT p.*, a.username as admin_nome FROM planilhas p LEFT JOIN users a ON p.admin_id = a.id WHERE p.cliente_id = %s ORDER BY p.data_upload DESC", (session['user_id'],)); planilhas = [dict(row) for row in cursor.fetchall()]; conn.close()
    return render_template('cliente_planilhas.html', active_page='planilhas', planilhas=planilhas)

@app.route('/cliente/fluxo_caixa') 
def cliente_fluxo_caixa():
    if not session.get('logged_in') or session.get('role') != 'cliente': return redirect(url_for('login_page'))
    user_id = session['user_id']; conn = get_db(); cursor = conn.cursor()
    cursor.execute("SELECT DISTINCT categoria FROM transacoes_financeiras WHERE cliente_id = %s AND categoria IS NOT NULL ORDER BY categoria", (user_id,)); categorias = [row['categoria'] for row in cursor.fetchall()]
    cursor.execute("SELECT id, apelido_conta FROM contas_bancarias WHERE user_id = %s ORDER BY apelido_conta", (user_id,)); contas_bancarias = [dict(row) for row in cursor.fetchall()]; conn.close()
    return render_template('cliente_fluxo_caixa.html', active_page='fluxo_caixa', categorias=categorias, contas_bancarias=contas_bancarias)

@app.route('/cliente/contas', methods=['GET', 'POST'])
def cliente_contas():
    if not session.get('logged_in') or session.get('role') != 'cliente': return redirect(url_for('login_page'))
    user_id = session['user_id']; conn = get_db(); cursor = conn.cursor()
    if request.method == 'POST':
        action = request.form.get('action')
        if action == 'add_conta':
            try:
                cursor.execute("INSERT INTO contas_pagar_receber (user_id, fornecedor_id, meu_cliente_id, descricao, valor_previsto, data_vencimento, status, tipo, recorrencia, plano_conta_id) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)", (user_id, request.form.get('fornecedor_id') or None, request.form.get('meu_cliente_id') or None, request.form['descricao'], float(request.form['valor_previsto']), request.form['data_vencimento'], 'Pendente', request.form['tipo'], request.form.get('recorrencia'), request.form.get('plano_conta_id')))
                conn.commit(); flash('Conta adicionada com sucesso!', 'success')
            except Exception as e: conn.rollback(); flash(f'Erro ao adicionar conta: {e}', 'danger')
        elif action == 'dar_baixa':
            try:
                conta_id, valor_real, data_pagamento = request.form['conta_id'], abs(float(request.form['valor_real'])), request.form['data_pagamento']
                cursor.execute("UPDATE contas_pagar_receber SET status = 'Pago/Recebido', valor_real = %s, data_pagamento = %s WHERE id = %s AND user_id = %s", (valor_real, data_pagamento, conta_id, user_id))
                cursor.execute("SELECT * FROM contas_pagar_receber WHERE id = %s", (conta_id,)); conta = cursor.fetchone()
                if conta:
                    cursor.execute("INSERT INTO transacoes_financeiras (cliente_id, conta_id, data_transacao, descricao, tipo, valor, plano_conta_id) VALUES (%s, %s, %s, %s, %s, %s, %s)", (user_id, conta_id, data_pagamento, conta['descricao'], 'saida' if conta['tipo'] == 'pagar' else 'entrada', valor_real, conta['plano_conta_id']))
                    conn.commit(); flash('Baixa da conta realizada com sucesso!', 'success')
                else: conn.rollback(); flash('Conta não encontrada para dar baixa.', 'danger')
            except Exception as e: conn.rollback(); flash(f'Erro ao dar baixa na conta: {e}', 'danger')
        conn.close()
        return redirect(url_for('cliente_contas'))
    cursor.execute("SELECT cpr.*, f.nome as fornecedor_nome, mc.nome as meu_cliente_nome, pc.nome as plano_conta_nome FROM contas_pagar_receber cpr LEFT JOIN fornecedores f ON cpr.fornecedor_id = f.id LEFT JOIN meus_clientes mc ON cpr.meu_cliente_id = mc.id LEFT JOIN plano_contas pc ON cpr.plano_conta_id = pc.id WHERE cpr.user_id = %s ORDER BY cpr.data_vencimento", (user_id,)); contas = [dict(row) for row in cursor.fetchall()]
    cursor.execute("SELECT * FROM fornecedores WHERE user_id = %s ORDER BY nome", (user_id,)); fornecedores = [dict(row) for row in cursor.fetchall()]
    cursor.execute("SELECT * FROM meus_clientes WHERE user_id = %s ORDER BY nome", (user_id,)); meus_clientes = [dict(row) for row in cursor.fetchall()]
    cursor.execute("SELECT id, codigo, nome FROM plano_contas WHERE user_id = %s AND aceita_lancamentos = TRUE ORDER BY codigo", (user_id,)); contas_disponiveis = [dict(row) for row in cursor.fetchall()]; conn.close()
    return render_template('cliente_contas.html', active_page='contas', contas=contas, fornecedores=fornecedores, meus_clientes=meus_clientes, contas_disponiveis=contas_disponiveis)

@app.route('/cliente/fornecedores', methods=['GET', 'POST'])
def cliente_fornecedores():
    if not session.get('logged_in') or session.get('role') != 'cliente': return redirect(url_for('login_page'))
    user_id = session['user_id']; conn = get_db(); cursor = conn.cursor()
    if request.method == 'POST':
        action, conta_padrao_id = request.form.get('action'), request.form.get('conta_padrao_id') or None
        if action == 'add':
            nome = request.form.get('nome')
            if nome:
                cursor.execute("INSERT INTO fornecedores (user_id, nome, info_adicional, telefone, tipo_pessoa, cpf_cnpj, endereco_rua, endereco_numero, endereco_bairro, endereco_cidade, endereco_cep, conta_padrao_id) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)", (user_id, nome, request.form.get('info_adicional'), request.form.get('telefone'), request.form.get('tipo_pessoa', 'fisica'), request.form.get('cpf_cnpj'), request.form.get('endereco_rua'), request.form.get('endereco_numero'), request.form.get('endereco_bairro'), request.form.get('endereco_cidade'), request.form.get('endereco_cep'), conta_padrao_id))
                conn.commit(); flash('Fornecedor adicionado com sucesso!', 'success')
            else: flash('O nome do fornecedor é obrigatório.', 'warning')
        elif action == 'delete':
            fornecedor_id = request.form.get('fornecedor_id'); cursor.execute("DELETE FROM fornecedores WHERE id = %s AND user_id = %s", (fornecedor_id, user_id)); conn.commit(); flash('Fornecedor apagado com sucesso!', 'success')
        conn.close()
        return redirect(url_for('cliente_fornecedores'))
    cursor.execute("SELECT f.*, pc.codigo, pc.nome as nome_conta FROM fornecedores f LEFT JOIN plano_contas pc ON f.conta_padrao_id = pc.id WHERE f.user_id = %s ORDER BY f.nome", (user_id,)); fornecedores = [dict(row) for row in cursor.fetchall()]
    cursor.execute("SELECT id, codigo, nome FROM plano_contas WHERE user_id = %s AND aceita_lancamentos = TRUE ORDER BY codigo", (user_id,)); contas_disponiveis = [dict(row) for row in cursor.fetchall()]; conn.close()
    return render_template('cliente_fornecedores.html', active_page='fornecedores', fornecedores=fornecedores, contas_disponiveis=contas_disponiveis)

@app.route('/cliente/meus_clientes', methods=['GET', 'POST'])
def cliente_meus_clientes():
    if not session.get('logged_in') or session.get('role') != 'cliente': return redirect(url_for('login_page'))
    user_id = session['user_id']; conn = get_db(); cursor = conn.cursor()
    if request.method == 'POST':
        action, conta_padrao_id = request.form.get('action'), request.form.get('conta_padrao_id') or None
        if action == 'add':
            nome = request.form.get('nome')
            if nome:
                cursor.execute("INSERT INTO meus_clientes (user_id, nome, email, telefone, tipo_pessoa, cpf_cnpj, endereco_rua, endereco_numero, endereco_bairro, endereco_cidade, endereco_cep, conta_padrao_id) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)", (user_id, nome, request.form.get('email'), request.form.get('telefone'), request.form.get('tipo_pessoa', 'fisica'), request.form.get('cpf_cnpj'), request.form.get('endereco_rua'), request.form.get('endereco_numero'), request.form.get('endereco_bairro'), request.form.get('endereco_cidade'), request.form.get('endereco_cep'), conta_padrao_id))
                conn.commit(); flash('Cliente adicionado com sucesso!', 'success')
            else: flash('O nome do cliente é obrigatório.', 'warning')
        elif action == 'delete':
            cliente_id = request.form.get('cliente_id'); cursor.execute("DELETE FROM meus_clientes WHERE id = %s AND user_id = %s", (cliente_id, user_id)); conn.commit(); flash('Cliente apagado com sucesso!', 'success')
        conn.close()
        return redirect(url_for('cliente_meus_clientes'))
    cursor.execute("SELECT mc.*, pc.codigo, pc.nome as nome_conta FROM meus_clientes mc LEFT JOIN plano_contas pc ON mc.conta_padrao_id = pc.id WHERE mc.user_id = %s ORDER BY mc.nome", (user_id,)); meus_clientes = [dict(row) for row in cursor.fetchall()]
    cursor.execute("SELECT id, codigo, nome FROM plano_contas WHERE user_id = %s AND aceita_lancamentos = TRUE AND codigo LIKE '01%%' ORDER BY codigo", (user_id,)); contas_disponiveis = [dict(row) for row in cursor.fetchall()]; conn.close()
    return render_template('cliente_meus_clientes.html', active_page='meus_clientes', meus_clientes=meus_clientes, contas_disponiveis=contas_disponiveis)

@app.route('/cliente/contas_bancarias', methods=['GET', 'POST'])
def cliente_contas_bancarias():
    if not session.get('logged_in') or session.get('role') != 'cliente': return redirect(url_for('login_page'))
    user_id = session['user_id']; conn = get_db()
    check_and_add_column(conn, 'contas_bancarias', 'numero_banco', 'TEXT')
    cursor = conn.cursor()
    
    if request.method == 'POST':
        action = request.form.get('action')
        if action == 'add':
            apelido, banco, conta = request.form.get('apelido_conta'), request.form.get('nome_banco'), request.form.get('numero_conta')
            numero_banco = request.form.get('numero_banco')
            if all([apelido, banco, conta]):
                cursor.execute("INSERT INTO contas_bancarias (user_id, apelido_conta, nome_banco, agencia, numero_conta, numero_banco) VALUES (%s, %s, %s, %s, %s, %s)", (user_id, apelido, banco, request.form.get('agencia'), conta, numero_banco)); conn.commit(); flash('Conta bancária adicionada com sucesso!', 'success')
            else: flash('Apelido, Banco e Número da Conta são obrigatórios.', 'warning')
        elif action == 'delete':
            conta_id = request.form.get('conta_id'); cursor.execute("DELETE FROM contas_bancarias WHERE id = %s AND user_id = %s", (conta_id, user_id)); conn.commit(); flash('Conta bancária apagada com sucesso!', 'success')
        conn.close()
        return redirect(url_for('cliente_contas_bancarias'))
    
    cursor.execute("SELECT * FROM contas_bancarias WHERE user_id = %s ORDER BY apelido_conta", (user_id,)); contas = [dict(row) for row in cursor.fetchall()]; conn.close()
    return render_template('cliente_contas_bancarias.html', active_page='contas_bancarias', contas=contas)


@app.route('/cliente/configuracoes')
def cliente_configuracoes():
    if not session.get('logged_in') or session.get('role') != 'cliente': return redirect(url_for('login_page'))
    return render_template('cliente_configuracoes.html', active_page='configuracoes')

# --- ROTAS DE ADMIN ---

@app.route('/admin/dashboard')
def admin_dashboard_home():
    if not session.get('logged_in') or session.get('role') != 'admin': return redirect(url_for('login_page'))
    conn = get_db(); cursor = conn.cursor()
    cursor.execute("SELECT COUNT(id) as total FROM users WHERE role = 'cliente'"); total_clientes = cursor.fetchone()['total']
    cursor.execute("SELECT COUNT(id) as total FROM planilhas"); total_planilhas = cursor.fetchone()['total']; conn.close()
    return render_template('admin_dashboard_home.html', active_page='home', total_clientes=total_clientes, total_planilhas=total_planilhas)

@app.route('/reset_admin')
def reset_admin():
    usuario = 'admin_user' 
    nova_senha = 'admin123'
    conn = get_db()
    cursor = conn.cursor()
    hashed_pw = generate_password_hash(nova_senha)
    try:
        cursor.execute("UPDATE users SET password = %s WHERE username = %s", (hashed_pw, usuario))
        conn.commit()
        return f"Senha do usuário '{usuario}' resetada com sucesso para '{nova_senha}'!"
    except Exception as e:
        conn.rollback()
        return f"Erro ao resetar senha: {e}"
    finally:
        conn.close()

@app.route('/admin/demonstrativo/exportar')
def admin_demonstrativo_exportar():
    if not session.get('logged_in') or session.get('role') != 'admin':
        return redirect(url_for('login_page'))
    
    cliente_id = request.args.get('cliente_id', type=int)
    ano = request.args.get('ano', type=int)
    mes_1 = request.args.get('mes_1', type=int)
    mes_2 = request.args.get('mes_2', type=int)

    if not cliente_id or not ano:
        flash('Selecione um cliente e um ano para exportar.', 'warning')
        return redirect(url_for('admin_demonstrativo'))

    dre_data = calcular_dre_anual_refatorado(cliente_id, ano)
    
    meses_pt = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
    colunas_meses = meses_pt 
    
    meses_indices = []
    if mes_1: meses_indices.append(mes_1 - 1)
    if mes_2: meses_indices.append(mes_2 - 1)
    meses_indices.sort()

    if meses_indices:
        colunas_meses = [meses_pt[i] for i in meses_indices]
        
        def filter_values(node):
            original = node.get('monthly_values', [0.0]*12)
            node['monthly_values'] = [original[i] for i in meses_indices if i < len(original)]
            node['total'] = sum(node['monthly_values'])
            n = len(node['monthly_values'])
            node['avg'] = node['total'] / n if n > 0 else 0.0
            
            for child in node.get('children', []):
                filter_values(child)

        for root_node in dre_data.get('dre_tree', []):
            filter_values(root_node)
            
        mc_orig = dre_data.get('margem_contribuicao', [0.0]*12)
        ro_orig = dre_data.get('resultado_operacional', [0.0]*12)
        dre_data['margem_contribuicao'] = [mc_orig[i] for i in meses_indices if i < len(mc_orig)]
        dre_data['resultado_operacional'] = [ro_orig[i] for i in meses_indices if i < len(ro_orig)]

    rows = []

    def process_node(node, level=0):
        indent = "    " * level
        row = {
            "Descrição": indent + node['codigo'] + " - " + node['nome'],
        }
        for idx, mes_nome in enumerate(colunas_meses):
            val = node['monthly_values'][idx] if idx < len(node['monthly_values']) else 0.0
            row[mes_nome] = val
            
        row["Total"] = node['total']
        row["Média"] = node['avg']
        rows.append(row)

        for child in node.get('children', []):
            process_node(child, level + 1)

    for group in dre_data.get('dre_tree', []):
        process_node(group)

    row_mc = {"Descrição": "(=) MARGEM DE CONTRIBUIÇÃO"}
    for idx, mes_nome in enumerate(colunas_meses):
        val = dre_data['margem_contribuicao'][idx]
        row_mc[mes_nome] = val
    row_mc["Total"] = sum(dre_data['margem_contribuicao'])
    row_mc["Média"] = row_mc["Total"] / len(colunas_meses) if colunas_meses else 0
    rows.append(row_mc)

    row_rl = {"Descrição": "(=) RESULTADO LÍQUIDO"}
    for idx, mes_nome in enumerate(colunas_meses):
        val = dre_data['resultado_operacional'][idx]
        row_rl[mes_nome] = val
    row_rl["Total"] = sum(dre_data['resultado_operacional'])
    row_rl["Média"] = row_rl["Total"] / len(colunas_meses) if colunas_meses else 0
    rows.append(row_rl)

    df = pd.DataFrame(rows)
    cols_order = ["Descrição"] + colunas_meses + ["Total", "Média"]
    df = df[cols_order]

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='DRE')
        worksheet = writer.sheets['DRE']
        worksheet.column_dimensions['A'].width = 50 
        for col in ['B','C','D','E','F','G','H','I','J','K','L','M','N','O','P']:
             worksheet.column_dimensions[col].width = 15

    output.seek(0)
    filename = f"DRE_Cliente_{cliente_id}_{ano}.xlsx"
    return send_file(output, as_attachment=True, download_name=filename, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

@app.route('/admin/reclassificar_transacoes')
def admin_reclassificar_transacoes():
    if not session.get('logged_in') or session.get('role') != 'admin':
        return redirect(url_for('login_page'))

    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("SELECT id, username FROM users WHERE role = 'cliente' ORDER BY username")
    clientes = cursor.fetchall()
    
    cursor.execute("SELECT id, codigo, nome, user_id as cliente_id FROM plano_contas WHERE aceita_lancamentos = TRUE ORDER BY user_id, codigo")
    all_plano_contas_raw = cursor.fetchall()
    conn.close()

    plano_contas_por_cliente = defaultdict(list)
    for pc in all_plano_contas_raw:
        plano_contas_por_cliente[pc['cliente_id']].append(dict(pc))

    return render_template(
        'admin_reclassificar_transacoes.html',
        active_page='reclassificar',
        clientes=clientes,
        plano_contas_json=json.dumps(plano_contas_por_cliente)
    )

@app.route('/admin/transacao_manual', methods=['GET', 'POST'])
def admin_transacao_manual():
    if not session.get('logged_in') or session.get('role') != 'admin':
        return redirect(url_for('login_page'))

    conn = get_db()
    cursor = conn.cursor()

    if request.method == 'POST':
        try:
            cliente_id = request.form.get('cliente_id')
            conta_bancaria_id = request.form.get('conta_bancaria_id')
            tipo = request.form.get('tipo')
            data_transacao = request.form.get('data_transacao')
            valor = request.form.get('valor')
            descricao = request.form.get('descricao')
            plano_conta_id = request.form.get('plano_conta_id')

            if not all([cliente_id, conta_bancaria_id, tipo, data_transacao, valor, descricao]):
                flash('Todos os campos obrigatórios devem ser preenchidos.', 'warning')
                return redirect(request.url)

            valor_float = abs(float(valor.replace(',', '.')))

            cursor.execute("""
                INSERT INTO transacoes_financeiras 
                (cliente_id, conta_bancaria_id, plano_conta_id, data_transacao, descricao, tipo, valor) 
                VALUES (%s, %s, %s, %s, %s, %s, %s)
            """, (cliente_id, conta_bancaria_id, plano_conta_id or None, data_transacao, descricao, tipo, valor_float))
            
            conn.commit()
            flash('Transação lançada com sucesso!', 'success')
            return redirect(url_for('admin_transacao_manual'))

        except Exception as e:
            conn.rollback()
            flash(f'Erro ao lançar transação: {e}', 'danger')
            return redirect(request.url)
        finally:
            conn.close()

    cursor.execute("SELECT id, username FROM users WHERE role = 'cliente' ORDER BY username")
    clientes = cursor.fetchall()

    cursor.execute("SELECT id, apelido_conta, nome_banco, user_id FROM contas_bancarias ORDER BY apelido_conta")
    contas_raw = cursor.fetchall()
    contas_por_cliente = defaultdict(list)
    for c in contas_raw:
        contas_por_cliente[c['user_id']].append(dict(c))

    cursor.execute("SELECT id, codigo, nome, user_id FROM plano_contas WHERE aceita_lancamentos = TRUE ORDER BY codigo")
    planos_raw = cursor.fetchall()
    planos_por_cliente = defaultdict(list)
    for p in planos_raw:
        planos_por_cliente[p['user_id']].append(dict(p))

    conn.close()

    return render_template(
        'admin_transacao_manual.html',
        active_page='transacao_manual',
        clientes=clientes,
        contas_json=json.dumps(contas_por_cliente),
        planos_json=json.dumps(planos_por_cliente)
    )

@app.route('/admin/contas_bancarias', methods=['GET', 'POST'])
def admin_contas_bancarias():
    if not session.get('logged_in') or session.get('role') != 'admin':
        return redirect(url_for('login_page'))

    conn = get_db()
    check_and_add_column(conn, 'contas_bancarias', 'numero_banco', 'TEXT')
    check_and_add_column(conn, 'contas_bancarias', 'saldo_inicial', 'REAL DEFAULT 0.0')
    cursor = conn.cursor()

    if request.method == 'POST':
        action = request.form.get('action')
        cliente_id_alvo = request.form.get('cliente_id_alvo')

        if not cliente_id_alvo:
            flash('Erro: Cliente alvo não especificado.', 'danger')
            conn.close()
            return redirect(url_for('admin_contas_bancarias'))

        try:
            if action == 'add':
                cursor.execute("""
                    INSERT INTO contas_bancarias (user_id, apelido_conta, nome_banco, agencia, numero_conta, numero_banco, saldo_inicial) 
                    VALUES (%s, %s, %s, %s, %s, %s, %s)
                """, (
                    cliente_id_alvo, 
                    request.form.get('apelido_conta'), 
                    request.form.get('nome_banco'), 
                    request.form.get('agencia'), 
                    request.form.get('numero_conta'), 
                    request.form.get('numero_banco'),
                    request.form.get('saldo_inicial') or 0.0
                ))
                conn.commit()
                flash('Conta bancária adicionada com sucesso!', 'success')
            
            elif action == 'edit':
                cursor.execute("""
                    UPDATE contas_bancarias 
                    SET apelido_conta = %s, nome_banco = %s, agencia = %s, numero_conta = %s, numero_banco = %s, saldo_inicial = %s
                    WHERE id = %s AND user_id = %s
                """, (
                    request.form.get('apelido_conta_edit'), 
                    request.form.get('nome_banco_edit'), 
                    request.form.get('agencia_edit'), 
                    request.form.get('numero_conta_edit'), 
                    request.form.get('numero_banco_edit'),
                    request.form.get('saldo_inicial_edit') or 0.0,
                    request.form.get('conta_id'), 
                    cliente_id_alvo
                ))
                conn.commit()
                flash('Conta bancária atualizada com sucesso!', 'success')

            elif action == 'delete':
                cursor.execute("DELETE FROM contas_bancarias WHERE id = %s AND user_id = %s", (request.form.get('conta_id'), cliente_id_alvo))
                conn.commit()
                flash('Conta bancária apagada com sucesso!', 'success')
        
        except Exception as e:
            conn.rollback()
            flash(f'Erro na operação: {e}', 'danger')
        finally:
            conn.close()
        
        return redirect(url_for('admin_contas_bancarias', cliente_id=cliente_id_alvo))

    cursor.execute("SELECT id, username FROM users WHERE role = 'cliente' ORDER BY username")
    clientes = cursor.fetchall()
    
    cliente_id_filtro = request.args.get('cliente_id', type=int)
    contas_bancarias = []

    if cliente_id_filtro:
        cursor.execute("SELECT * FROM contas_bancarias WHERE user_id = %s ORDER BY apelido_conta", (cliente_id_filtro,))
        rows = cursor.fetchall()
        for row in rows:
            conta_dict = dict(row)
            conta_dict['_json'] = json.dumps(conta_dict)
            contas_bancarias.append(conta_dict)

    conn.close()
    return render_template(
        'admin_contas_bancarias.html',
        active_page='contas_bancarias_admin',
        clientes=clientes,
        cliente_id_filtro=cliente_id_filtro,
        contas_bancarias=contas_bancarias
    )

@app.route('/uploads/planilhas/<filename>') 
def uploaded_file(filename):
    if not session.get('logged_in'): return redirect(url_for('login_page'))
    return send_from_directory(app.config['UPLOAD_FOLDER'], filename)

@app.route('/admin/planilhas', methods=['GET', 'POST'])
def admin_planilhas():
    if not session.get('logged_in') or session.get('role') != 'admin':
        return redirect(url_for('login_page'))
    
    conn = get_db()
    cursor = conn.cursor()
    
    if request.method == 'POST':
        if 'planilha' not in request.files or request.files['planilha'].filename == '':
            flash('Nenhum arquivo selecionado.', 'warning')
            conn.close()
            return redirect(request.url)
            
        file = request.files['planilha']
        cliente_id = request.form.get('cliente_id')
        conta_bancaria_id = request.form.get('conta_bancaria_id')
        descricao_arquivo = request.form.get('descricao')

        if not cliente_id or not conta_bancaria_id:
            flash('Por favor, selecione um cliente e uma conta bancária.', 'danger')
            conn.close()
            return redirect(request.url)

        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            saved_filename = f"{timestamp}_{filename}"
            file.save(os.path.join(app.config['UPLOAD_FOLDER'], saved_filename))
            file.seek(0)
            
            try:
                cursor.execute("""
                    INSERT INTO planilhas (admin_id, cliente_id, nome_arquivo_original, nome_arquivo_servidor, descricao) 
                    VALUES (%s, %s, %s, %s, %s) RETURNING id
                """, (session['user_id'], cliente_id, filename, saved_filename, descricao_arquivo))
                planilha_id = cursor.fetchone()['id']
                conn.commit()
            except Exception as e:
                flash(f'Erro ao registrar planilha no banco: {e}', 'danger')
                conn.close()
                return redirect(request.url)

            try:
                cursor.execute("SELECT fitid FROM transacoes_financeiras WHERE cliente_id = %s AND fitid IS NOT NULL", (cliente_id,))
                fitids_existentes = {row['fitid'] for row in cursor.fetchall()}
                cursor.execute("SELECT fitid FROM transacoes_pendentes WHERE user_id = %s AND fitid IS NOT NULL", (cliente_id,))
                fitids_existentes.update({row['fitid'] for row in cursor.fetchall()})

                transacoes_temporarias = []
                duplicates_count = 0

                if file.filename.lower().endswith('.ofx'):
                    try:
                        ofx = ler_ofx_seguro(file)
                        all_accounts = ofx.accounts + getattr(ofx, 'credit_cards', [])
                        for account in all_accounts:
                            for t in account.statement.transactions:
                                if t.id in fitids_existentes:
                                    duplicates_count += 1
                                    continue
                                valor = float(t.amount)
                                transacoes_temporarias.append({
                                    'data': t.date.strftime('%Y-%m-%d'), 
                                    'descricao': t.memo, 
                                    'valor': abs(valor),
                                    'tipo': 'entrada' if valor >= 0 else 'saida', 
                                    'fitid': t.id
                                })
                    except Exception as e:
                        flash(f'Erro crítico de leitura OFX (Encoding): {e}', 'danger')
                        conn.close()
                        return redirect(request.url)

                elif file.filename.lower().endswith('.csv'):
                    try:
                        df = pd.read_csv(file, sep=None, engine='python', encoding='utf-8')
                    except UnicodeDecodeError:
                        file.seek(0)
                        df = pd.read_csv(file, sep=None, engine='python', encoding='latin-1')
                    
                    df.columns = df.columns.str.lower().str.strip()
                    col_data = next((c for c in df.columns if c in ['data', 'date', 'dt', 'data movimento', 'dia']), None)
                    col_desc = next((c for c in df.columns if c in ['descricao', 'descrição', 'historico', 'memo', 'description', 'lançamento']), None)
                    col_valor = next((c for c in df.columns if c in ['valor', 'value', 'amount', 'saldo', 'quantia']), None)

                    if not (col_data and col_desc and col_valor):
                        flash('CSV inválido. Colunas obrigatórias não encontradas.', 'danger')
                        conn.close()
                        return redirect(request.url)

                    for index, row in df.iterrows():
                        try:
                            data_raw = str(row[col_data])
                            data_obj = pd.to_datetime(data_raw, dayfirst=True, errors='coerce')
                            if pd.isna(data_obj): continue
                            data_fmt = data_obj.strftime('%Y-%m-%d')

                            valor_raw = str(row[col_valor])
                            valor_raw = valor_raw.replace('R$', '').replace('.', '').replace(',', '.').strip()
                            valor_float = float(valor_raw)

                            descricao = str(row[col_desc]).strip()
                            unique_str = f"{data_fmt}{descricao}{valor_float}{cliente_id}"
                            fitid_gerado = hashlib.md5(unique_str.encode()).hexdigest()

                            if fitid_gerado in fitids_existentes:
                                duplicates_count += 1
                                continue

                            transacoes_temporarias.append({
                                'data': data_fmt,
                                'descricao': descricao,
                                'valor': abs(valor_float),
                                'tipo': 'entrada' if valor_float >= 0 else 'saida',
                                'fitid': fitid_gerado
                            })
                        except: continue

                if duplicates_count > 0:
                    flash(f'{duplicates_count} transações duplicadas ignoradas.', 'info')
                
                if not transacoes_temporarias:
                    flash('Nenhuma transação nova encontrada.', 'warning')
                    conn.close()
                    return redirect(request.url)
                
                session['transacoes_para_mapear'] = transacoes_temporarias
                session['mapeamento_cliente_id'] = cliente_id
                session['mapeamento_conta_bancaria_id'] = conta_bancaria_id
                
                cursor.execute("UPDATE planilhas SET num_transacoes_importadas = %s WHERE id = %s", (len(transacoes_temporarias), planilha_id))
                conn.commit()
                conn.close()
                
                return redirect(url_for('admin_mapear_transacoes'))

            except Exception as e:
                flash(f'Erro ao processar o arquivo: {e}', 'danger')
                conn.close()
                return redirect(request.url)
        else:
            flash('Tipo de arquivo não permitido.', 'danger')
            conn.close()
            return redirect(request.url)
            
    cursor.execute("SELECT p.*, u_admin.username as admin_nome, u_cliente.username as cliente_nome FROM planilhas p JOIN users u_admin ON p.admin_id = u_admin.id LEFT JOIN users u_cliente ON p.cliente_id = u_cliente.id ORDER BY p.data_upload DESC")
    planilhas_list = [dict(row) for row in cursor.fetchall()]
    cursor.execute("SELECT id, username FROM users WHERE role = 'cliente' ORDER BY username")
    clientes_list = [dict(row) for row in cursor.fetchall()]
    cursor.execute("SELECT id, apelido_conta, user_id FROM contas_bancarias ORDER BY apelido_conta")
    contas_bancarias = [dict(row) for row in cursor.fetchall()]
    conn.close()
    
    return render_template('admin_planilhas.html', active_page='planilhas', planilhas=planilhas_list, clientes=clientes_list, contas_bancarias=contas_bancarias)

@app.route('/admin/mapear_transacoes')
def admin_mapear_transacoes():
    if not session.get('logged_in') or session.get('role') != 'admin':
        return redirect(url_for('login_page'))
    transacoes = session.get('transacoes_para_mapear', [])
    cliente_id = session.get('mapeamento_cliente_id')
    if not transacoes or not cliente_id:
        flash('Nenhuma transação para mapear. Por favor, importe um arquivo primeiro.', 'info')
        return redirect(url_for('admin_planilhas'))
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM tipos_conta WHERE user_id = %s ORDER BY prefixo", (cliente_id,))
    tipos_conta = [dict(row) for row in cursor.fetchall()]
    cursor.execute("SELECT * FROM plano_contas WHERE user_id = %s ORDER BY codigo", (cliente_id,))
    plano_contas_completo = [dict(row) for row in cursor.fetchall()]
    plano_contas_selecionavel = [conta for conta in plano_contas_completo if conta['aceita_lancamentos']]
    cursor.execute("SELECT username FROM users WHERE id = %s", (cliente_id,))
    cliente = cursor.fetchone()
    conn.close()
    for t in transacoes:
        t['plano_sugerido_id'] = sugerir_plano_contas_automatico(t['descricao'], cliente_id)
    return render_template(
        'admin_mapear_transacoes.html', 
        active_page='planilhas', 
        transacoes=transacoes, 
        plano_contas_selecionavel=plano_contas_selecionavel,
        plano_contas_completo=plano_contas_completo,
        tipos_conta=tipos_conta,
        cliente=cliente
    )

@app.route('/admin/salvar_transacoes_importadas', methods=['POST'])
def admin_salvar_transacoes_importadas():
    if not session.get('logged_in') or session.get('role') != 'admin':
        return redirect(url_for('login_page'))
    cliente_id = session.get('mapeamento_cliente_id')
    conta_bancaria_id = session.get('mapeamento_conta_bancaria_id')
    if not cliente_id or not conta_bancaria_id:
        flash("Sessão de importação inválida ou expirada. Por favor, comece novamente.", "danger")
        return redirect(url_for('admin_planilhas'))
    datas = request.form.getlist('data')
    descricoes = request.form.getlist('descricao')
    valores = request.form.getlist('valor')
    tipos = request.form.getlist('tipo')
    fitids = request.form.getlist('fitid')
    planos_conta_ids = request.form.getlist('plano_conta_id')
    transacoes_para_inserir = []
    transacoes_para_pendencia = []
    for i in range(len(datas)):
        if planos_conta_ids[i]:
            transacoes_para_inserir.append((
                cliente_id, conta_bancaria_id, planos_conta_ids[i], 
                datas[i], descricoes[i], tipos[i], 
                float(valores[i]), fitids[i]
            ))
        else:
            transacoes_para_pendencia.append((
                cliente_id, conta_bancaria_id, datas[i], 
                descricoes[i], float(valores[i]), tipos[i], fitids[i]
            ))
    conn = get_db()
    cursor = conn.cursor()
    if transacoes_para_inserir:
        try:
            cursor.executemany("""
                INSERT INTO transacoes_financeiras 
                (cliente_id, conta_bancaria_id, plano_conta_id, data_transacao, descricao, tipo, valor, fitid) 
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s) ON CONFLICT (cliente_id, fitid) DO NOTHING
            """, transacoes_para_inserir)
        except psycopg2.Error as e:
            conn.rollback()
            flash(f"Erro ao salvar transações. Detalhe: {e}", "danger")
            conn.close()
            return redirect(url_for('admin_planilhas'))
    if transacoes_para_pendencia:
        try:
            cursor.executemany("""
                INSERT INTO transacoes_pendentes 
                (user_id, conta_bancaria_id, data_transacao, descricao, valor, tipo, fitid) 
                VALUES (%s, %s, %s, %s, %s, %s, %s) ON CONFLICT (user_id, fitid) DO NOTHING
            """, transacoes_para_pendencia)
        except psycopg2.Error as e:
            conn.rollback()
            flash(f"Erro ao salvar pendências. Detalhe: {e}", "danger")
            conn.close()
            return redirect(url_for('admin_planilhas'))
    conn.commit()
    conn.close()
    flash(f'{len(transacoes_para_inserir)} transações importadas e {len(transacoes_para_pendencia)} salvas como pendentes para o cliente.', 'success')
    session.pop('transacoes_para_mapear', None)
    session.pop('mapeamento_cliente_id', None)
    session.pop('mapeamento_conta_bancaria_id', None)
    return redirect(url_for('admin_transacoes_pendentes', cliente_id=cliente_id))

@app.route('/admin/transacoes_pendentes', methods=['GET', 'POST'])
def admin_transacoes_pendentes():
    if not session.get('logged_in') or session.get('role') != 'admin':
        return redirect(url_for('login_page'))
    conn = get_db()
    cursor = conn.cursor()
    
    if request.method == 'POST':
        pendencia_id = request.form.get('pendencia_id')
        plano_conta_id = request.form.get('plano_conta_id')
        cliente_id_alvo = request.form.get('cliente_id_alvo')
        if pendencia_id and plano_conta_id and cliente_id_alvo:
            cursor.execute("SELECT * FROM transacoes_pendentes WHERE id = %s", (pendencia_id,))
            pendencia = cursor.fetchone()
            if pendencia:
                cursor.execute("""
                    INSERT INTO transacoes_financeiras (cliente_id, conta_bancaria_id, plano_conta_id, data_transacao, descricao, tipo, valor, fitid) 
                    VALUES (%s, %s, %s, %s, %s, %s, %s, %s) ON CONFLICT (cliente_id, fitid) DO NOTHING
                """, (cliente_id_alvo, pendencia['conta_bancaria_id'], plano_conta_id, pendencia['data_transacao'], pendencia['descricao'], pendencia['tipo'], pendencia['valor'], pendencia['fitid']))
                cursor.execute("DELETE FROM transacoes_pendentes WHERE id = %s", (pendencia_id,))
                conn.commit()
                flash('Transação categorizada com sucesso para o cliente!', 'success')
        conn.close()
        return redirect(url_for('admin_transacoes_pendentes', cliente_id=cliente_id_alvo))
        
    cursor.execute("SELECT id, username FROM users WHERE role = 'cliente' ORDER BY username")
    clientes = cursor.fetchall()
    cliente_id_filtro = request.args.get('cliente_id', type=int)
    pendencias = []
    plano_contas = []
    
    if cliente_id_filtro:
        cursor.execute("""
            SELECT tp.*, cb.apelido_conta 
            FROM transacoes_pendentes tp 
            LEFT JOIN contas_bancarias cb ON tp.conta_bancaria_id = cb.id 
            WHERE tp.user_id = %s 
            ORDER BY tp.data_transacao DESC
        """, (cliente_id_filtro,))
        pendencias = cursor.fetchall()
        
        cursor.execute("SELECT id, codigo, nome FROM plano_contas WHERE user_id = %s AND aceita_lancamentos = TRUE ORDER BY codigo", (cliente_id_filtro,))
        plano_contas = cursor.fetchall()
    
    conn.close()
    return render_template('admin_transacoes_pendentes.html', active_page='pendentes', clientes=clientes, cliente_id_filtro=cliente_id_filtro, pendencias=pendencias, plano_contas=plano_contas)

@app.route('/admin/planilhas/apagar/<int:planilha_id>', methods=['POST'])
def admin_apagar_planilha(planilha_id):
    if not session.get('logged_in') or session.get('role') != 'admin': flash('Acesso não autorizado.', 'danger'); return redirect(url_for('login_page'))
    conn = get_db(); cursor = conn.cursor()
    cursor.execute("SELECT nome_arquivo_servidor FROM planilhas WHERE id = %s", (planilha_id,)); planilha = cursor.fetchone()
    if planilha:
        try:
            cursor.execute("DELETE FROM transacoes_financeiras WHERE planilha_id = %s", (planilha_id,))
            cursor.execute("DELETE FROM planilhas WHERE id = %s", (planilha_id,)); conn.commit()
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], planilha['nome_arquivo_servidor'])
            if os.path.exists(file_path): os.remove(file_path)
            flash('Planilha e transações associadas foram apagadas com sucesso.', 'success')
        except Exception as e: conn.rollback(); flash(f'Erro ao apagar a planilha: {e}', 'danger')
        finally: conn.close()
    else: flash('Planilha não encontrada.', 'warning'); conn.close()
    return redirect(url_for('admin_planilhas'))

@app.route('/admin/plano_de_contas', methods=['GET', 'POST'])
def admin_plano_de_contas():
    if not session.get('logged_in') or session.get('role') != 'admin':
        return redirect(url_for('login_page'))
    
    conn = get_db()
    cursor = conn.cursor()

    if request.method == 'POST':
        action = request.form.get('action')
        user_id = request.form.get('cliente_id_alvo', type=int)
        
        if not user_id:
            flash("Erro: Cliente alvo não especificado.", "danger")
            conn.close()
            return redirect(url_for('admin_plano_de_contas'))

        try:
            if action == 'add_or_edit_conta':
                codigo = request.form.get('codigo').strip()
                nome = request.form.get('nome')
                conta_id = request.form.get('conta_id')
                aceita_lancamentos = 'aceita_lancamentos' in request.form

                parent_code = '.'.join(codigo.split('.')[:-1])
                if parent_code: 
                    cursor.execute("SELECT id FROM plano_contas WHERE user_id = %s AND codigo = %s", (user_id, parent_code))
                    if not cursor.fetchone():
                        flash(f'Erro: A conta pai com código "{parent_code}" não existe.', 'danger')
                        conn.close()
                        return redirect(url_for('admin_plano_de_contas', cliente_id=user_id))
                
                tipo_conta = get_tipo_by_codigo(codigo, user_id)

                if conta_id:
                    cursor.execute(
                        "UPDATE plano_contas SET nome = %s, tipo = %s, aceita_lancamentos = %s WHERE id = %s AND user_id = %s",
                        (nome, tipo_conta, aceita_lancamentos, conta_id, user_id)
                    )
                    flash('Conta atualizada com sucesso!', 'success')
                else:
                    cursor.execute(
                        "INSERT INTO plano_contas (user_id, codigo, nome, tipo, aceita_lancamentos) VALUES (%s, %s, %s, %s, %s)",
                        (user_id, codigo, nome, tipo_conta, aceita_lancamentos)
                    )
                    flash('Conta adicionada com sucesso!', 'success')

            elif action == 'delete_conta':
                conta_id_para_deletar = request.form.get('conta_id')
                cursor.execute("SELECT codigo FROM plano_contas WHERE id = %s", (conta_id_para_deletar,))
                conta = cursor.fetchone()
                if conta:
                    cursor.execute("SELECT 1 FROM plano_contas WHERE user_id = %s AND codigo LIKE %s LIMIT 1", (user_id, conta['codigo'] + '.%'))
                    if cursor.fetchone():
                        flash('Erro: Não é possível apagar uma conta que possui sub-contas.', 'danger')
                    else:
                        cursor.execute("SELECT 1 FROM transacoes_financeiras WHERE plano_conta_id = %s LIMIT 1", (conta_id_para_deletar,))
                        if cursor.fetchone():
                            flash('Erro: Não é possível apagar uma conta que já possui transações lançadas.', 'danger')
                        else:
                            cursor.execute("DELETE FROM plano_contas WHERE id = %s AND user_id = %s", (conta_id_para_deletar, user_id))
                            flash('Conta apagada com sucesso!', 'success')
            
            conn.commit()
        except psycopg2.IntegrityError:
            conn.rollback()
            flash(f'Erro de integridade: O código de conta já existe para este cliente.', 'danger')
        except Exception as e:
            conn.rollback()
            flash(f'Ocorreu um erro inesperado: {e}', 'danger')
        
        conn.close()
        return redirect(url_for('admin_plano_de_contas', cliente_id=user_id))

    cursor.execute("SELECT id, username FROM users WHERE role = 'cliente' ORDER BY username")
    clientes = cursor.fetchall()
    cliente_id_filtro = request.args.get('cliente_id', type=int)
    
    plano_contas_tree = []
    if cliente_id_filtro:
        cursor.execute("SELECT * FROM plano_contas WHERE user_id = %s ORDER BY codigo", (cliente_id_filtro,))
        contas = [dict(c) for c in cursor.fetchall()]
        
        contas_dict = {c['codigo']: c for c in contas}
        root_contas = []
        
        for conta in sorted(contas, key=lambda x: x['codigo']):
            conta['children'] = []
            parent_code = '.'.join(conta['codigo'].split('.')[:-1])

            if parent_code and parent_code in contas_dict:
                contas_dict[parent_code]['children'].append(conta)
            else:
                root_contas.append(conta)
        plano_contas_tree = root_contas

    conn.close()
    return render_template('admin_plano_de_contas.html', active_page='plano_de_contas', clientes=clientes, cliente_id_filtro=cliente_id_filtro, plano_contas_tree=plano_contas_tree)

@app.route('/admin/demonstrativo')
def admin_demonstrativo():
    if not session.get('logged_in') or session.get('role') != 'admin':
        return redirect(url_for('login_page'))
    
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("SELECT id, username FROM users WHERE role = 'cliente' ORDER BY username")
    clientes = cursor.fetchall()
    conn.close()

    cliente_id_filtro = request.args.get('cliente_id', type=int)
    ano_selecionado = request.args.get('ano', type=int)
    mes_1_filtro = request.args.get('mes_1', type=int)
    mes_2_filtro = request.args.get('mes_2', type=int)
    
    anos_disponiveis = range(datetime.now().year + 1, datetime.now().year - 5, -1)
    meses_pt = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
    
    dre_data_processada = {}
    meses_exibidos = [] 

    if cliente_id_filtro and ano_selecionado:
        dre_data_completa = calcular_dre_anual_refatorado(user_id=cliente_id_filtro, ano=ano_selecionado)
        meses_indices_filtro = []
        if mes_1_filtro: meses_indices_filtro.append(mes_1_filtro - 1)
        if mes_2_filtro: meses_indices_filtro.append(mes_2_filtro - 1)
        meses_indices_filtro.sort()

        if meses_indices_filtro:
            meses_exibidos = [meses_pt[i] for i in meses_indices_filtro]
            def filter_node_values(node):
                original_values = node.get('monthly_values', [0.0]*12)
                node['monthly_values'] = [original_values[i] for i in meses_indices_filtro if i < len(original_values)]
                node['total'] = sum(node['monthly_values'])
                num_months = len(node['monthly_values'])
                node['avg'] = (node['total'] / num_months) if num_months > 0 else 0.0
                for child in node.get('children', []):
                    filter_node_values(child)

            for root_node in dre_data_completa.get('dre_tree', []):
                filter_node_values(root_node)

            mc_original = dre_data_completa.get('margem_contribuicao', [0.0]*12)
            ro_original = dre_data_completa.get('resultado_operacional', [0.0]*12)
            dre_data_completa['margem_contribuicao'] = [mc_original[i] for i in meses_indices_filtro if i < len(mc_original)]
            dre_data_completa['resultado_operacional'] = [ro_original[i] for i in meses_indices_filtro if i < len(ro_original)]
            dre_data_processada = dre_data_completa
        else:
            meses_exibidos = meses_pt
            dre_data_processada = dre_data_completa

    return render_template(
        'admin_demonstrativo.html',
        active_page='demonstrativo',
        clientes=clientes,
        cliente_id_filtro=cliente_id_filtro,
        ano_selecionado=ano_selecionado,
        anos_disponiveis=anos_disponiveis,
        meses_pt=meses_pt, 
        meses_exibidos=meses_exibidos, 
        mes_1_filtro=mes_1_filtro,
        mes_2_filtro=mes_2_filtro,
        **dre_data_processada
    )

@app.route('/admin/fluxo_caixa')
def admin_fluxo_caixa():
    if not session.get('logged_in') or session.get('role') != 'admin':
        return redirect(url_for('login_page'))
    
    conn = get_db()
    cursor = conn.cursor()
    
    cursor.execute("SELECT id, username FROM users WHERE role = 'cliente' ORDER BY username")
    clientes = [dict(row) for row in cursor.fetchall()]
    
    cursor.execute("SELECT id, nome_arquivo_original FROM planilhas ORDER BY data_upload DESC")
    planilhas = [dict(row) for row in cursor.fetchall()]
    
    cursor.execute("SELECT cb.id, cb.apelido_conta, u.username FROM contas_bancarias cb JOIN users u ON cb.user_id = u.id ORDER BY u.username, cb.apelido_conta")
    contas_bancarias = [dict(row) for row in cursor.fetchall()]
    
    cursor.execute("SELECT id, codigo, nome, user_id as cliente_id FROM plano_contas WHERE aceita_lancamentos = TRUE ORDER BY user_id, codigo")
    all_plano_contas_raw = cursor.fetchall()
    plano_contas_por_cliente = defaultdict(list)
    for pc in all_plano_contas_raw:
        plano_contas_por_cliente[pc['cliente_id']].append(dict(pc))

    conn.close()
    
    return render_template(
        'admin_fluxo_caixa.html',
        active_page='fluxo_caixa',
        clientes=clientes,
        planilhas=planilhas,
        contas_bancarias=contas_bancarias,
        plano_contas_json=json.dumps(plano_contas_por_cliente)
    )

@app.route('/admin/analise_financeira')
def admin_analise_financeira():
    if not session.get('logged_in') or session.get('role') != 'admin': return redirect(url_for('login_page'))
    conn = get_db(); cursor = conn.cursor()
    cursor.execute("SELECT id, username FROM users WHERE role = 'cliente' ORDER BY username"); clientes = [dict(row) for row in cursor.fetchall()]; conn.close()
    return render_template('admin_analise_financeira.html', active_page='analise_financeira', clientes=clientes)

@app.route('/admin/regras_mapeamento', methods=['GET', 'POST'])
def admin_regras_mapeamento():
    if not session.get('logged_in') or session.get('role') != 'admin':
        return redirect(url_for('login_page'))

    conn = get_db()
    cursor = conn.cursor()

    if request.method == 'POST':
        action = request.form.get('action')
        cliente_id_alvo = request.form.get('cliente_id_alvo')

        if action == 'add':
            texto_chave = request.form.get('texto_chave')
            plano_conta_id = request.form.get('plano_conta_id')
            if texto_chave and plano_conta_id and cliente_id_alvo:
                try:
                    cursor.execute(
                        "INSERT INTO mapeamento_regras (user_id, texto_chave, plano_conta_id) VALUES (%s, %s, %s)",
                        (cliente_id_alvo, texto_chave.strip().upper(), plano_conta_id)
                    )
                    conn.commit()
                    flash('Regra de mapeamento adicionada com sucesso!', 'success')
                except psycopg2.IntegrityError:
                    conn.rollback()
                    flash(f'Erro: A regra para "{texto_chave}" já existe para este cliente.', 'danger')

        elif action == 'delete':
            regra_id = request.form.get('regra_id')
            if regra_id:
                cursor.execute("DELETE FROM mapeamento_regras WHERE id = %s", (regra_id,))
                conn.commit()
                flash('Regra apagada com sucesso.', 'success')
        
        conn.close()
        return redirect(url_for('admin_regras_mapeamento', cliente_id=cliente_id_alvo))

    cursor.execute("SELECT id, username FROM users WHERE role = 'cliente' ORDER BY username")
    clientes = cursor.fetchall()
    
    cliente_id_filtro = request.args.get('cliente_id', type=int)
    regras = []
    plano_contas_selecionavel = []

    if cliente_id_filtro:
        cursor.execute("""
            SELECT mr.id, mr.texto_chave, pc.nome as plano_conta_nome, pc.codigo as plano_conta_codigo
            FROM mapeamento_regras mr
            JOIN plano_contas pc ON mr.plano_conta_id = pc.id
            WHERE mr.user_id = %s 
            ORDER BY mr.texto_chave
        """, (cliente_id_filtro,))
        regras = cursor.fetchall()

        cursor.execute(
            "SELECT id, codigo, nome FROM plano_contas WHERE user_id = %s AND aceita_lancamentos = TRUE ORDER BY codigo",
            (cliente_id_filtro,)
        )
        plano_contas_selecionavel = cursor.fetchall()

    conn.close()
    
    return render_template(
        'admin_regras_mapeamento.html',
        active_page='regras_mapeamento',
        clientes=clientes,
        cliente_id_filtro=cliente_id_filtro,
        regras=regras,
        plano_contas_selecionavel=plano_contas_selecionavel
    )

@app.route('/admin/assinaturas', methods=['GET', 'POST'])
def admin_assinaturas():
    if not session.get('logged_in') or session.get('role') != 'admin': return redirect(url_for('login_page'))
    conn = get_db(); cursor = conn.cursor()
    if request.method == 'POST':
        action, cliente_id = request.form.get('action'), request.form.get('cliente_id')
        if not cliente_id or not cliente_id.isdigit(): flash('ID de Cliente inválido.', 'danger'); conn.close(); return redirect(url_for('admin_assinaturas'))
        if action == 'add_or_update':
            data_inicio, data_fim, status, plano_nome = request.form.get('data_inicio'), request.form.get('data_fim'), request.form.get('status'), request.form.get('plano_nome')
            if not all([data_inicio, data_fim, status]): flash('Todos os campos da assinatura são obrigatórios.', 'warning')
            else:
                try:
                    data_inicio_obj, data_fim_obj = datetime.strptime(data_inicio, '%Y-%m-%d').date(), datetime.strptime(data_fim, '%Y-%m-%d').date()
                    cursor.execute("SELECT id FROM assinaturas WHERE cliente_id = %s", (cliente_id,));
                    if cursor.fetchone(): 
                        cursor.execute("UPDATE assinaturas SET data_inicio = %s, data_fim = %s, status = %s, plano_nome = %s WHERE cliente_id = %s", (data_inicio_obj, data_fim_obj, status, plano_nome, cliente_id)); flash('Assinatura atualizada!', 'success')
                    else: 
                        cursor.execute("INSERT INTO assinaturas (cliente_id, data_inicio, data_fim, status, plano_nome) VALUES (%s, %s, %s, %s, %s)", (cliente_id, data_inicio_obj, data_fim_obj, status, plano_nome)); flash('Assinatura adicionada!', 'success')
                    conn.commit()
                except ValueError: flash('Formato de data inválido.', 'danger')
        elif action == 'cancel': cursor.execute("UPDATE assinaturas SET status = 'cancelada' WHERE id = %s", (request.form.get('assinatura_id_cancel'),)); conn.commit(); flash('Assinatura cancelada.', 'info')
        conn.close()
        return redirect(url_for('admin_assinaturas'))
        
    cursor.execute("SELECT a.*, u.username as cliente_nome, u.id as cliente_id FROM users u LEFT JOIN assinaturas a ON u.id = a.cliente_id WHERE u.role = 'cliente' ORDER BY u.username"); clientes_com_assinaturas_raw = cursor.fetchall()
    clientes_com_assinaturas_processed = []
    for row_data in clientes_com_assinaturas_raw:
        processed_row = dict(row_data) 
        if processed_row.get('data_inicio') and not isinstance(processed_row['data_inicio'], date):
            try: processed_row['data_inicio'] = datetime.strptime(processed_row['data_inicio'], '%Y-%m-%d').date()
            except (ValueError, TypeError): processed_row['data_inicio'] = None 
        if processed_row.get('data_fim') and not isinstance(processed_row['data_fim'], date):
            try: processed_row['data_fim'] = datetime.strptime(processed_row['data_fim'], '%Y-%m-%d').date()
            except (ValueError, TypeError): processed_row['data_fim'] = None
        clientes_com_assinaturas_processed.append(processed_row)
    cursor.execute("SELECT u.id, u.username FROM users u LEFT JOIN assinaturas a ON u.id = a.cliente_id WHERE u.role = 'cliente' AND a.id IS NULL"); clientes_sem_assinatura = [dict(row) for row in cursor.fetchall()]; conn.close()
    return render_template('admin_assinaturas.html', active_page='assinaturas', clientes_assinaturas=clientes_com_assinaturas_processed, clientes_sem_assinatura=clientes_sem_assinatura)

@app.route('/admin/clientes', methods=['GET'])
def admin_listar_clientes():
    if not session.get('logged_in') or session.get('role') != 'admin': return redirect(url_for('login_page'))
    conn = get_db(); cursor = conn.cursor()
    cursor.execute("SELECT * FROM users WHERE role = 'cliente' ORDER BY username"); clientes = [dict(row) for row in cursor.fetchall()]; conn.close()
    return render_template('admin_listar_clientes.html', active_page='clientes_listar', clientes=clientes)

@app.route('/admin/utilizadores')
def admin_utilizadores():
    if not session.get('logged_in') or session.get('role') != 'admin': return redirect(url_for('login_page'))
    conn = get_db(); cursor = conn.cursor()
    cursor.execute("SELECT id, username, email, role FROM users ORDER BY username"); users = [dict(row) for row in cursor.fetchall()]; conn.close()
    return render_template('admin_utilizadores.html', active_page='utilizadores', users=users)

@app.route('/admin/configuracoes')
def admin_configuracoes():
    if not session.get('logged_in') or session.get('role') != 'admin': return redirect(url_for('login_page'))
    return render_template('admin_configuracoes.html', active_page='configuracoes')

@app.route('/admin/clientes/cadastrar', methods=['GET', 'POST'])
def admin_cadastrar_cliente():
    if not session.get('logged_in') or session.get('role') != 'admin':
        return redirect(url_for('login_page'))
    
    conn = get_db()
    check_and_add_column(conn, 'users', 'role', 'TEXT')
    
    if request.method == 'POST':
        form = request.form
        username = form.get('username')
        password = form.get('password')
        email = form.get('email')
        role = form.get('role', 'cliente')
        
        if not all([username, password, email, role]):
            flash('Todos os campos de acesso são obrigatórios.', 'warning')
            return redirect(url_for('admin_cadastrar_cliente'))

        cursor = conn.cursor() 
        try:
            hashed_pw = generate_password_hash(password)

            cursor.execute("""
                INSERT INTO users (
                    username, password, role, email, telefone, tipo_pessoa, cpf_cnpj, 
                    razao_social, nome_fantasia, endereco_rua, endereco_numero, 
                    endereco_bairro, endereco_cidade, endereco_cep
                ) 
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s) 
                RETURNING id
            """, (
                username, hashed_pw, role, email, form.get('telefone'), form.get('tipo_pessoa'),
                form.get('cpf_cnpj'), form.get('razao_social'), form.get('nome_fantasia'),
                form.get('endereco_rua'), form.get('endereco_numero'),
                form.get('endereco_bairro'), form.get('endereco_cidade'), form.get('endereco_cep')
            ))
            
            result = cursor.fetchone()
            new_id = result['id'] if isinstance(result, dict) else result[0]
            
            if role == 'cliente':
                criar_plano_contas_padrao(cursor, new_id)
                cursor.execute("""
                    INSERT INTO assinaturas (cliente_id, data_inicio, data_fim, status, plano_nome) 
                    VALUES (%s, %s, %s, 'ativa', 'Plano Inicial')
                """, (new_id, date.today(), date.today() + timedelta(days=30)))
            
            conn.commit()
            flash(f'Usuário "{username}" ({role}) cadastrado com sucesso!', 'success')
            return redirect(url_for('admin_listar_clientes'))
            
        except Exception as e:
            conn.rollback()
            flash(f'Erro no cadastro: {e}', 'danger')
        finally:
            conn.close()

    return render_template('admin_cadastrar_cliente.html', active_page='clientes_cadastrar')

@app.route('/admin/usuarios/excluir/<int:user_id>', methods=['POST'])
def admin_excluir_usuario(user_id):
    if not session.get('logged_in') or session.get('role') != 'admin':
        return redirect(url_for('login_page'))

    if user_id == session.get('user_id'):
        flash('Erro: Não pode excluir a sua própria conta de administrador.', 'danger')
        return redirect(url_for('admin_utilizadores'))

    conn = get_db()
    cursor = conn.cursor()
    try:
        cursor.execute("DELETE FROM users WHERE id = %s", (user_id,))
        conn.commit()
        flash('Utilizador e todos os dados associados foram removidos permanentemente.', 'success')
    except Exception as e:
        conn.rollback()
        flash(f'Erro ao excluir utilizador: {e}', 'danger')
    finally:
        conn.close()

    return redirect(url_for('admin_utilizadores'))

# --- ROTAS DE API ---

@app.route('/api/dados_fluxo_caixa')
def api_dados_fluxo_caixa():
    if not session.get('logged_in'): return jsonify({'error': 'Acesso não autorizado'}), 403
    
    args = request.args
    data_inicio, data_fim = args.get('data_inicio'), args.get('data_fim')
    cliente_id_filtro, planilha_id_filtro, conta_bancaria_id_filtro = args.get('cliente_id'), args.get('planilha_id'), args.get('conta_bancaria_id')
    
    conn = get_db(); cursor = conn.cursor(); base_query = "FROM transacoes_financeiras tf"; where_clauses, params = [], []
    
    if session['role'] == 'cliente': 
        where_clauses.append("tf.cliente_id = %s"); params.append(session['user_id'])
    elif session['role'] == 'admin' and cliente_id_filtro and cliente_id_filtro != 'todos': 
        where_clauses.append("tf.cliente_id = %s"); params.append(int(cliente_id_filtro))
    
    if planilha_id_filtro and planilha_id_filtro != 'todas': where_clauses.append("tf.planilha_id = %s"); params.append(int(planilha_id_filtro))
    if conta_bancaria_id_filtro and conta_bancaria_id_filtro != 'todas': where_clauses.append("tf.conta_bancaria_id = %s"); params.append(int(conta_bancaria_id_filtro))
    if data_inicio: where_clauses.append("tf.data_transacao >= %s"); params.append(data_inicio)
    if data_fim: where_clauses.append("tf.data_transacao <= %s"); params.append(data_fim)
    
    query_filter = f"WHERE {' AND '.join(where_clauses)}" if where_clauses else ""
    
    cursor.execute(f"SELECT tipo, SUM(valor) as total {base_query} {query_filter} GROUP BY tipo", params); entradas_saidas = {row['tipo']: (row['total'] or 0) for row in cursor.fetchall()}
    pizza_data = {'labels': ['Total Entradas', 'Total Saídas'], 'valores': [entradas_saidas.get('entrada', 0), entradas_saidas.get('saida', 0)]}
    
    cursor.execute(f"SELECT TO_CHAR(data_transacao, 'YYYY-MM') as mes, tipo, SUM(valor) as total {base_query} {query_filter} GROUP BY mes, tipo ORDER BY mes", params); line_data_raw = cursor.fetchall()
    line_data_processed = {}
    for row in line_data_raw:
        mes, tipo, total = row['mes'], row['tipo'], (row['total'] or 0)
        if mes not in line_data_processed: line_data_processed[mes] = {'entradas': 0, 'saidas': 0}
        line_data_processed[mes][tipo + 's'] = total
    line_data = {'labels': list(line_data_processed.keys()), 'entradas': [v['entradas'] for v in line_data_processed.values()], 'saidas': [v['saidas'] for v in line_data_processed.values()]}
    
    query_transacoes = f"""
        SELECT 
            tf.id, tf.data_transacao, tf.descricao, tf.tipo, tf.valor,
            tf.cliente_id, tf.plano_conta_id,
            u.username as cliente_username,
            pc.codigo as plano_conta_codigo,
            pc.nome as plano_conta_nome,
            cb.apelido_conta as banco_nome
        FROM transacoes_financeiras tf 
        LEFT JOIN users u ON tf.cliente_id = u.id
        LEFT JOIN plano_contas pc ON tf.plano_conta_id = pc.id
        LEFT JOIN contas_bancarias cb ON tf.conta_bancaria_id = cb.id
        {query_filter} 
        ORDER BY tf.data_transacao DESC, tf.id DESC
        LIMIT 100 
    """
    cursor.execute(query_transacoes, params)
    transacoes_recentes = [dict(row) for row in cursor.fetchall()]

    conn.close()
    
    return jsonify({'pizzaData': pizza_data, 'lineData': line_data, 'tableData': transacoes_recentes})

@app.route('/api/admin/transacao/atualizar_plano_conta', methods=['POST'])
def api_admin_atualizar_plano_conta():
    if not session.get('logged_in') or session.get('role') != 'admin':
        return jsonify({'success': False, 'error': 'Acesso não autorizado'}), 403
    
    data = request.json
    transacao_id = data.get('transacao_id')
    plano_conta_id = data.get('plano_conta_id')

    if not plano_conta_id or str(plano_conta_id) == '0':
        plano_conta_id = None
    
    if not transacao_id:
        return jsonify({'success': False, 'error': 'ID da transação não fornecido.'}), 400

    try:
        conn = get_db()
        cursor = conn.cursor()
        cursor.execute(
            "UPDATE transacoes_financeiras SET plano_conta_id = %s WHERE id = %s",
            (plano_conta_id, transacao_id)
        )
        conn.commit()
        conn.close()
        return jsonify({'success': True, 'message': 'Plano de conta atualizado com sucesso!'})
    except Exception as e:
        return jsonify({'success': False, 'error': f'Erro no servidor: {e}'}), 500

@app.route('/api/admin/pendencia/excluir', methods=['POST'])
def api_admin_excluir_pendencia():
    if not session.get('logged_in') or session.get('role') != 'admin':
        return jsonify({'success': False, 'error': 'Acesso não autorizado'}), 403
    
    data = request.json
    pendencia_id = data.get('pendencia_id')

    if not pendencia_id:
        return jsonify({'success': False, 'error': 'ID da pendência não fornecido.'}), 400

    try:
        conn = get_db()
        cursor = conn.cursor()
        cursor.execute("DELETE FROM transacoes_pendentes WHERE id = %s", (pendencia_id,))
        conn.commit()
        conn.close()
        return jsonify({'success': True, 'message': 'Pendência excluída com sucesso!'})
    except Exception as e:
        return jsonify({'success': False, 'error': f'Erro no servidor: {e}'}), 500

@app.route('/api/admin/transacao/excluir', methods=['POST'])
def api_admin_excluir_transacao():
    if not session.get('logged_in') or session.get('role') != 'admin':
        return jsonify({'success': False, 'error': 'Acesso não autorizado'}), 403
    
    data = request.json
    transacao_id = data.get('transacao_id')

    if not transacao_id:
        return jsonify({'success': False, 'error': 'ID da transação não fornecido.'}), 400

    try:
        conn = get_db()
        cursor = conn.cursor()
        cursor.execute("SELECT id FROM transacoes_financeiras WHERE id = %s", (transacao_id,))
        if not cursor.fetchone():
             conn.close()
             return jsonify({'success': False, 'error': 'Transação não encontrada.'}), 404

        cursor.execute("DELETE FROM transacoes_financeiras WHERE id = %s", (transacao_id,))
        conn.commit()
        conn.close()
        return jsonify({'success': True, 'message': 'Transação excluída com sucesso!'})
    except Exception as e:
        return jsonify({'success': False, 'error': f'Erro no servidor: {e}'}), 500

@app.route('/api/admin/transacoes_para_reclassificar')
def api_admin_transacoes_para_reclassificar():
    if not session.get('logged_in') or session.get('role') != 'admin':
        return jsonify({'error': 'Não autorizado'}), 403

    args = request.args
    cliente_id = args.get('cliente_id', type=int)
    plano_conta_id = args.get('plano_conta_atual_id', type=int)
    data_inicio = args.get('data_inicio')
    data_fim = args.get('data_fim')
    descricao = args.get('descricao')

    if not cliente_id:
        return jsonify([])

    conn = get_db()
    cursor = conn.cursor()
    
    where_clauses = ["tf.cliente_id = %s"]
    params = [cliente_id]

    if plano_conta_id:
        where_clauses.append("tf.plano_conta_id = %s")
        params.append(plano_conta_id)
    if data_inicio:
        where_clauses.append("tf.data_transacao >= %s")
        params.append(data_inicio)
    if data_fim:
        where_clauses.append("tf.data_transacao <= %s")
        params.append(data_fim)
    if descricao:
        where_clauses.append("tf.descricao ILIKE %s")
        params.append(f"%{descricao}%")
    
    query_filter = f"WHERE {' AND '.join(where_clauses)}"

    query = f"""
        SELECT tf.id, tf.data_transacao, tf.descricao, tf.tipo, tf.valor, tf.cliente_id, tf.plano_conta_id
        FROM transacoes_financeiras tf
        {query_filter}
        ORDER BY tf.data_transacao DESC, tf.id DESC
        LIMIT 200
    """
    
    cursor.execute(query, params)
    transacoes = [dict(row) for row in cursor.fetchall()]
    conn.close()

    return jsonify(transacoes)

@app.route('/api/cliente/notificacoes')
def api_cliente_notificacoes():
    if not session.get('logged_in') or session.get('role') != 'cliente': return jsonify({'error': 'Não autorizado'}), 403
    user_id = session['user_id']; conn = get_db(); cursor = conn.cursor()
    cursor.execute("SELECT id, mensagem, url, data_criacao FROM notificacoes WHERE user_id = %s AND lida = FALSE ORDER BY data_criacao DESC LIMIT 5", (user_id,)); notificacoes = [dict(row) for row in cursor.fetchall()]
    cursor.execute("SELECT COUNT(id) as total FROM notificacoes WHERE user_id = %s AND lida = FALSE", (user_id,)); total_nao_lidas = cursor.fetchone()['total']; conn.close()
    return jsonify({'notificacoes': notificacoes, 'total_nao_lidas': total_nao_lidas})

@app.route('/api/cliente/transacao/atualizar_categoria', methods=['POST'])
def cliente_atualizar_categoria():
    if not session.get('logged_in') or session.get('role') != 'cliente': return jsonify({'success': False, 'error': 'Acesso não autorizado'}), 403
    data = request.json; transacao_id, nova_categoria, cliente_id = data.get('transacao_id'), data.get('nova_categoria'), session.get('user_id')
    if not transacao_id or not nova_categoria: return jsonify({'success': False, 'error': 'Dados em falta'}), 400
    conn = get_db(); cursor = conn.cursor(); cursor.execute("SELECT categoria FROM transacoes_financeiras WHERE id = %s AND cliente_id = %s", (transacao_id, cliente_id)); transacao = cursor.fetchone()
    if not transacao: conn.close(); return jsonify({'success': False, 'error': 'Transação não encontrada ou não pertence a este utilizador'}), 404
    categoria_antiga = transacao['categoria']
    if categoria_antiga != nova_categoria:
        cursor.execute("UPDATE transacoes_financeiras SET categoria = %s WHERE id = %s", (nova_categoria, transacao_id))
        cursor.execute("INSERT INTO categoria_alteracoes_log (transacao_id, user_id, categoria_antiga, categoria_nova) VALUES (%s, %s, %s, %s)", (transacao_id, cliente_id, categoria_antiga, nova_categoria)); conn.commit()
    conn.close()
    return jsonify({'success': True, 'message': 'Categoria atualizada com sucesso.'})

@app.route('/cliente/transacao_manual', methods=['POST'])
def cliente_adicionar_transacao_manual():
    if not session.get('logged_in') or session.get('role') != 'cliente': return redirect(url_for('login_page'))
    try:
        conta_bancaria_id, data_transacao, descricao, tipo, valor, categoria, plano_conta_id = request.form.get('conta_bancaria_id') or None, request.form['data_transacao'], request.form['descricao'], request.form['tipo'], abs(float(request.form['valor'])), request.form['categoria'], request.form.get('plano_conta_id') or None
        conn = get_db(); cursor = conn.cursor()
        cursor.execute("INSERT INTO transacoes_financeiras (cliente_id, conta_bancaria_id, data_transacao, descricao, tipo, valor, categoria, plano_conta_id) VALUES (%s, %s, %s, %s, %s, %s, %s, %s)", (session['user_id'], conta_bancaria_id, data_transacao, descricao, tipo, valor, categoria, plano_conta_id)); conn.commit(); conn.close()
        flash('Transação manual adicionada com sucesso!', 'success')
    except Exception as e: flash(f'Erro ao adicionar transação: {e}', 'danger')
    return redirect(url_for('cliente_contas'))
    
@app.route('/api/fornecedor/<int:fornecedor_id>')
def get_fornecedor_data(fornecedor_id):
    if not session.get('logged_in') or session.get('role') != 'cliente': return jsonify({'error': 'Acesso não autorizado'}), 403
    conn = get_db(); cursor = conn.cursor()
    cursor.execute("SELECT conta_padrao_id FROM fornecedores WHERE id = %s AND user_id = %s", (fornecedor_id, session['user_id'])); fornecedor = cursor.fetchone(); conn.close()
    if fornecedor: return jsonify(dict(fornecedor))
    return jsonify({'error': 'Fornecedor não encontrado'}), 404
    
@app.route('/api/meu_cliente/<int:cliente_id>')
def get_meu_cliente_data(cliente_id):
    if not session.get('logged_in') or session.get('role') != 'cliente': return jsonify({'error': 'Acesso não autorizado'}), 403
    conn = get_db(); cursor = conn.cursor()
    cursor.execute("SELECT conta_padrao_id FROM meus_clientes WHERE id = %s AND user_id = %s", (cliente_id, session['user_id'])); cliente = cursor.fetchone(); conn.close()
    if cliente: return jsonify(dict(cliente))
    return jsonify({'error': 'Cliente não encontrado'}), 404

@app.route('/admin/importar_plano_csv', methods=['GET', 'POST'])
def admin_importar_plano_csv():
    if not session.get('logged_in') or session.get('role') != 'admin':
        return redirect(url_for('login_page'))

    conn = get_db()
    cursor = conn.cursor()

    if request.method == 'POST':
        cliente_id = request.form.get('cliente_id')
        file = request.files.get('arquivo_csv')

        if not cliente_id or not file:
            flash('Selecione o cliente e o arquivo CSV.', 'warning')
            return redirect(request.url)

        try:
            df = pd.read_csv(file, sep=None, engine='python', encoding='utf-8-sig')
            df.columns = [c.strip().lower() for c in df.columns]

            col_codigo = next((c for c in df.columns if 'cod' in c), None)
            col_nome = next((c for c in df.columns if 'nom' in c), None)

            if not col_codigo or not col_nome:
                flash(f'Erro: Colunas não encontradas. O sistema leu: {list(df.columns)}', 'danger')
                return redirect(request.url)

            contas_importadas = 0
            for _, row in df.iterrows():
                codigo = str(row[col_codigo]).strip()
                nome = str(row[col_nome]).strip()
                tipo_conta = get_tipo_by_codigo(codigo, cliente_id)
                
                cursor.execute("""
                    INSERT INTO plano_contas (user_id, codigo, nome, tipo, aceita_lancamentos)
                    VALUES (%s, %s, %s, %s, TRUE)
                    ON CONFLICT (user_id, codigo) DO UPDATE SET nome = EXCLUDED.nome
                """, (cliente_id, codigo, nome, tipo_conta))
                contas_importadas += 1

            conn.commit()
            flash(f'Sucesso! {contas_importadas} contas processadas.', 'success')
            return redirect(url_for('admin_plano_de_contas', cliente_id=cliente_id))

        except Exception as e:
            conn.rollback()
            flash(f'Erro na leitura do arquivo: {e}', 'danger')
        finally:
            conn.close()

    cursor.execute("SELECT id, username FROM users WHERE role = 'cliente' ORDER BY username")
    clientes = cursor.fetchall()
    conn.close()
    return render_template('admin_importar_plano.html', clientes=clientes)

@app.route('/admin/limpar_travamento_ofx')
def limpar_travamento_ofx():
    if not session.get('logged_in') or session.get('role') != 'admin':
        return "Acesso negado", 403

    conn = get_db()
    cursor = conn.cursor()
    
    try:
        cursor.execute("DELETE FROM transacoes_pendentes WHERE user_id IN (SELECT id FROM users WHERE role = 'cliente')")
        pendentes_removidos = cursor.rowcount
        conn.commit()
        return f"Limpeza concluída! {pendentes_removidos} transações pendentes foram removidas. Tente importar o OFX novamente."
        
    except Exception as e:
        conn.rollback()
        return f"Erro ao limpar: {e}"
    finally:
        conn.close()

@app.route('/api/cliente/plano_contas/adicionar', methods=['POST'])
def api_cliente_adicionar_plano_conta():
    if not session.get('logged_in'):
        return jsonify({'success': False, 'error': 'Não autorizado'}), 403

    data = request.json
    codigo = data.get('codigo')
    nome = data.get('nome')
    user_id = data.get('user_id')

    if not all([codigo, nome, user_id]):
        return jsonify({'success': False, 'error': 'Código, nome e ID do cliente são obrigatórios.'}), 400
    
    if '.' not in codigo:
         return jsonify({'success': False, 'error': 'Código inválido. Contas criadas rapidamente devem ser sub-contas (ex: 01.10).'}), 400

    tipo_conta = get_tipo_by_codigo(codigo, user_id)
    if tipo_conta == 'Outros':
        return jsonify({'success': False, 'error': f'O código "{codigo}" não pertence a um grupo principal válido.'}), 400
    
    aceita_lancamentos = True

    try:
        conn = get_db()
        cursor = conn.cursor()
        
        cursor.execute("SELECT id FROM plano_contas WHERE codigo = %s AND user_id = %s", (codigo, user_id))
        if cursor.fetchone():
            conn.close()
            return jsonify({'success': False, 'error': 'Este código de conta já está em uso.'}), 409

        cursor.execute(
            "INSERT INTO plano_contas (user_id, codigo, nome, tipo, aceita_lancamentos) VALUES (%s, %s, %s, %s, %s) RETURNING id",
            (user_id, codigo, nome, tipo_conta, aceita_lancamentos)
        )
        new_id = cursor.fetchone()['id']
        conn.commit()
        conn.close()

        return jsonify({
            'success': True, 
            'message': 'Conta adicionada com sucesso!',
            'nova_conta': { 'id': new_id, 'codigo': codigo, 'nome': nome }
        })

    except Exception as e:
        print(f"ERRO na API /api/cliente/plano_contas/adicionar: {e}")
        traceback.print_exc()
        return jsonify({'success': False, 'error': 'Ocorreu um erro interno no servidor.'}), 500

with app.app_context():
    create_tables()

if __name__ == '__main__':
    app.run(debug=True)
    #app.run()