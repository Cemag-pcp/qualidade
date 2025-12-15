from flask import Flask, jsonify, request, render_template, Blueprint
from flask_cors import CORS
from models import db, Item, Recurso           
from flask import send_file
from sqlalchemy import or_, and_

from io import BytesIO
import gspread
from google.oauth2 import service_account
import pandas as pd
from datetime import datetime
import re
import numpy as np
import pdfkit
import zipfile
import tempfile
from concurrent.futures import ProcessPoolExecutor, as_completed

checklist = Blueprint('checklist', __name__)
planilha_cargas_cache = {}
config = pdfkit.configuration(wkhtmltopdf=r'C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe')

def extrair_cor(texto):
    """
    Procura a primeira sigla de cor em 'texto' e devolve a cor correspondente.
    Se não encontrar, devolve 'Laranja'.
    """

    df_cores = pd.DataFrame({'Recurso_cor':['AN','VJ','LC','VM','AV','sem_cor'], 
        'cor':['Azul','Verde','Laranja','Vermelho','Amarelo','Laranja']})
        
    map_cor = dict(zip(df_cores['Recurso_cor'], df_cores['cor']))
    padrao = re.compile(r'(' + '|'.join(map_cor.keys()) + r')', flags=re.I)

    m = padrao.search(str(texto))
    return map_cor.get(m.group(1).upper(), 'Laranja') if m else 'Laranja'

def remove_sigla_cor(texto: str) -> str:
    """Remove AN, VJ, LC, VM, AV (case‑insensitive) de 'texto'."""

    SIGLAS_COR = ['AN', 'VJ', 'LC', 'VM', 'AV']
    PADRAO_SIGLA = re.compile(r'\b(?:' + '|'.join(SIGLAS_COR) + r')\b', flags=re.I)
    sem_sigla = PADRAO_SIGLA.sub('', str(texto))
    # elimina espaços duplos que sobrarem e tira espaços nas pontas
    return re.sub(r'\s{2,}', ' ', sem_sigla).strip()

def planilha_cargas(data_carga, reset_cache=False):
    global planilha_cargas_cache

    if not reset_cache and data_carga in planilha_cargas_cache:
        return planilha_cargas_cache[data_carga]

    # Conversão e carregamento dos dados
    data_carga_formatada = datetime.strptime(data_carga, "%Y-%m-%d").strftime("%d/%m/%Y")

    scope = ['https://www.googleapis.com/auth/spreadsheets', "https://www.googleapis.com/auth/drive"]
    credentials = service_account.Credentials.from_service_account_file('credentials.json', scopes=scope)

    name_sheet = 'RQ AV-002-000 (PLANILHA DE CARGAS)'
    sa = gspread.authorize(credentials)
    sh = sa.open(name_sheet)
    wks1 = sh.worksheet('Importar Dados')

    list1 = wks1.get()
    importar_dados = pd.DataFrame(list1)
    importar_dados = importar_dados[[1,3,6,9,10,11,13]]
    importar_dados.columns = ['Estado','Cidade','Datas','Pessoa','Recurso','Descrição','Serie']
    importar_dados = importar_dados[1:]  # remove header
    importar_dados['Cidade/Estado'] = importar_dados['Cidade'] + '/' + importar_dados['Estado']
    importar_dados.drop(columns=['Estado', 'Cidade'], inplace=True)

    importar_dados = importar_dados[importar_dados['Datas'].notnull() & (importar_dados['Datas'] != '')]
    importar_dados['Datas'] = pd.to_datetime(importar_dados['Datas'], format="%d/%m/%Y", errors='coerce')
    importar_dados['Datas'] = importar_dados['Datas'].dt.strftime("%d/%m/%Y")

    dados_filtrados = importar_dados[importar_dados['Datas'] == data_carga_formatada]
    dados_filtrados['cor'] = dados_filtrados['Recurso'].apply(extrair_cor)

    # salva no cache
    planilha_cargas_cache[data_carga] = dados_filtrados

    return dados_filtrados

def create_app():
    app = Flask(__name__)
    app.config["SQLALCHEMY_DATABASE_URI"] = "sqlite:///items.db"
    app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False

    db.init_app(app)               # agora a instância sabe qual app usar
    CORS(app)

    # Criar tabelas
    with app.app_context():
        db.create_all() 

    app.register_blueprint(checklist)
                           
    return app

@checklist.route("/planilha-cargas", methods=["GET"])
def rota_planilha_cargas():
    data_carga = request.args.get("data")
    reset = request.args.get("reset", "false").lower() == "true"

    if not data_carga:
        return jsonify({"error": "Parâmetro 'data' é obrigatório."}), 400

    try:
        df = planilha_cargas(data_carga, reset_cache=reset)
        return df.to_json(orient="records", force_ascii=False)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@checklist.route("/api/items")
def listar_items():
    query = Item.query

    # lista dos campos que você permite filtrar
    for campo in ("carreta", "codigo", "tipo", "item_code"):
        if valor := request.args.get(campo):
            # filtra por LIKE case‑insensitive → "%valor%"
            query = query.filter(getattr(Item, campo).ilike(f"%{valor}%"))

    itens = query.order_by(Item.id).all()
    return jsonify([i.as_dict() for i in itens])

@checklist.route("/api/recursos")
def listar_recursos():
    """
    Retorna os registros da tabela Recurso permitindo filtros basicos.
    """
    query = Recurso.query

    for campo in ("carreta", "recurso"):
        if valor := request.args.get(campo):
            query = query.filter(getattr(Recurso, campo).ilike(f"%{valor}%"))

    recursos = query.order_by(Recurso.id).all()
    return jsonify([r.as_dict() for r in recursos])

@checklist.route("/api/recurso", methods=["POST"])
def adicionar_recurso():
    data = request.get_json()

    required = ("carreta", "desc_carreta", "recurso", "descricao_recurso", "qt")
    if not data or any(not str(data.get(campo, "")).strip() for campo in required):
        return jsonify({"error": "Todos os campos sao obrigatorios."}), 400

    existente = (
        Recurso.query
        .filter(and_(Recurso.carreta == data["carreta"], Recurso.recurso == data["recurso"]))
        .first()
    )
    if existente:
        return jsonify({"error": "Recurso ja cadastrado para essa carreta."}), 400

    recurso = Recurso(
        carreta=str(data["carreta"]).strip(),
        desc_carreta=str(data["desc_carreta"]).strip(),
        recurso=str(data["recurso"]).strip(),
        descricao_recurso=str(data["descricao_recurso"]).strip(),
        qt=float(data["qt"])
    )

    db.session.add(recurso)
    db.session.commit()

    return jsonify(recurso.as_dict()), 201

@checklist.route("/api/recurso/<int:recurso_id>", methods=["PUT"])
def editar_recurso(recurso_id):
    data = request.get_json()

    required = ("carreta", "desc_carreta", "recurso", "descricao_recurso", "qt")
    if not data or any(not str(data.get(campo, "")).strip() for campo in required):
        return jsonify({"error": "Todos os campos sao obrigatorios."}), 400

    recurso = Recurso.query.get_or_404(recurso_id)

    duplicado = (
        Recurso.query
        .filter(
            Recurso.carreta == data["carreta"],
            Recurso.recurso == data["recurso"],
            Recurso.id != recurso_id,
        )
        .first()
    )
    if duplicado:
        return jsonify({"error": "Ja existe outro recurso com essa combinacao."}), 400

    recurso.carreta = str(data["carreta"]).strip()
    recurso.desc_carreta = str(data["desc_carreta"]).strip()
    recurso.recurso = str(data["recurso"]).strip()
    recurso.descricao_recurso = str(data["descricao_recurso"]).strip()
    recurso.qt = float(data["qt"])

    db.session.commit()
    return jsonify(recurso.as_dict())

@checklist.route("/api/recurso/<int:recurso_id>", methods=["DELETE"])
def deletar_recurso(recurso_id):
    recurso = Recurso.query.get_or_404(recurso_id)
    db.session.delete(recurso)
    db.session.commit()
    return jsonify({"success": True})

@checklist.route("/api/items/<int:item_id>")
def item_unico(item_id):
    return jsonify(Item.query.get_or_404(item_id).as_dict())

@checklist.route("/")
def index():

    return render_template("index.html")

@checklist.route("/recursos")
def recursos_page():
    """
    Tela simples para visualizar recursos cadastrados.
    """
    return render_template("recursos.html")

def get_dados_por_recurso(df_rc, model):
    resultado = []

    for _, linha in df_rc.iterrows():
        recurso_nome = remove_sigla_cor(linha['Recurso'])
        cor = linha['cor']

        itens = (model.query
                      .filter_by(carreta=recurso_nome)
                      .all())

        resultado.append({
            "recurso": recurso_nome,
            "cor": cor,
            "itens": [i.as_dict() for i in itens]
        })

    return resultado

@checklist.route("/modelo")
def gerar_modelo():
    data_carga = request.args.get("dataCarga")
    recurso     = request.args.get("recurso").rstrip()
    serie     = request.args.get("serie")

    df = planilha_cargas(data_carga)
    df = df[df['Serie'] == serie]

    # filtra só pela carreta pedida
    df_rc = (
        df[['Recurso', 'cor']]
        .drop_duplicates()
        .query("Recurso == @recurso")
        .reset_index(drop=True)
    )

    acessorios = get_dados_por_recurso(df_rc, Item)
    recursos   = get_dados_por_recurso(df_rc, Recurso)
    
    acessorios_extras = [{"codigo":"Acessórios 01", "descricao":"Acessórios", "qt":1}]
    if 'bb' in recurso.lower():
        acessorios_extras.append({
            "codigo": "200391",
            "descricao": "BOMBA",
            "qt": 1
        })

    if 'rs/rd' in recurso.lower():
        acessorios_extras.append({
            "codigo": "214108",
            "descricao": "RODA 6 FUROS TANDEM FA6 Flag Romaneio",
            "qt": 2
        })

    if recurso.lower().find('rs/rd') != -1 and recurso.lower().find('roda 20') != -1: 
        
        # Se tiver rs/rd dentro do código e aro 20
        acessorios_extras.append({
            "codigo": "035390",
            "descricao": "RODA RS R20 CINZA [CBH12/FT12500] Flag Romaneio",
            "qt": 2
        })

    if recurso.lower().find('rs/rd') != -1 and (recurso.lower().find('(i)') != -1 or recurso.lower().find('(r)') != -1):
        
        # Se tiver rs/rd dentro do código
        acessorios_extras.append({
            "codigo": "214108",
            "descricao": "RODA 6 FUROS TANDEM FA6 Flag Romaneio",
            "qt": 2
        })
        # Se tiver pneu dentro do código
        acessorios_extras.append({
            "codigo": "Pneus",
            "descricao": "PNEUS",
            "qt": 6
        })

    if not recurso.lower().find('rs/rd') != -1 and (recurso.lower().find('(i)') != -1 or recurso.lower().find('(r)') != -1):

        # Se tiver pneu dentro do código
        acessorios_extras.append({
            "codigo": "Pneus",
            "descricao": "PNEUS",
            "qt": 4
        })

    #Tirante
    if not recurso.lower().find('roçax') != -1 or not recurso.lower().find('f6') != -1 or not recurso.lower().find('f4') != -1 or not recurso.lower().find('fa5') != -1 or not recurso.lower().find('ftc4300') != -1 or not recurso.lower().find('ftc6500') != -1 or not recurso.lower().find('ftc10500') != -1: 
        
        acessorios_extras.append({
            "codigo": "TIRANTE DA TRAVA COMPLETO",
            "descricao": "TIRANTE DA TRAVA COMPLETO",
            "qt": 2
        })

    context = {
        "n_serie": df['Serie'].dropna().drop_duplicates().tolist()[0],
        "desc_carreta": df['Descrição'].dropna().drop_duplicates().tolist()[0],
        "carreta": recurso,
        "cor": df_rc['cor'].dropna().drop_duplicates().tolist()[0],
        "cliente": df['Pessoa'].dropna().drop_duplicates().tolist()[0],
        "data_carga": datetime.strptime(data_carga, "%Y-%m-%d").strftime("%d/%m/%Y"),
        "cidade_estado": df['Cidade/Estado'].dropna().drop_duplicates().tolist()[0],

        "acessorios": acessorios,
        "recursos": recursos,
        "acessorios_extras": acessorios_extras,
    }

    return render_template("modelo.html", **context)

@checklist.route("/api/carretas")
def listar_carretas():
    data_carga = request.args.get("dataCarga")
    df = planilha_cargas(data_carga)
    carretas = df[['Recurso', 'Serie']].dropna(subset=['Recurso', 'Serie']).drop_duplicates().values.tolist()
    return jsonify(carretas=carretas)

@checklist.route("/api/listar-carretas")
def carretas_existentes():
    """
    Retorna sugestões de carretas conforme o texto digitado.
    """
    termo = request.args.get("carreta", "").strip()

    if not termo:
        return jsonify(carretas=[])

    carretas = (
        db.session.query(Item.carreta)
        .filter(Item.carreta.ilike(f"%{termo}%"))
        .distinct()
        .limit(10)
        .all()
    )

    # carretas será uma lista de tuplas: [('CBHM1000',), ('CBH2000',)...]
    return jsonify(carretas=[c[0] for c in carretas])

@checklist.route("/api/descricao-carreta")
def descricao_carreta():
    """
    Retorna a descrição da carreta selecionada.
    """
    carreta = request.args.get("carreta", "").strip()
    print(carreta)
    if not carreta:
        return jsonify(descricao=None)

    # Busca a primeira descrição relacionada à carreta
    resultado = (
        db.session.query(Item.desc)
        .filter(Item.carreta == carreta)
        .first()
    )

    descricao = resultado[0] if resultado else None

    print(descricao)

    return jsonify({"descricao": descricao})

@checklist.route("/api/codigo-carreta")
def codigo_carreta():
    """
    Retorna o codigo da carreta selecionada.
    """

    carreta = request.args.get("carreta", "").strip()

    if not carreta:
        return jsonify(codigo=None)

    # Busca a primeira descrição relacionada à carreta
    resultado = (
        db.session.query(Item.codigo)
        .filter(Item.carreta == carreta)
        .first()
    )

    codigo = resultado[0] if resultado else None

    return jsonify({"codigo": codigo})

@checklist.route("/api/item", methods=["POST"])
def adicionar_item():
    data = request.get_json()

    # Verifica se a combinação já existe
    existente = Item.query.filter_by(carreta=data["carreta"], item_code=data["itemCode"]).first()
    if existente:
        return jsonify({"error": "Item já cadastrado para essa carreta"}), 400

    if not data['carreta'] or not data['codigo'] or not data['desc'] or not data['itemCode'] or not data['qt'] or not data['itemDescription'] or not data['tipo']:
        return jsonify({"error": "Todos os campos são obrigatórios!"}), 400

    item = Item(
        carreta=data["carreta"],
        codigo=data["codigo"],
        desc=data["desc"],
        item_code=data["itemCode"],
        qt=data["qt"],
        item_description=data["itemDescription"],
        tipo=data["tipo"]
    )

    db.session.add(item)
    db.session.commit()

    return jsonify(item.as_dict()), 201

@checklist.route("/api/item/<int:item_id>", methods=["PUT"])
def editar_item(item_id):
    data = request.get_json()

    if not data['carreta'] or not data['codigo'] or not data['desc'] or not data['itemCode'] or not data['qt'] or not data['itemDescription'] or not data['tipo']:
        return jsonify({"error": "Todos os campos são obrigatórios!"}), 400

    item = Item.query.get_or_404(item_id)

    # Verifica se há outro item com mesma carreta + item_code
    duplicado = (
        Item.query
        .filter(Item.carreta == data["carreta"], Item.item_code == data["itemCode"], Item.id != item_id)
        .first()
    )
    
    if duplicado:
        return jsonify({"error": "Já existe outro item com essa combinação"}), 400

    item.carreta = data["carreta"]
    item.codigo = data["codigo"]
    item.desc = data["desc"]
    item.item_code = data["itemCode"]
    item.qt = data["qt"]
    item.item_description = data["itemDescription"]
    item.tipo = data["tipo"]

    db.session.commit()
    return jsonify(item.as_dict())

@checklist.route("/api/item/<int:item_id>", methods=["DELETE"])
def deletar_item(item_id):
    item = Item.query.get_or_404(item_id)
    db.session.delete(item)
    db.session.commit()
    return jsonify({"success": True})

@checklist.route("/api/planilha-modelo")
def baixar_modelo_planilha():
    # Exemplo de estrutura de planilha
    df = pd.DataFrame([
        {"carreta": "CBHM4500 SS RD P750(I) M17", "desc_carreta": "614233M17 - CBHM 5000 RD SC M17 - LARANJA - (CARRETA BASCULANTE HIDRAULICA 5T) (Un)", "recurso": "teste", "descricao_recurso": "teste", "qt": 2},
    ])

    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name="Modelo")

    output.seek(0)
    return send_file(output,
                     download_name="modelo_planilha.xlsx",
                     as_attachment=True,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

@checklist.route("/api/importar-recurso", methods=["POST"])
def importar_recurso():
    file = request.files.get("arquivo")

    if not file or not file.filename.endswith(".xlsx"):
        return jsonify({"error": "Arquivo inválido. Envie um .xlsx"}), 400

    try:
        df = pd.read_excel(file)

        # Verifica colunas obrigatórias
        required_cols = {"carreta", "desc_carreta", "recurso", "descricao_recurso", "qt"}
        if not required_cols.issubset(df.columns):
            return jsonify({"error": f"Colunas obrigatórias: {', '.join(required_cols)}"}), 400

        inseridos = []
        duplicados = []

        for _, row in df.iterrows():
            carreta = str(row["carreta"]).strip()
            recurso = str(row["recurso"]).strip()

            # Verifica duplicação
            ja_existe = Recurso.query.filter(and_(
                Recurso.carreta == carreta,
                Recurso.recurso == recurso
            )).first()

            if ja_existe:
                duplicados.append({"carreta": carreta, "recurso": recurso})
                continue

            novo = Recurso(
                carreta=carreta,
                desc_carreta=str(row["desc_carreta"]),
                recurso=recurso,
                descricao_recurso=str(row["descricao_recurso"]),
                qt=float(row["qt"])
            )

            db.session.add(novo)
            inseridos.append(novo)

        db.session.commit()

        return jsonify({
            "inseridos": len(inseridos),
            "duplicados": duplicados
        })

    except Exception as e:
        return jsonify({"error": f"Erro ao processar arquivo: {str(e)}"}), 500

def gerar_pdf_da_carreta(args):
    data_carga, recurso, serie = args
    modelo_url = f"http://192.168.3.3:8083/modelo?dataCarga={data_carga}&recurso={recurso}&serie={serie}"

    options = {
        'page-size': 'A4',
        'margin-top': '0.75in',
        'margin-right': '0.75in',
        'margin-bottom': '0.75in',
        'margin-left': '0.75in',
        'encoding': "UTF-8",
        'no-outline': None,
        'enable-local-file-access': None,
        'javascript-delay': 2000,
        'load-error-handling': 'ignore',
        'load-media-error-handling': 'ignore'
    }
    
    try:
        pdf_bytes = pdfkit.from_url(modelo_url, False, options=options, configuration=config)
        filename = f'checklist_{recurso}_{serie}.pdf'
        return (filename, pdf_bytes)
    except Exception as e:
        print(f"Erro ao gerar PDF para {recurso}-{serie}: {str(e)}")
        return None

@checklist.route('/api/gerar-pdfs-zip', methods=['POST'])
def gerar_pdfs_zip():
    try:
        data = request.get_json()
        data_carga = data.get('dataCarga')
        carretas = data.get('carretas', [])
        
        if not data_carga or not carretas:
            return jsonify({'error': 'Parâmetros obrigatórios não fornecidos'}), 400
        
        zip_buffer = BytesIO()
        
        # Preparar args para cada carreta
        args_list = [(data_carga, carreta[0], carreta[1]) for carreta in carretas]
        
        with ProcessPoolExecutor() as executor:
            futures = [executor.submit(gerar_pdf_da_carreta, args) for args in args_list]
            
            with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                for future in as_completed(futures):
                    result = future.result()
                    if result:
                        filename, pdf_bytes = result
                        zip_file.writestr(filename, pdf_bytes)
        
        zip_buffer.seek(0)
        filename = f'checklists_{data_carga}.zip'
        
        return send_file(
            zip_buffer,
            as_attachment=True,
            download_name=filename,
            mimetype='application/zip'
        )
        
    except Exception as e:
        print(f"Erro ao gerar ZIP dos PDFs: {str(e)}")
        return jsonify({'error': f'Erro interno do servidor: {str(e)}'}), 500
    
if __name__ == "__main__":
    app = create_app()
    app.run(debug=True, host="0.0.0.0", port=8083)