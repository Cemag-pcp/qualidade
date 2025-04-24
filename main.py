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

checklist = Blueprint('checklist', __name__)

def extrair_cor(texto):
    """
    Procura a primeira sigla de cor em 'texto' e devolve a cor correspondente.
    Se não encontrar, devolve 'sem_cor'.
    """

    df_cores = pd.DataFrame({'Recurso_cor':['AN','VJ','LC','VM','AV','sem_cor'], 
        'cor':['Azul','Verde','Laranja','Vermelho','Amarelo','Laranja']})
        
    map_cor = dict(zip(df_cores['Recurso_cor'], df_cores['cor']))
    padrao = re.compile(r'(' + '|'.join(map_cor.keys()) + r')', flags=re.I)

    m = padrao.search(str(texto))
    return map_cor.get(m.group(1).upper(), 'sem_cor') if m else 'sem_cor'

def remove_sigla_cor(texto: str) -> str:
    """Remove AN, VJ, LC, VM, AV (case‑insensitive) de 'texto'."""

    SIGLAS_COR = ['AN', 'VJ', 'LC', 'VM', 'AV']
    PADRAO_SIGLA = re.compile(r'\b(?:' + '|'.join(SIGLAS_COR) + r')\b', flags=re.I)
    sem_sigla = PADRAO_SIGLA.sub('', str(texto))
    # elimina espaços duplos que sobrarem e tira espaços nas pontas
    return re.sub(r'\s{2,}', ' ', sem_sigla).strip()

def planilha_cargas(data_carga):
    
    data_carga = datetime.strptime(data_carga, "%Y-%m-%d").strftime("%d/%m/%Y")

    scope = ['https://www.googleapis.com/auth/spreadsheets',
         "https://www.googleapis.com/auth/drive"]

    credentials = service_account.Credentials.from_service_account_file('credentials.json', scopes=scope)

    name_sheet = 'RQ AV-002-000 (PLANILHA DE CARGAS)'

    sa = gspread.authorize(credentials)
    sh = sa.open(name_sheet)

    worksheet1 = 'Importar Dados'
    wks1 = sh.worksheet(worksheet1)

    list1 = wks1.get()
    importar_dados = pd.DataFrame(list1)
    importar_dados = importar_dados[[1,3,6,9,10,11,13]]
    importar_dados = importar_dados.rename(columns={1:'Estado',3:'Cidade',6: 'Datas',9:'Pessoa',10:'Recurso',11:'Descrição',13:'Serie'})
    importar_dados['Cidade/Estado'] = importar_dados['Cidade'] + '/' + importar_dados['Estado']
    importar_dados.drop(columns=['Estado', 'Cidade'], inplace=True)
    importar_dados = importar_dados[1:]
    importar_dados = importar_dados[importar_dados['Datas'].notnull() & (importar_dados['Datas'] != '')]

    importar_dados['Datas'] = pd.to_datetime(importar_dados['Datas'], format="%d/%m/%Y", errors='coerce')
    importar_dados['Datas'] = importar_dados['Datas'].dt.strftime("%d/%m/%Y")
    
    dados_filtrados = importar_dados[importar_dados['Datas'] == data_carga]

    dados_filtrados['cor'] = dados_filtrados['Recurso'].apply(extrair_cor)

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

@checklist.route("/api/items/<int:item_id>")
def item_unico(item_id):
    return jsonify(Item.query.get_or_404(item_id).as_dict())

@checklist.route("/")
def index():

    return render_template("index.html")

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
    recurso     = request.args.get("recurso")
    
    df = planilha_cargas(data_carga)
    df = df[df['Recurso'] == recurso]

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
    carretas = df['Recurso'].dropna().drop_duplicates().tolist()
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

if __name__ == "__main__":
    app = create_app()
    app.run(debug=True)
