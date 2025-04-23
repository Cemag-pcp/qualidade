import pandas as pd
from models import Recurso
from extensions import db
from main import create_app
from models import Item

app = create_app()  # sua função que retorna o Flask app

# with app.app_context():

#     # DataFrame carregado de algum lugar:
#     df = pd.read_csv('dados_checklist.csv', sep=',')

#     df['Recurso'] = df['Recurso'].apply(lambda x: f"0{str(x)}" if len(str(x)) == 5 else str(x))
#     df.dropna(inplace=True, subset=['qnt'])

#     df['qnt'] = df['qnt'].apply(lambda x: float(str(x).replace(',', '.')))

#     # mesma lógica de antes:
#     df = df.rename(columns={
#         "REF PRINCIPAL": "carreta",
#         "Código": "codigo_carreta",
#         "Produto": "desc_carreta",
#         "Recurso": "recurso",
#         "Descrição": "descricao_recurso",
#         "qnt": "qt",
#     })
#     df = df[["carreta", "codigo_carreta", "desc_carreta", "recurso", "descricao_recurso", "qt"]]

#     recursos = [
#         Recurso(
#             carreta=row["carreta"],
#             desc_carreta=row["desc_carreta"],
#             recurso=row["recurso"],
#             descricao_recurso=row["descricao_recurso"],
#             qt=int(row["qt"])
#         )
#         for _, row in df.iterrows()
#     ]

#     db.session.bulk_save_objects(recursos)
#     db.session.commit()
#     print("Registros inseridos com sucesso!")


