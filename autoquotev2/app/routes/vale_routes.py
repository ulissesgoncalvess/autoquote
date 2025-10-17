from flask import Blueprint, jsonify, send_file, request
import os, glob, threading, time

vale_bp = Blueprint('vale', __name__)

# Caminho padrão onde o robô salva as planilhas
OUTPUT_DIR = os.path.join('output')
PLANILHA_PATH = os.path.join(OUTPUT_DIR, 'coleta_vale.xlsx')

# Variáveis de status
status_robo = {
    "executando": False,
    "mensagem": "Pronto para executar",
    "progresso": 0
}

# ============================
# ROTA: STATUS DO ROBÔ
# ============================
@vale_bp.route('/vale/status')
def vale_status():
    data = status_robo.copy()
    data["planilha_disponivel"] = os.path.exists(PLANILHA_PATH)
    return jsonify(data)

# ============================
# ROTA: EXECUTAR ROBÔ
# ============================
@vale_bp.route('/vale/executar', methods=['POST'])
def vale_executar():
    if status_robo["executando"]:
        return jsonify({"erro": "Robô já está em execução"}), 400

    data = request.get_json()
    data_coleta = data.get("data_coleta")

    thread = threading.Thread(target=rodar_robo_vale, args=(data_coleta,))
    thread.start()

    return jsonify({"mensagem": "Execução iniciada"})

# ============================
# LÓGICA SIMULADA DO ROBÔ
# ============================
def rodar_robo_vale(data_coleta):
    try:
        status_robo.update({
            "executando": True,
            "mensagem": f"Executando coleta {data_coleta}...",
            "progresso": 0
        })

        # Simulação de execução (substitua por sua função real)
        for i in range(1, 101, 10):
            time.sleep(1)
            status_robo["progresso"] = i

        # Aqui você colocaria sua chamada real, tipo:
        # robo_vale.executar(data_coleta)

        # Simula criação da planilha
        os.makedirs(OUTPUT_DIR, exist_ok=True)
        with open(PLANILHA_PATH, 'w') as f:
            f.write("Exemplo de planilha gerada.")

        status_robo.update({
            "executando": False,
            "mensagem": "Execução concluída com sucesso!",
            "progresso": 100
        })
    except Exception as e:
        status_robo.update({
            "executando": False,
            "mensagem": f"Erro: {e}",
            "progresso": 0
        })

# ============================
# ROTA: DOWNLOAD
# ============================
@vale_bp.route('/vale/download')
def vale_download():
    arquivos = glob.glob(os.path.join(OUTPUT_DIR, '*.xlsx'))
    if arquivos:
        arquivo_recente = max(arquivos, key=os.path.getctime)
        return send_file(arquivo_recente, as_attachment=True)
    else:
        return jsonify({"erro": "Nenhum arquivo encontrado"}), 404
