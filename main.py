from slack_bolt import App
from slack_bolt.adapter.socket_mode import SocketModeHandler
from slack_sdk import WebClient
from dotenv import load_dotenv
import os
from datetime import datetime

from create_tables import *
from services import (
    montar_blocos_modal_financeiro,
    criar_ordem_servico_financeiro,
    formatar_mensagem_chamado_financeiro,
    capturar_chamado,
    finalizar_chamado,
    abrir_modal_reabertura,
    abrir_modal_edicao,
    abrir_modal_cancelamento,
    reabrir_chamado,
    cancelar_chamado,
    editar_chamado
)

load_dotenv()

app = App(token=os.getenv("SLACK_BOT_TOKEN"))
client = WebClient(token=os.getenv("SLACK_BOT_TOKEN"))

# 📌 Abrir modal
@app.command("/financeiro-os")
def handle_chamado_financeiro_command(ack, body, client, logger):
    ack()  # ⚡️ Ack IMEDIATO

    try:
        # 💡 Montar blocos depois do ack, pois é leve
        blocks = montar_blocos_modal_financeiro()

        client.views_open(
            trigger_id=body["trigger_id"],
            view={
                "type": "modal",
                "callback_id": "modal_abertura_chamado_financeiro",
                "title": {"type": "plain_text", "text": "Chamado Financeiro"},
                "submit": {"type": "plain_text", "text": "Abrir"},
                "blocks": blocks
            }
        )

    except Exception as e:
        logger.error(f"❌ Erro ao abrir modal financeiro: {e}")
        client.chat_postEphemeral(
            channel=body.get("channel_id", os.getenv("SLACK_CANAL_ID", "#financeiro-comercial")),
            user=body["user_id"],
            text="❌ Ocorreu um erro ao abrir o formulário. Tente novamente."
        )

# 📌 Submissão do modal
@app.view("modal_abertura_chamado_financeiro")
def handle_modal_submission(ack, body, view, client):
    ack()
    user = body["user"]["id"]
    canal_id = os.getenv("SLACK_CANAL_ID", "C08KMCDNEFR")
    valores = view["state"]["values"]

    def pegar_valor(campo):
        bloco = valores.get(campo, {})
        if not bloco:
            return ""
        item = list(bloco.values())[0]
        return item.get("selected_option", {}).get("value") or item.get("value") or ""

    data = {
        "tipo_ticket": pegar_valor("tipo_ticket"),
        "tipo_contrato": pegar_valor("tipo_contrato"),
        "locatario": pegar_valor("locatario"),
        "empreendimento_unidade": pegar_valor("empreendimento_unidade"),
        "numero_reserva": pegar_valor("numero_reserva"),
        "responsavel": pegar_valor("responsavel"),
        "solicitante": user
    }

    response = client.chat_postMessage(
        channel=canal_id,
        text=f"🧾 ({data['locatario']}) - {data['empreendimento_unidade']} <@{user}>: *{data['tipo_ticket']}*",
        blocks=[
            {
                "type": "section",
                "text": {"type": "mrkdwn", "text": f"🧾 (*{data['locatario']}*) - {data['empreendimento_unidade']} <@{user}>: *{data['tipo_ticket']}*"}
            },
            {
                "type": "actions",
                "elements": [
                    {"type": "button", "text": {"type": "plain_text", "text": "🔄 Capturar"}, "action_id": "capturar_chamado"},
                    {"type": "button", "text": {"type": "plain_text", "text": "✅ Finalizar"}, "action_id": "finalizar_chamado"},
                    {"type": "button", "text": {"type": "plain_text", "text": "♻️ Reabrir"}, "action_id": "reabrir_chamado"},
                    {"type": "button", "text": {"type": "plain_text", "text": "❌ Cancelar"}, "action_id": "cancelar_chamado"}
                    # {"type": "button", "text": {"type": "plain_text", "text": "✏️ Editar"}, "action_id": "editar_chamado"}
                ]
            }
        ]
    )

    thread_ts = response["ts"]
    criar_ordem_servico_financeiro(data, thread_ts, canal_id)

    client.chat_postMessage(
        channel=canal_id,
        thread_ts=thread_ts,
        text=formatar_mensagem_chamado_financeiro(data, user)
    )

# 🔘 AÇÕES DE BOTÕES
@app.action("capturar_chamado")
def handle_capturar(ack, body, client):
    ack()
    capturar_chamado(client, body)

@app.action("finalizar_chamado")
def handle_finalizar(ack, body, client):
    ack()
    finalizar_chamado(client, body)

@app.action("reabrir_chamado")
def handle_reabrir(ack, body, client):
    ack()
    abrir_modal_reabertura(client, body)

@app.view("reabrir_chamado_modal")
def handle_reabrir_submit(ack, body, view, client):
    ack()
    reabrir_chamado(client, body, view)

@app.action("editar_chamado")
def handle_editar(ack, body, client):
    ack()
    abrir_modal_edicao(client, body["trigger_id"], body["message"]["ts"])

@app.action("cancelar_chamado")
def handle_cancelar(ack, body, client):
    ack()
    abrir_modal_cancelamento(client, body)

@app.view("cancelar_chamado_modal")
def handle_cancelar_submit(ack, body, view, client):
    ack()
    cancelar_chamado(client, body, view)

@app.view("editar_chamado_modal")
def handle_editar_submit(ack, body, view, client):
    ack()
    editar_chamado(client, body, view)

@app.command("/minhas-os-financeiro")
def handle_meus_chamados(ack, body, client):
    ack()
    user_id = body["user_id"]

    from services import listar_chamados_por_usuario
    chamados = listar_chamados_por_usuario(user_id)

    response = client.conversations_open(users=user_id)
    channel_id = response["channel"]["id"]

    if not chamados:
        client.chat_postEphemeral(
            channel=channel_id,
            user=user_id,
            text="❗ Você ainda não abriu nenhum chamado financeiro."
        )
        return

    mensagem = "*📄 Seus chamados financeiros:*\n\n"
    for ch in chamados:
        mensagem += f"• *{ch.tipo_ticket}* - {ch.empreendimento_unidade} (Status: `{ch.status}`)\n"

    client.chat_postEphemeral(
        channel=channel_id,
        user=user_id,
        text=mensagem
    )

@app.command("/exportar-os-financeiro")
def handle_exportar_command_financeiro(ack, body, client, logger):
    ack()

    try:
        from services import montar_blocos_exportacao
        blocks = montar_blocos_exportacao()

        client.views_open(
            trigger_id=body["trigger_id"],
            view={
                "type": "modal",
                "callback_id": "escolher_exportacao_financeiro",
                "title": {"type": "plain_text", "text": "Exportar OS Financeiro"},
                "submit": {"type": "plain_text", "text": "Exportar"},
                "private_metadata": body["user_id"],
                "blocks": blocks
            }
        )
    except Exception as e:
        logger.error(f"❌ Erro ao abrir modal de exportação financeira: {e}")
        client.chat_postEphemeral(
            channel=body.get("channel_id"),
            user=body["user_id"],
            text="❌ Erro ao abrir modal de exportação. Tente novamente."
        )

@app.view("escolher_exportacao_financeiro")
def exportar_chamados_financeiro_handler(ack, body, view, client):
    ack()
    user_id = body["user"]["id"]
    valores = view["state"]["values"]

    tipo = valores["tipo_arquivo"]["value"]["selected_option"]["value"]
    data_inicio = valores["data_inicio"]["value"]["selected_date"]
    data_fim = valores["data_fim"]["value"]["selected_date"]

    from services import exportar_chamados_financeiro

    data_inicio = datetime.strptime(data_inicio, "%Y-%m-%d") if data_inicio else None
    data_fim = datetime.strptime(data_fim, "%Y-%m-%d") if data_fim else None

    arquivos = exportar_chamados_financeiro(data_inicio, data_fim, tipo=tipo)

    response = client.conversations_open(users=user_id)
    channel_id = response["channel"]["id"]

    client.files_upload_v2(
        channel=channel_id,
        initial_comment="📎 Exportação de chamados financeiros",
        file=arquivos[tipo][1],
        filename=arquivos[tipo][0],
        title=f"chamados-financeiro.{tipo}"
)

# 🚀 Inicializar o app
if __name__ == "__main__":
    SocketModeHandler(app, os.getenv("SLACK_APP_TOKEN")).start()
