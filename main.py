from slack_bolt import App
from slack_bolt.adapter.socket_mode import SocketModeHandler
from slack_sdk import WebClient
import os
from datetime import datetime

from create_tables import *
from services import (
    montar_blocos_modal_servicos,
    criar_ordem_servico_servicos,
    formatar_mensagem_chamado_servicos,
    listar_chamados_por_usuario_servicos,
    montar_blocos_exportacao_servicos,
    exportar_chamados_servicos
)

# Inicializa o app com tokens do Railway (env vars)
app = App(token=os.getenv("SLACK_BOT_TOKEN"))
client = WebClient(token=os.getenv("SLACK_BOT_TOKEN"))

# 📌 Abrir modal de Serviços
@app.command("/chamado-servicos")
def handle_chamado_servicos_command(ack, body, client, logger):
    ack()

    try:
        blocks = montar_blocos_modal_servicos()

        client.views_open(
            trigger_id=body["trigger_id"],
            view={
                "type": "modal",
                "callback_id": "modal_abertura_chamado_servicos",
                "title": {"type": "plain_text", "text": "Chamado - Serviços"},
                "submit": {"type": "plain_text", "text": "Abrir"},
                "blocks": blocks
            }
        )

    except Exception as e:
        logger.error(f"❌ Erro ao abrir modal serviços: {e}")
        client.chat_postEphemeral(
            channel=body.get("channel_id"),
            user=body["user_id"],
            text="❌ Ocorreu um erro ao abrir o formulário. Tente novamente."
        )

# 📌 Submissão do modal
@app.view("modal_abertura_chamado_servicos")
def handle_modal_submission_servicos(ack, body, view, client):
    ack()
    user = body["user"]["id"]
    canal_id = os.getenv("SLACK_CANAL_ID_SERVICOS", "CXXXXXXX")  # ⚠️ Trocar pelo canal real
    valores = view["state"]["values"]

    def pegar_valor(campo):
        bloco = valores.get(campo, {})
        if not bloco:
            return ""
        item = list(bloco.values())[0]
        return item.get("selected_option", {}).get("value") or item.get("value") or ""

    data = {
        "tipo_ticket": pegar_valor("tipo_ticket"),
        "locatario": pegar_valor("locatario"),
        "empreendimento_unidade": pegar_valor("empreendimento_unidade"),
        "responsavel": pegar_valor("responsavel"),
        "solicitante": user
    }

    response = client.chat_postMessage(
        channel=canal_id,
        text=f"🧾 ({data['locatario']}) - {data['empreendimento_unidade']} <@{user}>: *{data['tipo_ticket']}*",
        blocks=[
            {
                "type": "section",
                "text": {
                    "type": "mrkdwn",
                    "text": f"🧾 (*{data['locatario']}*) - {data['empreendimento_unidade']} <@{user}>: *{data['tipo_ticket']}*"
                }
            },
            {
                "type": "actions",
                "elements": [
                    {"type": "button", "text": {"type": "plain_text", "text": "🔄 Capturar"}, "action_id": "capturar_chamado"},
                    {"type": "button", "text": {"type": "plain_text", "text": "✅ Encerrar"}, "action_id": "finalizar_chamado"}
                ]
            }
        ]
    )

    thread_ts = response["ts"]
    criar_ordem_servico_servicos(data, thread_ts, canal_id)

    client.chat_postMessage(
        channel=canal_id,
        thread_ts=thread_ts,
        text=formatar_mensagem_chamado_servicos(data, user)
    )

@app.action("capturar_chamado")
def handle_capturar(ack, body, client):
    ack()
    user = body["user"]["id"]
    client.chat_postEphemeral(
        channel=body["channel"]["id"],
        user=user,
        text="🔄 Função de *capturar chamado* ainda não implementada."
    )

@app.action("finalizar_chamado")
def handle_finalizar(ack, body, client):
    ack()
    user = body["user"]["id"]
    client.chat_postEphemeral(
        channel=body["channel"]["id"],
        user=user,
        text="✅ Função de *finalizar chamado* ainda não implementada."
    )

# 📌 Listar chamados do usuário
@app.command("/minhas-os-servicos")
def handle_meus_chamados_servicos(ack, body, client):
    ack()
    user_id = body["user_id"]

    chamados = listar_chamados_por_usuario_servicos(user_id)

    response = client.conversations_open(users=user_id)
    channel_id = response["channel"]["id"]

    if not chamados:
        client.chat_postEphemeral(
            channel=channel_id,
            user=user_id,
            text="❗ Você ainda não abriu nenhum chamado de serviços."
        )
        return

    mensagem = "*📄 Seus chamados de serviços:*\n\n"
    for ch in chamados:
        mensagem += f"• *{ch.tipo_ticket}* - {ch.empreendimento_unidade} (Status: `{ch.status}`)\n"

    client.chat_postEphemeral(
        channel=channel_id,
        user=user_id,
        text=mensagem
    )

# 📌 Exportar chamados
@app.command("/exportar-os-servicos")
def handle_exportar_command_servicos(ack, body, client, logger):
    ack()

    try:
        blocks = montar_blocos_exportacao_servicos()

        client.views_open(
            trigger_id=body["trigger_id"],
            view={
                "type": "modal",
                "callback_id": "escolher_exportacao_servicos",
                "title": {"type": "plain_text", "text": "Exportar OS - Serviços"},
                "submit": {"type": "plain_text", "text": "Exportar"},
                "private_metadata": body["user_id"],
                "blocks": blocks
            }
        )
    except Exception as e:
        logger.error(f"❌ Erro ao abrir modal de exportação serviços: {e}")
        client.chat_postEphemeral(
            channel=body.get("channel_id"),
            user=body["user_id"],
            text="❌ Erro ao abrir modal de exportação. Tente novamente."
        )

@app.view("escolher_exportacao_servicos")
def exportar_chamados_servicos_handler(ack, body, view, client):
    ack()
    user_id = body["user"]["id"]
    valores = view["state"]["values"]

    tipo = valores["tipo_arquivo"]["value"]["selected_option"]["value"]
    data_inicio = valores["data_inicio"]["value"]["selected_date"]
    data_fim = valores["data_fim"]["value"]["selected_date"]

    data_inicio = datetime.strptime(data_inicio, "%Y-%m-%d") if data_inicio else None
    data_fim = datetime.strptime(data_fim, "%Y-%m-%d") if data_fim else None

    arquivos = exportar_chamados_servicos(data_inicio, data_fim, tipo=tipo)

    response = client.conversations_open(users=user_id)
    channel_id = response["channel"]["id"]

    client.files_upload_v2(
        channel=channel_id,
        initial_comment="📎 Exportação de chamados de serviços",
        file=arquivos[tipo][1],
        filename=arquivos[tipo][0],
        title=f"chamados-servicos.{tipo}"
    )

# 🚀 Inicializar o app
if __name__ == "__main__":
    SocketModeHandler(app, os.getenv("SLACK_APP_TOKEN")).start()
