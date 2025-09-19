from database import SessionLocal
from models import OrdemServicoFinanceiro
from datetime import datetime
import json
import os

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter
from slack_sdk import WebClient

client_slack = WebClient(token=os.getenv("SLACK_BOT_TOKEN"))

def get_nome_slack(user_id):
    if not user_id or not user_id.startswith("U"):
        return user_id
    try:
        info = client_slack.users_info(user=user_id)
        return info["user"]["real_name"]
    except Exception as e:
        print(f"Erro ao buscar nome real de {user_id}: {e}")
        return user_id
        

# 🔧 Função para montar o modal do financeiro
def montar_blocos_modal_financeiro():
    return [
        {
            "type": "input",
            "block_id": "tipo_ticket",
            "element": {
                "type": "static_select",
                "action_id": "value",
                "placeholder": {"type": "plain_text", "text": "Escolha"},
                "options": [{"text": {"type": "plain_text", "text": opt}, "value": opt}
                            for opt in [
                                "Ajuste de Fatura",
                                "Ajuste de Reserva",
                                "Prorrogação de Data",
                                "Saída Antecipada",
                                "Prorrogação de Desconto",
                                "Encerramento de Contrato",
                                "Alteração de Apto",
                                "Incluir reserva no LH",
                                "Geração de link Manual",
                                "Histórico de Morador",
                                "Recibo",
                                "Comprovante de Pagamento",
                                "Falta de Contato Morador"
                            ]]
            },
            "label": {"type": "plain_text", "text": "Tipo de Ticket"}
        },
        {
            "type": "input",
            "block_id": "tipo_contrato",
            "element": {
                "type": "static_select",
                "action_id": "value",
                "placeholder": {"type": "plain_text", "text": "Escolha"},
                "options": [{"text": {"type": "plain_text", "text": opt}, "value": opt}
                            for opt in ["Short Stay", "Temporada", "Long Stay", "Comodato", "Cortesia"]],
            },
            "label": {"type": "plain_text", "text": "Tipo de Contrato"}
        },
        {
            "type": "input",
            "block_id": "locatario",
            "element": {"type": "plain_text_input", "action_id": "value"},
            "label": {"type": "plain_text", "text": "Locatário"}
        },
        {
            "type": "input",
            "block_id": "empreendimento_unidade",
            "element": {
            "type": "plain_text_input",
            "action_id": "value",
            "max_length": 10
        },
        "label": {"type": "plain_text", "text": "Empreendimento e Unidade"}
        },
        {
            "type": "input",
            "block_id": "numero_reserva",
            "element": {"type": "plain_text_input", "action_id": "value"},
            "label": {"type": "plain_text", "text": "Número da Reserva"}
        },
        {
            "type": "input",
            "block_id": "responsavel",
            "element": {
                "type": "static_select",
                "action_id": "value",
                "placeholder": {"type": "plain_text", "text": "Escolha o responsável"},
                "options": [
                    {"text": {"type": "plain_text", "text": "Marta Cabral"}, "value": "U07S6RFBGE6"},
                    {"text": {"type": "plain_text", "text": "Marina Macena"}, "value": "U06UKKKNJTG"},
                    {"text": {"type": "plain_text", "text": "Alef Nunes"}, "value": "U07DN49NT6V"},
                    {"text": {"type": "plain_text", "text": "Rafaela Assis"}, "value": "U092WBQKE11"},
                    {"text": {"type": "plain_text", "text": "Jelifer Neves"}, "value": "U085ME3BYFP"},
                    {"text": {"type": "plain_text", "text": "Victor Adas"}, "value": "U07B2130TKQ"},
                    {"text": {"type": "plain_text", "text": "Reservas"}, "value": "S08STJCNMHR"}
                ]
            },
            "label": {"type": "plain_text", "text": "Responsável"}
        }
    ]

# 💾 Criação de chamado no banco
def criar_ordem_servico_financeiro(data, thread_ts=None, canal_id=None):
    db = SessionLocal()
    try:
        chamado = OrdemServicoFinanceiro(
            tipo_ticket=data["tipo_ticket"],
            tipo_contrato=data["tipo_contrato"],
            locatario=data["locatario"],
            empreendimento_unidade=data["empreendimento_unidade"],
            numero_reserva=data["numero_reserva"],
            responsavel=data["responsavel"],
            solicitante=data["solicitante"],
            status="aberto",
            data_abertura=datetime.utcnow(),
            thread_ts=thread_ts,
            canal_id=canal_id
        )
        db.add(chamado)
        db.commit()
    except Exception as e:
        print("❌ Erro ao salvar chamado:", e)
        db.rollback()
    finally:
        db.close()

# 🧾 Mensagem formatada
def formatar_mensagem_chamado_financeiro(data, user_id):
    def limpar(valor):
        return valor if valor else "–"

    def formatar_mencao(slack_id):
        if not slack_id:
            return "–"
        if slack_id.startswith("U"):
            return f"<@{slack_id}>"
        if slack_id.startswith("S"):
            return f"<!subteam^{slack_id}>"
        return slack_id

    return (
        f"*Tipo:* {limpar(data.get('tipo_ticket'))}\n"
        f"*Contrato:* {limpar(data.get('tipo_contrato'))}\n"
        f"*Locatário:* {limpar(data.get('locatario'))}\n"
        f"*Empreendimento e Unidade:* {limpar(data.get('empreendimento_unidade'))}\n"
        f"*Número da Reserva:* {limpar(data.get('numero_reserva'))}\n"
        f"*Responsável:* {formatar_mencao(data.get('responsavel'))}\n"
        f"*Solicitante:* {formatar_mencao(user_id)}"
    )

# 🔄 Captura
def capturar_chamado(client, body):
    thread_ts = body["message"]["ts"]
    user_id = body["user"]["id"]
    canal_id = body["channel"]["id"]

    db = SessionLocal()
    chamado = db.query(OrdemServicoFinanceiro).filter_by(thread_ts=thread_ts).first()
    if chamado:
        chamado.status = "em atendimento"
        chamado.data_captura = datetime.now()
        chamado.capturado_por = user_id
        db.commit()
        client.chat_postMessage(channel=canal_id, thread_ts=thread_ts, text=f"🔄 Chamado capturado por <@{user_id}>!")
    else:
        client.chat_postMessage(channel=canal_id, thread_ts=thread_ts, text="❗ Chamado não encontrado.")
    db.close()

# ✅ Finalizar
def finalizar_chamado(client, body):
    thread_ts = body["message"]["ts"]
    user_id = body["user"]["id"]
    canal_id = body["channel"]["id"]

    db = SessionLocal()
    chamado = db.query(OrdemServicoFinanceiro).filter_by(thread_ts=thread_ts).first()
    if chamado:
        chamado.status = "finalizado"
        chamado.data_fechamento = datetime.now()
        db.commit()
        client.chat_postMessage(channel=canal_id, thread_ts=thread_ts, text=f"✅ Chamado finalizado por <@{user_id}>.")
    else:
        client.chat_postMessage(channel=canal_id, thread_ts=thread_ts, text="❗ Chamado não encontrado.")
    db.close()

# ♻️ Reabrir
def abrir_modal_reabertura(client, body):
    client.views_open(
        trigger_id=body["trigger_id"],
        view={
            "type": "modal",
            "callback_id": "reabrir_chamado_modal",
            "title": {"type": "plain_text", "text": "Reabrir Chamado"},
            "submit": {"type": "plain_text", "text": "Reabrir"},
            "private_metadata": body["message"]["ts"],
            "blocks": [{
                "type": "input",
                "block_id": "motivo",
                "element": {
                    "type": "plain_text_input",
                    "action_id": "value",
                    "multiline": True
                },
                "label": {"type": "plain_text", "text": "Motivo da reabertura"}
            }]
        }
    )

def reabrir_chamado(client, body, view):
    ts = view["private_metadata"]
    user_id = body["user"]["id"]
    canal_id = os.getenv("SLACK_CANAL_ID", "C08KMCDNEFR")

    # Recuperar o valor do campo "motivo" de forma segura
    valores = view["state"]["values"]
    motivo = ""
    for block in valores.values():
        for action in block.values():
            motivo = action.get("value", "")
            break  # Só tem um campo

    db = SessionLocal()
    chamado = db.query(OrdemServicoFinanceiro).filter_by(thread_ts=ts).first()

    if chamado:
        chamado.status = "reaberto"
        chamado.historico_reaberturas = json.dumps({
            "motivo": motivo,
            "data": datetime.now().isoformat(),
            "usuario": user_id
        })
        db.commit()
        client.chat_postMessage(
            channel=canal_id,
            thread_ts=ts,
            text=f"♻️ Chamado reaberto por <@{user_id}>!\n*Motivo:* {motivo}"
        )

    db.close()

# ❌ Cancelar
def abrir_modal_cancelamento(client, body):
    client.views_open(
        trigger_id=body["trigger_id"],
        view={
            "type": "modal",
            "callback_id": "cancelar_chamado_modal",
            "title": {"type": "plain_text", "text": "Cancelar Chamado"},
            "submit": {"type": "plain_text", "text": "Confirmar"},
            "private_metadata": body["message"]["ts"],
            "blocks": [{
                "type": "input",
                "block_id": "motivo",
                "element": {
                    "type": "plain_text_input",
                    "action_id": "value",
                    "multiline": True
                },
                "label": {"type": "plain_text", "text": "Motivo do cancelamento"}
            }]
        }
    )

# ✏️ Editar
def abrir_modal_edicao(client, trigger_id, ts):
    from models import OrdemServicoFinanceiro
    db = SessionLocal()
    chamado = db.query(OrdemServicoFinanceiro).filter_by(thread_ts=ts).first()
    db.close()

    if not chamado:
        client.chat_postMessage(
            channel=os.getenv("SLACK_CANAL_ID", "C08KMCDNEFR"),
            thread_ts=ts,
            text="❗ Chamado não encontrado para edição."
        )
        return

    def option(text):
        return {"text": {"type": "plain_text", "text": text}, "value": text}

    def initial_option(valor):
        return option(valor) if valor else None

    client.views_open(
        trigger_id=trigger_id,
        view={
            "type": "modal",
            "callback_id": "editar_chamado_modal",
            "private_metadata": ts,
            "title": {"type": "plain_text", "text": "Editar Chamado"},
            "submit": {"type": "plain_text", "text": "Salvar"},
            "blocks": [
                {
                    "type": "input",
                    "block_id": "tipo_ticket",
                    "element": {
                        "type": "static_select",
                        "action_id": "value",
                        "placeholder": {"type": "plain_text", "text": "Escolha"},
                        "initial_option": initial_option(chamado.tipo_ticket),
                        "options": [option(t) for t in [
                            "Prorrogação de Data", "Saída Antecipada", "Ajuste de Fatura",
                            "Ajuste de Reserva", "Prorrogação de Desconto", "Encerramento de Contrato"
                        ]]
                    },
                    "label": {"type": "plain_text", "text": "Tipo de Ticket"}
                },
                {
                    "type": "input",
                    "block_id": "tipo_contrato",
                    "element": {
                        "type": "static_select",
                        "action_id": "value",
                        "placeholder": {"type": "plain_text", "text": "Escolha"},
                        "initial_option": initial_option(chamado.tipo_contrato),
                        "options": [option(c) for c in ["Short Stay", "Temporada", "Long Stay", "Comodato", "Cortesia"]]
                    },
                    "label": {"type": "plain_text", "text": "Tipo de Contrato"}
                },
                {
                    "type": "input",
                    "block_id": "locatario",
                    "element": {
                        "type": "plain_text_input",
                        "action_id": "value",
                        "initial_value": chamado.locatario or ""
                    },
                    "label": {"type": "plain_text", "text": "Locatário"}
                },
                {
                    "type": "input",
                    "block_id": "empreendimento_unidade",
                    "element": {
                        "type": "plain_text_input",
                        "action_id": "value",
                        "max_length": 10
                    },
                    "label": {"type": "plain_text", "text": "Empreendimento e Unidade"}
                },
                {
                    "type": "input",
                    "block_id": "numero_reserva",
                    "element": {
                        "type": "plain_text_input",
                        "action_id": "value",
                        "initial_value": chamado.numero_reserva or ""
                    },
                    "label": {"type": "plain_text", "text": "Número da Reserva"}
                },
                {
                    "type": "input",
                    "block_id": "responsavel",
                    "element": {
                        "type": "static_select",
                        "action_id": "value",
                        "placeholder": {"type": "plain_text", "text": "Escolha"},
                        "initial_option": initial_option(chamado.responsavel),
                        "options": [
                            option("U06TAJU7C95"),
                            option("U08DRE18RR7"),
                            option("U07S6RFBGE6"),
                            option("U06UKKKNJTG"),
                            option("U07DN49NT6V"),
                            option("U085ME3BYFP"),
                            option("U07B2130TKQ")
                        ]
                    },
                    "label": {"type": "plain_text", "text": "Responsável"}
                }
            ]
        }
    )

# ❌ Cancelamento de chamado
def cancelar_chamado(client, body, view):
    ts = view["private_metadata"]
    motivo = view["state"]["values"]["motivo"]["value"]["value"]
    user_id = body["user"]["id"]
    canal_id = os.getenv("SLACK_CANAL_ID", "C08KMCDNEFR")

    db = SessionLocal()
    chamado = db.query(OrdemServicoFinanceiro).filter_by(thread_ts=ts).first()
    if chamado:
        chamado.status = "cancelado"
        chamado.motivo_cancelamento = motivo
        db.commit()
        client.chat_postMessage(
            channel=canal_id,
            thread_ts=ts,
            text=f"❌ Chamado cancelado por <@{user_id}>!\n*Motivo:* {motivo}"
        )
    db.close()

# ✏️ Edição de chamado
def editar_chamado(client, body, view):
    ts = view["private_metadata"]
    valores = view["state"]["values"]
    user_id = body["user"]["id"]
    canal_id = os.getenv("SLACK_CANAL_ID", "C08KMCDNEFR")

    def pegar(campo):
        bloco = valores.get(campo, {})
        if not bloco:
            return ""
        item = list(bloco.values())[0]
        return item.get("selected_option", {}).get("value") or item.get("value") or ""

    novos_dados = {
        "tipo_ticket": pegar("tipo_ticket"),
        "tipo_contrato": pegar("tipo_contrato"),
        "locatario": pegar("locatario"),
        "empreendimento_unidade": pegar("empreendimento_unidade"),
        "numero_reserva": pegar("numero_reserva"),
        "responsavel": pegar("responsavel")
    }

    db = SessionLocal()
    chamado = db.query(OrdemServicoFinanceiro).filter_by(thread_ts=ts).first()
    if chamado:
        log = {}
        for campo, novo_valor in novos_dados.items():
            antigo = getattr(chamado, campo)
            if str(antigo) != str(novo_valor):
                log[campo] = {"de": antigo, "para": novo_valor}
                setattr(chamado, campo, novo_valor)

        chamado.data_ultima_edicao = datetime.now()
        chamado.ultimo_editor = user_id
        if log:
            chamado.log_edicoes = json.dumps(log, indent=2, ensure_ascii=False)

        db.commit()

        client.chat_postMessage(
            channel=canal_id,
            thread_ts=ts,
            text=f"✏️ Chamado editado por <@{user_id}>"
        )

        # Atualiza visual da thread
        client.chat_update(
            channel=canal_id,
            ts=ts,
            text=f"🧾 ({novos_dados['locatario']}) - {novos_dados['empreendimento_unidade']} <@{chamado.solicitante}>: *{novos_dados['tipo_ticket']}*",
            blocks=[
                {
                    "type": "section",
                    "text": {
                        "type": "mrkdwn",
                        "text": f"🧾 (*{novos_dados['locatario']}*) - {novos_dados['empreendimento_unidade']} <@{chamado.solicitante}>: *{novos_dados['tipo_ticket']}*"
                    }
                },
                {
                    "type": "actions",
                    "elements": [
                        {"type": "button", "text": {"type": "plain_text", "text": "🔄 Capturar"}, "action_id": "capturar_chamado"},
                        {"type": "button", "text": {"type": "plain_text", "text": "✅ Finalizar"}, "action_id": "finalizar_chamado"},
                        {"type": "button", "text": {"type": "plain_text", "text": "♻️ Reabrir"}, "action_id": "reabrir_chamado"},
                        {"type": "button", "text": {"type": "plain_text", "text": "❌ Cancelar"}, "action_id": "cancelar_chamado"},
                        {"type": "button", "text": {"type": "plain_text", "text": "✏️ Editar"}, "action_id": "editar_chamado"}
                    ]
                }
            ]
        )

    db.close()

import csv
import io
import os
import json
import urllib.request
from fpdf import FPDF
from database import SessionLocal
from models import OrdemServicoFinanceiro
from datetime import datetime
from slack_sdk import WebClient

import unicodedata

# 🔧 Limpa texto para evitar erros de encoding (PDF e outros)
def limpar_texto_pdf(texto):
    if not texto:
        return ""
    texto = texto.replace("–", "-").replace("—", "-")
    texto = texto.replace("“", '"').replace("”", '"')
    texto = texto.replace("‘", "'").replace("’", "'")
    texto = unicodedata.normalize("NFKD", texto).encode("latin-1", "ignore").decode("latin-1")
    return texto

# 🗂️ Gerar PDF
def gerar_pdf_chamados_financeiro(chamados, user_id):
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()

    LOGO_URL = "https://raw.githubusercontent.com/jflrealty/images/main/JFL_logotipo_completo.jpg"
    logo_path = "/tmp/logo_jfl.jpg"
    try:
        urllib.request.urlretrieve(LOGO_URL, logo_path)
        pdf.image(logo_path, x=10, y=8, w=60)
    except Exception as e:
        print(f"❌ Erro ao carregar logo: {e}")

    pdf.set_font("Arial", "B", 10)
    pdf.set_fill_color(230, 230, 230)
    col_widths = [10, 40, 30, 40, 25, 30, 30, 20, 30]
    headers = ["ID", "Tipo", "Contrato", "Locatário", "Empreendimento", "Reserva", "Responsável", "Status", "Aberto em"]

    for i, header in enumerate(headers):
        pdf.cell(col_widths[i], 8, limpar_texto_pdf(header), border=1, fill=True)
    pdf.ln()

    pdf.set_font("Arial", "", 9)
    for ch in chamados:
        row = [
            ch.id,
            ch.tipo_ticket,
            ch.tipo_contrato,
            ch.locatario,
            ch.empreendimento_unidade,
            ch.numero_reserva,
            get_nome_slack(ch.responsavel),
            ch.status,
            ch.data_abertura.strftime('%d/%m/%Y %H:%M')
        ]
        for i, cell in enumerate(row):
            pdf.cell(col_widths[i], 8, limpar_texto_pdf(str(cell)), border=1)
        pdf.ln()

    pdf_bytes = pdf.output(dest="S").encode("latin-1")
    return io.BytesIO(pdf_bytes)

# 🧾 Gerar CSV
def gerar_csv_chamados_financeiro(chamados):
    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow([
        "ID", "Tipo", "Contrato", "Locatário", "Empreendimento/Unidade",
        "Reserva", "Responsável", "Solicitante", "Status", "Aberto em"
    ])
    for ch in chamados:
        writer.writerow([
            ch.id,
            limpar_texto_pdf(ch.tipo_ticket),
            limpar_texto_pdf(ch.tipo_contrato),
            limpar_texto_pdf(ch.locatario),
            limpar_texto_pdf(ch.empreendimento_unidade),
            limpar_texto_pdf(ch.numero_reserva),
            limpar_texto_pdf(get_nome_slack(ch.responsavel)),
            limpar_texto_pdf(get_nome_slack(ch.solicitante)),
            limpar_texto_pdf(ch.status),
            ch.data_abertura.strftime('%d/%m/%Y %H:%M')
        ])
    output.seek(0)
    return io.BytesIO(output.read().encode("utf-8"))

# 🧵 Buscar chamados por usuário
def buscar_chamados_por_usuario_financeiro(user_id):
    db = SessionLocal()
    chamados = db.query(OrdemServicoFinanceiro).filter_by(solicitante=user_id).order_by(OrdemServicoFinanceiro.data_abertura.desc()).all()
    db.close()
    return chamados

# 📁 Buscar todos os chamados
def buscar_todos_chamados_financeiro():
    db = SessionLocal()
    chamados = db.query(OrdemServicoFinanceiro).order_by(OrdemServicoFinanceiro.data_abertura.desc()).all()
    db.close()
    return chamados
# ✅ Listar chamados do usuário (para o comando /meus-chamados-financeiro)
def listar_chamados_por_usuario(user_id):
    db = SessionLocal()
    chamados = db.query(OrdemServicoFinanceiro)\
                 .filter_by(solicitante=user_id)\
                 .order_by(OrdemServicoFinanceiro.data_abertura.desc())\
                 .all()
    db.close()
    return chamados


# ✅ Exportar chamados (para o comando /exportar-chamados-financeiro)
def exportar_chamados_financeiro(data_inicio=None, data_fim=None, tipo="pdf"):
    chamados = buscar_todos_chamados_financeiro()

    # Filtrar por data de abertura
    if data_inicio:
        chamados = [c for c in chamados if c.data_abertura >= data_inicio]
    if data_fim:
        chamados = [c for c in chamados if c.data_abertura <= data_fim]

    # Timestamp para nome do arquivo
    agora = datetime.now().strftime("%Y%m%d%H%M%S")

    return {
        "pdf": (
            f"chamados-financeiro-{agora}.pdf",
            gerar_pdf_chamados_financeiro(chamados, "admin")
        ),
        "csv": (
            f"chamados-financeiro-{agora}.csv",
            gerar_csv_chamados_financeiro(chamados)
        ),
        "xlsx": (
            f"chamados-financeiro-{agora}.xlsx",
            gerar_xlsx_chamados_financeiro(chamados)
        )
    }

# 📦 Modal de exportação com filtro por data e tipo
def montar_blocos_exportacao():
    return [
        {
            "type": "input",
            "block_id": "tipo_arquivo",
            "label": {"type": "plain_text", "text": "Formato do Arquivo"},
            "element": {
                "type": "static_select",
                "action_id": "value",
                "placeholder": {"type": "plain_text", "text": "Escolha o formato"},
                "options": [
                    {"text": {"type": "plain_text", "text": "PDF"}, "value": "pdf"},
                    {"text": {"type": "plain_text", "text": "CSV"}, "value": "csv"},
                    {"text": {"type": "plain_text", "text": "Excel"}, "value": "xlsx"}
                ]
            }
        },
        {
            "type": "input",
            "block_id": "data_inicio",
            "label": {"type": "plain_text", "text": "Data Inicial"},
            "element": {
                "type": "datepicker",
                "action_id": "value",
                "placeholder": {"type": "plain_text", "text": "Escolha a data inicial"}
            }
        },
        {
            "type": "input",
            "block_id": "data_fim",
            "label": {"type": "plain_text", "text": "Data Final"},
            "element": {
                "type": "datepicker",
                "action_id": "value",
                "placeholder": {"type": "plain_text", "text": "Escolha a data final"}
            }
        }
    ]

def gerar_xlsx_chamados_financeiro(chamados):
    wb = Workbook()
    ws = wb.active
    ws.title = "Chamados Financeiros"

    colunas = [
        "ID", "Tipo", "Contrato", "Locatário", "Empreendimento/Unidade",
        "Reserva", "Responsável", "Capturado por", "Solicitante", "Status", "Aberto em"
    ]
    ws.append(colunas)

    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center")

    def nome(user_id):
        return "Reservas" if user_id == "S08STJCNMHR" else get_nome_slack(user_id)

    for c in chamados:
        ws.append([
            c.id,
            limpar_texto_pdf(c.tipo_ticket),
            limpar_texto_pdf(c.tipo_contrato),
            limpar_texto_pdf(c.locatario),
            limpar_texto_pdf(c.empreendimento_unidade),
            limpar_texto_pdf(c.numero_reserva),
            limpar_texto_pdf(nome(c.responsavel)),
            limpar_texto_pdf(nome(c.capturado_por)),
            limpar_texto_pdf(nome(c.solicitante)),
            limpar_texto_pdf(c.status),
            c.data_abertura.strftime('%d/%m/%Y %H:%M')
        ])

    for col in ws.columns:
        max_length = max(len(str(cell.value)) for cell in col)
        ws.column_dimensions[get_column_letter(col[0].column)].width = max(max_length + 2, 12)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output
