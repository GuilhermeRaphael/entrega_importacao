"""
core.py
Lógica de negócio do Importador Goalfy.

Portado a partir do VBA original (frmSelecaoMatrizes + Módulo de consolidação),
com o passo final de envio por e-mail (Outlook) substituído por envio direto
ao webhook do Goalfy Flow (a mesma lógica que já existia em app.py/teste2.py).
"""

from __future__ import annotations

import json
import os
from dataclasses import dataclass
from datetime import datetime, timedelta
from typing import Any

import openpyxl
import pandas as pd
import requests

CONFIG_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")

# Colunas da MATRIZ (1-indexed, como no VBA)
COL_MATRIZ_N = 1
COL_MATRIZ_NOME = 2
COL_MATRIZ_TELEFONE = 3
COL_MATRIZ_DOC = 4  # CPF ou CNPJ "cru"
COL_MATRIZ_EMAIL = 5
COL_MATRIZ_CHEGADA_LEAD = 6
COL_MATRIZ_DIA_RECOLHE = 7
COL_MATRIZ_RESPONSAVEL = 8
COL_MATRIZ_PRODUTO = 9
COL_MATRIZ_EMPRESA = 10
COL_MATRIZ_MODELO = 11
COL_MATRIZ_DATA_ENTREGA = 12

HEADERS_BASE = [
    "n",
    "nome",
    "telefone",
    "colarCpf",
    "colarCnpj",
    "email",
    "chegadaDoLead",
    "diaRecolhe",
    "responsavel",
    "produto",
    "empresa",
    "modelo",
    "cpf",
    "cnpj",
]
HEADER_BONIF = "bonif"


# ---------------------------------------------------------------------------
# Configuração
# ---------------------------------------------------------------------------


def load_config(path: str = CONFIG_PATH) -> dict:
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def save_config(config: dict, path: str = CONFIG_PATH) -> None:
    with open(path, "w", encoding="utf-8") as f:
        json.dump(config, f, ensure_ascii=False, indent=2)


def listar_produtos_ativos(config: dict) -> list[str]:
    """Equivalente ao carregamento do cboProduto no UserForm_Initialize."""
    nomes = []
    for m in config.get("matrizes", []):
        if m.get("caminho") and m.get("ativado", False):
            nomes.append(m.get("exibicao") or m.get("id"))
    return nomes


def obter_caminho_matriz(config: dict, exibicao_ou_id: str) -> str:
    """Equivalente a ObterCaminhoMatrizPorExibicao."""
    alvo = (exibicao_ou_id or "").strip().upper()
    for m in config.get("matrizes", []):
        if (m.get("exibicao") or "").strip().upper() == alvo or (
            m.get("id") or ""
        ).strip().upper() == alvo:
            return m.get("caminho", "")
    return ""


# ---------------------------------------------------------------------------
# Item de fila (remessa)
# ---------------------------------------------------------------------------


@dataclass
class ItemFila:
    produto: str
    modelo: str
    vendedor: str
    quantidade: int


# ---------------------------------------------------------------------------
# Leitura da matriz / cálculo de saldo por modelo
# (equivalente a AtualizarListaModelosComQtd do UserForm)
# ---------------------------------------------------------------------------


def _linha_elegivel(row: tuple) -> bool:
    """Uma linha é elegível se tem nome/telefone/email e ainda não foi puxada
    (responsavel, diaRecolhe e dataEntrega em branco)."""
    nome = str(row[COL_MATRIZ_NOME - 1] or "").strip()
    tel = str(row[COL_MATRIZ_TELEFONE - 1] or "").strip()
    email = str(row[COL_MATRIZ_EMAIL - 1] or "").strip()
    if not (nome or tel or email):
        return False

    resp = str(row[COL_MATRIZ_RESPONSAVEL - 1] or "").strip()
    dia_recolhe = str(row[COL_MATRIZ_DIA_RECOLHE - 1] or "").strip()
    dt_entrega = str(row[COL_MATRIZ_DATA_ENTREGA - 1] or "").strip()
    return resp == "" and dia_recolhe == "" and dt_entrega == ""


def obter_modelos_disponiveis(
    caminho_arq: str,
    fila_atual: list[ItemFila] | None = None,
    produto_atual: str | None = None,
) -> dict[str, int]:
    """Retorna {MODELO: saldo_disponivel} para a matriz, descontando o que já
    está enfileirado nesta sessão para o mesmo produto."""
    if not caminho_arq or not os.path.isfile(caminho_arq):
        raise FileNotFoundError(f"Matriz não encontrada: {caminho_arq}")

    wb = openpyxl.load_workbook(caminho_arq, read_only=True, data_only=True)
    ws = wb.worksheets[0]

    contagem: dict[str, int] = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or len(row) < COL_MATRIZ_DATA_ENTREGA:
            continue
        if _linha_elegivel(row):
            modelo = str(row[COL_MATRIZ_MODELO - 1] or "").strip().upper() or "LEAD"
            contagem[modelo] = contagem.get(modelo, 0) + 1
    wb.close()

    # Desconta o que já está na fila (mesmo produto)
    na_fila: dict[str, int] = {}
    if fila_atual and produto_atual:
        for item in fila_atual:
            if item.produto.strip().upper() == produto_atual.strip().upper():
                modelo_key = item.modelo.strip().upper()
                na_fila[modelo_key] = na_fila.get(modelo_key, 0) + item.quantidade

    saldo = {}
    for modelo, qtd in contagem.items():
        restante = qtd - na_fila.get(modelo, 0)
        if restante > 0:
            saldo[modelo] = restante
    return saldo


# ---------------------------------------------------------------------------
# Formatação (idêntico ao app.py / teste2.py originais)
# ---------------------------------------------------------------------------

import re

_ISO_DATE_RE = re.compile(r"^\d{4}-\d{2}-\d{2}")


def formatar_data_estrita(valor: Any) -> str:
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return ""

    # Já é datetime/date (caso comum: célula de data do Excel lida via openpyxl)
    if hasattr(valor, "strftime"):
        return valor.strftime("%Y-%m-%d")

    val_str = str(valor).strip()
    if not val_str or val_str.lower() in ("nan", "none", "nat"):
        return ""

    # Se já vier em ISO (YYYY-MM-DD...), não usar dayfirst (senão inverte mês/dia)
    if _ISO_DATE_RE.match(val_str):
        return val_str[:10]

    try:
        data_dt = pd.to_datetime(val_str, dayfirst=True)
        return data_dt.strftime("%Y-%m-%d")
    except (ValueError, TypeError, OverflowError):
        return val_str.split(" ")[0].split("T")[0]


def _limpar_documento(val_str: str) -> str:
    if "." in val_str and val_str.replace(".", "").isdigit() is False:
        pass
    val_str = val_str.removesuffix(".0")
    return val_str.replace("-", "").replace(".", "").replace("/", "").strip()


def formatar_cpf(valor: Any) -> str:
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return ""
    val_str = str(valor).strip()
    if not val_str or val_str.lower() in ("nan", "none"):
        return ""
    val_limpo = _limpar_documento(val_str)
    if not val_limpo:
        return ""
    return val_limpo.zfill(11)


def formatar_cnpj(valor: Any) -> str:
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return ""
    val_str = str(valor).strip()
    if not val_str or val_str.lower() in ("nan", "none"):
        return ""
    val_limpo = _limpar_documento(val_str)
    if not val_limpo:
        return ""
    return val_limpo.zfill(14)


# ---------------------------------------------------------------------------
# Consolidação (equivalente a ExecutarConsolidacao)
# ---------------------------------------------------------------------------


def consolidar_remessa(
    config: dict, fila: list[ItemFila]
) -> tuple[list[dict], bool, list[str]]:
    """Processa a fila de remessa: abre cada matriz, marca as linhas puxadas
    (responsavel/modelo/datas) e monta os registros consolidados.

    Retorna (registros, tem_auto_bonif, avisos).
    """
    registros: list[dict] = []
    avisos: list[str] = []

    tem_auto_bonif = any(item.produto.strip().upper() == "AUTO BONIF" for item in fila)

    for item in fila:
        caminho_arq = obter_caminho_matriz(config, item.produto)
        if not caminho_arq or not os.path.isfile(caminho_arq):
            avisos.append(
                f"Matriz não encontrada para o produto '{item.produto}'. Item ignorado."
            )
            continue

        wb = openpyxl.load_workbook(caminho_arq)  # mantém fórmulas eventuais intactas
        ws = wb.worksheets[0]

        modelo_nome = item.modelo.strip().upper()
        qtd_coletada = 0

        for row_cells in ws.iter_rows(min_row=2):
            if qtd_coletada >= item.quantidade:
                break

            def val(col):
                return row_cells[col - 1].value

            nome = str(val(COL_MATRIZ_NOME) or "").strip()
            tel = str(val(COL_MATRIZ_TELEFONE) or "").strip()
            email = str(val(COL_MATRIZ_EMAIL) or "").strip()
            if not (nome or tel or email):
                continue

            resp_existente = str(val(COL_MATRIZ_RESPONSAVEL) or "").strip()
            dia_recolhe_existente = str(val(COL_MATRIZ_DIA_RECOLHE) or "").strip()
            dt_entrega_existente = str(val(COL_MATRIZ_DATA_ENTREGA) or "").strip()
            if not (
                resp_existente == ""
                and dia_recolhe_existente == ""
                and dt_entrega_existente == ""
            ):
                continue

            modelo_matriz_existente = str(val(COL_MATRIZ_MODELO) or "").strip().upper()
            if not (
                modelo_matriz_existente == "" or modelo_matriz_existente == modelo_nome
            ):
                continue

            # --- marca a linha como puxada ---
            data_entrega_dt = datetime.now().date()
            # Weekday(vbMonday-based): segunda=1..domingo=7. Python weekday(): segunda=0..domingo=6
            dias_add = (
                5 if data_entrega_dt.weekday() >= 2 else 3
            )  # >=3 no VBA (1-based) == >=2 (0-based)
            dia_recolhe_dt = data_entrega_dt + timedelta(days=dias_add)

            row_cells[COL_MATRIZ_RESPONSAVEL - 1].value = item.vendedor
            row_cells[COL_MATRIZ_MODELO - 1].value = modelo_nome
            row_cells[COL_MATRIZ_DATA_ENTREGA - 1].value = data_entrega_dt.strftime(
                "%Y-%m-%d"
            )
            row_cells[COL_MATRIZ_DIA_RECOLHE - 1].value = dia_recolhe_dt.strftime(
                "%Y-%m-%d"
            )

            # --- monta o registro consolidado ---
            doc_raw = str(val(COL_MATRIZ_DOC) or "").strip()
            doc_limpo = (
                doc_raw.replace(".", "")
                .replace("-", "")
                .replace("/", "")
                .replace(" ", "")
            )

            if len(doc_limpo) > 11:
                colar_cpf, colar_cnpj = "", doc_raw
            else:
                colar_cpf, colar_cnpj = doc_raw, ""

            chegada_lead = val(COL_MATRIZ_CHEGADA_LEAD)
            chegada_lead_fmt = (
                formatar_data_estrita(chegada_lead) if chegada_lead else ""
            )

            registro = {
                "n": val(COL_MATRIZ_N),
                "nome": nome,
                "telefone": tel,
                "colarCpf": colar_cpf,
                "colarCnpj": colar_cnpj,
                "email": email,
                "chegadaDoLead": chegada_lead_fmt,
                "diaRecolhe": dia_recolhe_dt.strftime("%Y-%m-%d"),
                "responsavel": item.vendedor,
                "produto": val(COL_MATRIZ_PRODUTO) or item.produto,
                "empresa": val(COL_MATRIZ_EMPRESA) or "",
                "modelo": modelo_nome,
                "cpf": formatar_cpf(colar_cpf),
                "cnpj": formatar_cnpj(colar_cnpj),
            }
            if tem_auto_bonif:
                registro[HEADER_BONIF] = (
                    "BONIFICACAO"
                    if item.produto.strip().upper() == "AUTO BONIF"
                    else ""
                )

            registros.append(registro)
            qtd_coletada += 1

        wb.save(caminho_arq)
        wb.close()

        if qtd_coletada < item.quantidade:
            avisos.append(
                f"Produto '{item.produto}' / modelo '{item.modelo}': solicitados {item.quantidade}, "
                f"encontrados apenas {qtd_coletada} leads elegíveis."
            )

    return registros, tem_auto_bonif, avisos


# ---------------------------------------------------------------------------
# Salvamento da planilha de backup (equivalente à aba IMPORTAÇÃO + SaveAs)
# ---------------------------------------------------------------------------


def salvar_planilha_importacao(
    config: dict, registros: list[dict], tem_auto_bonif: bool
) -> str:
    from openpyxl.styles import Alignment, Font, PatternFill

    headers = HEADERS_BASE + ([HEADER_BONIF] if tem_auto_bonif else [])

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "IMPORTAÇÃO"
    ws.append(headers)

    header_fill = PatternFill(
        start_color="002060", end_color="002060", fill_type="solid"
    )
    header_font = Font(color="FFFFFF", bold=True)
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for reg in registros:
        ws.append([reg.get(h, "") for h in headers])

    for col_cells in ws.columns:
        largura = (
            max((len(str(c.value)) if c.value is not None else 0) for c in col_cells)
            + 2
        )
        ws.column_dimensions[col_cells[0].column_letter].width = min(largura, 40)

    pasta_destino = config.get("pasta_destino", "").rstrip("\\/")
    agora = datetime.now()
    nome_mes = agora.strftime("%B").upper()
    pasta_mes = os.path.join(
        pasta_destino, f"{agora.month} - {nome_mes} - {agora.strftime('%y')}"
    )
    pasta_dia = os.path.join(pasta_mes, agora.strftime("%d-%m-%Y"))
    os.makedirs(pasta_dia, exist_ok=True)

    nome_arquivo = f"IMPORTAÇÃO GERAL GOALFY - {agora.strftime('%Y-%m-%d_%Hh%M')}.xlsx"
    caminho_saida = os.path.join(pasta_dia, nome_arquivo)
    wb.save(caminho_saida)
    return caminho_saida


# ---------------------------------------------------------------------------
# Envio ao Goalfy (equivalente a app.py / teste2.py)
# ---------------------------------------------------------------------------


def montar_payload(registro: dict) -> dict:
    def s(key):
        v = registro.get(key, "")
        return "" if v is None else str(v)

    return {
        "n": s("n"),
        "nome": s("nome"),
        "telefone": s("telefone"),
        "email": s("email"),
        "chegadaDoLead": registro.get("chegadaDoLead", ""),
        "diaRecolhe": registro.get("diaRecolhe", ""),
        "responsavel": s("responsavel"),
        "produto": s("produto"),
        "empresa": s("empresa"),
        "modelo": s("modelo"),
        "cpf": s("cpf"),
        "cnpj": s("cnpj"),
        "bonif": s("bonif"),
    }


def enviar_para_goalfy(
    config: dict, registros: list[dict], progress_callback=None
) -> tuple[int, int, list[str]]:
    """Envia cada registro para o webhook do Goalfy.
    progress_callback(index, total, sucesso: bool) é chamado a cada envio."""
    webhook_url = config.get("webhook_url", "")
    headers = {"Content-Type": "application/json", "Accept": "application/json"}

    sucessos = 0
    falhas = 0
    erros: list[str] = []
    total = len(registros)

    for i, registro in enumerate(registros):
        payload = montar_payload(registro)
        ok = False
        status_code = None
        try:
            resp = requests.post(webhook_url, json=payload, headers=headers, timeout=15)
            status_code = resp.status_code
            ok = status_code in (200, 201)
        except requests.RequestException as e:
            erros.append(
                f"Linha {i + 1} ({registro.get('nome', '')}): erro de conexão - {e}"
            )

        if ok:
            sucessos += 1
        else:
            falhas += 1
            if status_code is not None:
                erros.append(
                    f"Linha {i + 1} ({registro.get('nome', '')}): status {status_code}"
                )

        if progress_callback:
            progress_callback(i + 1, total, ok)

    return sucessos, falhas, erros


# ---------------------------------------------------------------------------
# Agrupamento por vendedor / envio de e-mail (Outlook) / impressão
# (equivalente a DispararEntregaDeLeads e ImprimirTabelaLeads do VBA)
# ---------------------------------------------------------------------------

import platform
import tempfile

from openpyxl.utils import get_column_letter

HEADERS_GRUPO = [
    "Nº",
    "NOME",
    "TELEFONE",
    "CPF",
    "CNPJ",
    "E-MAIL",
    "CHEGADA DO LEAD",
    "DATA RECOLHE",
    "RESPONSAVEL",
    "PRODUTO",
    "EMPRESA",
    "MODELO",
]


def agrupar_leads(registros: list[dict]) -> dict[tuple, list[dict]]:
    """Agrupa registros por (vendedor, produto, empresa), na ordem de chegada,
    igual à lógica do dictGrupos no VBA."""
    grupos: dict[tuple, list[dict]] = {}
    for r in registros:
        vend = str(r.get("responsavel", "")).strip()
        prod = str(r.get("produto", "")).strip()
        emp = str(r.get("empresa", "")).strip()
        if not vend:
            continue
        chave = (vend, prod, emp)
        grupos.setdefault(chave, []).append(r)
    return grupos


def titulo_grupo(chave: tuple) -> str:
    vend, prod, emp = chave
    if prod and emp:
        return f"{prod} {emp}"
    return prod or emp


def obter_email_vendedor(config: dict, nome_vendedor: str) -> str:
    alvo = (nome_vendedor or "").strip().upper()
    for v in config.get("vendedores", []):
        if (v.get("nome") or "").strip().upper() == alvo:
            return v.get("email", "")
    return ""


def gerar_planilha_grupo(
    chave: tuple, registros_grupo: list[dict], pasta_temp: str | None = None
) -> str:
    """Gera um .xlsx formatado com os leads de um único grupo (vendedor/produto/empresa),
    equivalente à planilha temporária criada no VBA antes de anexar ao e-mail / imprimir."""
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

    vend, prod, emp = chave
    titulo = titulo_grupo(chave)
    agora = datetime.now()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "LEADS"

    ws.merge_cells("B2:M2")
    cel_titulo = ws["B2"]
    cel_titulo.value = f"{titulo} - {agora.strftime('%d/%m/%Y - %Hh%M')}"
    cel_titulo.font = Font(bold=True, size=14, name="Calibri")
    cel_titulo.alignment = Alignment(horizontal="center", vertical="center")

    for idx, h in enumerate(HEADERS_GRUPO, start=2):
        c = ws.cell(row=3, column=idx, value=h)
        c.fill = PatternFill(
            start_color="002060", end_color="002060", fill_type="solid"
        )
        c.font = Font(bold=True, color="FFFFFF", name="Calibri")
        c.alignment = Alignment(horizontal="center", vertical="center")

    thin = Side(style="thin", color="000000")
    borda = Border(left=thin, right=thin, top=thin, bottom=thin)

    linha = 4
    for r in registros_grupo:
        valores = [
            r.get("n", ""),
            r.get("nome", ""),
            r.get("telefone", ""),
            r.get("cpf", ""),
            r.get("cnpj", ""),
            r.get("email", ""),
            r.get("chegadaDoLead", ""),
            r.get("diaRecolhe", ""),
            r.get("responsavel", ""),
            r.get("produto", ""),
            r.get("empresa", ""),
            r.get("modelo", ""),
        ]
        for idx, v in enumerate(valores, start=2):
            c = ws.cell(row=linha, column=idx, value=v)
            c.border = borda
            c.alignment = Alignment(horizontal="center", vertical="center")
        linha += 1

    for col_idx in range(2, 14):
        col_letter = get_column_letter(col_idx)
        largura = (
            max(
                (
                    len(str(ws.cell(row=r, column=col_idx).value or ""))
                    for r in range(3, linha)
                ),
                default=10,
            )
            + 2
        )
        ws.column_dimensions[col_letter].width = min(largura, 40)

    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    pasta_temp = pasta_temp or tempfile.gettempdir()
    os.makedirs(pasta_temp, exist_ok=True)
    data_hora = agora.strftime("%d-%m-%Y %Hh%M")
    nome_arquivo = f"LEADS {prod} {emp} - {vend.upper()} - {data_hora}.xlsx".replace(
        "  ", " "
    )
    # remove caracteres inválidos para nome de arquivo no Windows
    for ch in '<>:"/\\|?*':
        nome_arquivo = nome_arquivo.replace(ch, "")
    caminho = os.path.join(pasta_temp, nome_arquivo)
    wb.save(caminho)
    return caminho


def enviar_email_outlook(
    destinatario: str, nome_vendedor: str, titulo: str, caminho_anexo: str
) -> None:
    """Envia e-mail com anexo via Outlook instalado (COM/pywin32).
    Só funciona no Windows com o Outlook instalado e aberto (ou com perfil configurado),
    exatamente como no VBA original."""
    if platform.system() != "Windows":
        raise RuntimeError(
            "Envio via Outlook só funciona no Windows com o Outlook instalado. "
            "(Ambiente atual não é Windows.)"
        )
    try:
        import pythoncom  # type: ignore
        import win32com.client  # type: ignore
    except ImportError as e:
        raise RuntimeError(
            "Biblioteca 'pywin32' não encontrada. Instale com: pip install pywin32"
        ) from e

    # O envio roda numa thread separada (para não travar a interface). O COM
    # exige que cada thread que o usa chame CoInitialize antes de criar
    # objetos COM (Outlook.Application), senão dá o erro -2147221008
    # "CoInitialize não foi chamado".
    pythoncom.CoInitialize()
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)  # olMailItem
        mail.To = destinatario
        mail.Subject = os.path.basename(caminho_anexo).rsplit(".", 1)[0]
        mail.Body = (
            f"Olá {nome_vendedor},\n\n"
            f"Segue em anexo a sua lista de leads referente a {titulo}."
        )
        mail.Attachments.Add(caminho_anexo)
        mail.Send()
    finally:
        pythoncom.CoUninitialize()


def imprimir_arquivo(caminho_anexo: str) -> None:
    """Envia o arquivo para a impressora padrão do Windows (usa o Excel/handler
    associado ao .xlsx, igual ao PrintOut do VBA)."""
    if platform.system() != "Windows":
        raise RuntimeError(
            "Impressão direta só é suportada no Windows (usa o Excel instalado)."
        )
    os.startfile(caminho_anexo, "print")  # type: ignore[attr-defined]


def enviar_emails_para_grupos(
    config: dict, grupos_selecionados: dict[tuple, list[dict]], progress_callback=None
) -> tuple[int, int, list[str]]:
    sucesso, falha = 0, 0
    avisos: list[str] = []
    total = len(grupos_selecionados)

    for i, (chave, registros_grupo) in enumerate(grupos_selecionados.items(), start=1):
        vend = chave[0]
        email_vend = obter_email_vendedor(config, vend)
        ok = False
        if not email_vend:
            falha += 1
            avisos.append(
                f"E-mail do vendedor '{vend}' não cadastrado na aba Vendedores. Envio pulado."
            )
        else:
            try:
                caminho = gerar_planilha_grupo(chave, registros_grupo)
                enviar_email_outlook(email_vend, vend, titulo_grupo(chave), caminho)
                sucesso += 1
                ok = True
            except Exception as e:
                falha += 1
                avisos.append(f"Falha ao enviar e-mail para '{vend}': {e}")
        if progress_callback:
            progress_callback(i, total, ok)

    return sucesso, falha, avisos


def imprimir_grupos(
    grupos_selecionados: dict[tuple, list[dict]], progress_callback=None
) -> tuple[int, int, list[str]]:
    sucesso, falha = 0, 0
    avisos: list[str] = []
    total = len(grupos_selecionados)

    for i, (chave, registros_grupo) in enumerate(grupos_selecionados.items(), start=1):
        ok = False
        try:
            caminho = gerar_planilha_grupo(chave, registros_grupo)
            imprimir_arquivo(caminho)
            sucesso += 1
            ok = True
        except Exception as e:
            falha += 1
            avisos.append(f"Falha ao imprimir grupo '{titulo_grupo(chave)}': {e}")
        if progress_callback:
            progress_callback(i, total, ok)

    return sucesso, falha, avisos


# ---------------------------------------------------------------------------
# CRUD de matrizes e vendedores (para edição dentro do app)
# ---------------------------------------------------------------------------


def adicionar_matriz(
    config: dict, id_: str, exibicao: str, caminho: str, ativado: bool
) -> None:
    config.setdefault("matrizes", []).append(
        {
            "id": id_,
            "exibicao": exibicao,
            "caminho": caminho,
            "ativado": ativado,
        }
    )


def remover_matriz(config: dict, id_: str) -> None:
    config["matrizes"] = [m for m in config.get("matrizes", []) if m.get("id") != id_]


def adicionar_vendedor(config: dict, nome: str, email: str) -> None:
    config.setdefault("vendedores", []).append({"nome": nome, "email": email})


def remover_vendedor(config: dict, nome: str) -> None:
    alvo = (nome or "").strip().upper()
    config["vendedores"] = [
        v
        for v in config.get("vendedores", [])
        if (v.get("nome") or "").strip().upper() != alvo
    ]
