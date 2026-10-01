"""Proteção persistente contra repetição da mesma importação no Goalfy."""
import hashlib
import json
import os
import tempfile
import threading
import core

HISTORICO_PATH = os.path.join(os.path.dirname(core.CONFIG_PATH), "historico_envios_goalfy.json")
_LOCK = threading.Lock()

class EnvioRepetidoError(ValueError):
    pass


def identidade_importacao(registros):
    payloads = [{k: "" if v is None else str(v) for k,v in core.montar_payload(r).items()} for r in registros]
    dados = json.dumps(payloads, sort_keys=True, ensure_ascii=False, separators=(",", ":"))
    return hashlib.sha256(dados.encode("utf-8")).hexdigest()


def _ler(caminho):
    if not os.path.exists(caminho):
        return {"importacoes": {}}
    with open(caminho, encoding="utf-8") as f:
        historico = json.load(f)
    if not isinstance(historico.get("importacoes"), dict):
        raise ValueError("Histórico de envios inválido. Não foi possível verificar a trava.")
    return historico


def _salvar(caminho, historico):
    pasta = os.path.dirname(os.path.abspath(caminho))
    os.makedirs(pasta, exist_ok=True)
    temp = None
    try:
        with tempfile.NamedTemporaryFile(mode="w", encoding="utf-8", dir=pasta, prefix=".goalfy_", suffix=".json", delete=False) as f:
            temp = f.name
            json.dump(historico, f, ensure_ascii=False, indent=2)
        os.replace(temp, caminho)
        temp = None
    finally:
        if temp and os.path.exists(temp):
            os.remove(temp)


def indices_pendentes(registros, caminho_historico=None):
    historico = _ler(caminho_historico or HISTORICO_PATH)
    dados = historico["importacoes"].get(identidade_importacao(registros), {})
    enviados = set(dados.get("enviados", []))
    return [i for i in range(len(registros)) if i not in enviados]


def enviar_protegido(config, registros, indices=None, progress_callback=None, caminho_historico=None):
    caminho = caminho_historico or HISTORICO_PATH
    with _LOCK:
        historico = _ler(caminho)
        chave = identidade_importacao(registros)
        entrega = historico["importacoes"].setdefault(chave, {"enviados": [], "total": len(registros)})
        enviados = set(entrega["enviados"])
        escolhidos = list(dict.fromkeys(range(len(registros)) if indices is None else indices))
        if any(i < 0 or i >= len(registros) for i in escolhidos):
            raise ValueError("Seleção de leads inválida.")
        pendentes = [i for i in escolhidos if i not in enviados]
        if not pendentes:
            raise EnvioRepetidoError("Envio bloqueado: esta importação (ou os leads selecionados) já foi enviada ao Goalfy. Não é possível enviar a mesma planilha duas vezes.")
        # Valida a gravação do histórico antes do primeiro disparo.
        _salvar(caminho, historico)
        def progresso(numero, total, ok):
            if ok:
                enviados.add(pendentes[numero - 1])
                entrega["enviados"] = sorted(enviados)
                _salvar(caminho, historico)
            if progress_callback:
                progress_callback(numero, total, ok)
        return core.enviar_para_goalfy(config, [registros[i] for i in pendentes], progresso)
