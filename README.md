# Importador Goalfy (Python)

Porte do seu fluxo em VBA (UserForm `frmSelecaoMatrizes` + Módulo de consolidação
+ macros `DispararEntregaDeLeads` e `ImprimirTabelaLeads`) para um app Python de
desktop, com envio direto ao webhook do Goalfy Flow, envio de e-mail por
vendedor (via Outlook) e impressão por vendedor.

## O que foi portado 1:1

- **Tela de seleção**: Produto → Modelo (com saldo em tempo real, descontando o
  que já está na fila) → Vendedor → Quantidade → "Adicionar à fila".
- **Consolidação**: abre a matriz do produto, pega linhas elegíveis (sem
  `responsavel`/`diaRecolhe`/`dataEntrega`), marca como puxadas (grava
  vendedor, modelo, data de entrega e data de recolhimento — regra "hoje +5
  dias se depois de terça-feira, senão +3 dias"), monta os leads.
- **Backup em disco**: `PASTA_DESTINO/mês - MÊS - ano/dd-mm-aaaa/IMPORTAÇÃO
  GERAL GOALFY - data_hora.xlsx`, igual à estrutura do VBA.
- **Coluna `bonif`** quando algum item da fila é `AUTO BONIF`.
- **E-mail por vendedor**: agrupa os leads por (vendedor, produto, empresa) —
  igual ao `dictGrupos` do VBA —, gera uma planilha formatada por grupo e
  dispara um e-mail com o Outlook instalado, um por vendedor.
- **Impressão por vendedor**: mesmo agrupamento, gera a planilha por grupo
  com layout paisagem/A4 e manda direto para a impressora padrão.
- **Cascata de seleção**: antes de enviar e-mail ou imprimir, abre uma janela
  com checkboxes de cada grupo (vendedor + produto + empresa), para você
  escolher quais disparar — com "Selecionar Todos" / "Limpar Seleção".

## O que mudou de propósito

- **Envio de leads para o Goalfy passou a ser direto pelo webhook** (botão
  "ENVIAR PARA O GOALFY"), além do e-mail — não substitui mais o Outlook,
  os dois convivem agora.
- **Configuração em `config.json`**, mas **editável dentro do próprio app**:
  aba Configurações → sub-abas "Geral", "Matrizes" e "Vendedores", cada uma
  com Adicionar / Editar / Remover. Para a matriz, dá pra usar o botão
  "Procurar..." pra escolher o arquivo `.xlsx` direto do explorador.

## Requisitos para e-mail e impressão

Essas duas funções replicam exatamente o que a planilha original fazia:
usam o **Outlook** e a **impressora padrão do Windows** instalados na
máquina. Por isso:
- Só funcionam rodando o app **no Windows**, com o Outlook instalado (e
  de preferência aberto/com perfil configurado).
- Exigem a biblioteca `pywin32` (já está no `requirements.txt`, instalada
  automaticamente só em Windows).
- Se você preferir enviar e-mail por SMTP (sem depender do Outlook aberto),
  me avise que eu troco a implementação — é só um ajuste em
  `core.enviar_email_outlook`.

## Bug corrigido na portagem

O `app.py`/`teste2.py` originais usavam `pd.to_datetime(valor,
dayfirst=True)` para qualquer data, o que inverte dia e mês quando o valor
já vem em `AAAA-MM-DD` (ex: virava "2026-01-08" em vez de "2026-08-01").
Corrigido no `core.py`.

## Como rodar

```bash
python -m venv venv
# Windows:
venv\Scripts\activate
pip install -r requirements.txt
python app.py
```

## Como gerar o .exe

```bash
pip install pyinstaller
pyinstaller --noconfirm --onefile --windowed --icon=imp.ico --name "ImportadorGoalfy" app.py
```

O executável final fica em `dist/ImportadorGoalfy.exe`. Copie o `config.json`
para a mesma pasta do `.exe` (ele é lido/gravado ao lado do script — e agora
também é atualizado automaticamente quando você edita matrizes/vendedores
pela tela).

## Estrutura de arquivos

```
importador_goalfy/
├── app.py           # Interface gráfica (customtkinter)
├── core.py          # Toda a lógica de negócio (testável sem GUI)
├── config.json       # Configuração (matrizes, pasta destino, webhook, vendedores)
├── requirements.txt
└── README.md
```
