"""
Importador Goalfy - v3
Tela de seleção de matrizes (porte do UserForm frmSelecaoMatrizes) +
consolidação (porte do Módulo VBA) + envio direto ao webhook do Goalfy +
envio de e-mail por vendedor (Outlook) + impressão por vendedor +
cadastro de matrizes e vendedores direto no app.
"""

import os
import threading
import traceback
from tkinter import filedialog, messagebox, ttk

import customtkinter as ctk

import core
import gestores
import envio_goalfy
from autocomplete import CampoAutocomplete, chave as chave_campo

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

PH_PRODUTO = "SELECIONE O PRODUTO"
PH_MODELO = "SELECIONE O MODELO"
PH_VENDEDOR = "SELECIONE O VENDEDOR"

AZUL_ESCURO = "#003366"
AZUL = "#0066CC"
VERMELHO = "#B42828"
VERDE = "#1f8d4b"


class AppImportador(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Importador Goalfy")
        self.geometry("1040x720")
        self.minsize(960, 650)

        self.config_data = gestores.migrar_cadastros_gestores(core.load_config())
        self._gerando = False
        self._assinaturas_geradas = set()
        self._assinatura_em_geracao = None
        self._goalfy_em_execucao = False
        self._exportacao_pendente = None
        self.fila: list[core.ItemFila] = []
        self.ultimos_registros: list[dict] = []
        self.tem_auto_bonif = False
        self._ultimo_bloco = None
        self._modelos_cache = {}
        self._produto_modelos = None

        self.tabview = ctk.CTkTabview(self, width=1020, height=700)
        self.tabview.pack(padx=10, pady=10, fill="both", expand=True)

        self.tab_remessa = self.tabview.add("Nova Remessa")
        self.tab_config = self.tabview.add("Configurações")

        self._montar_tab_remessa()
        self._montar_tab_config()

    # ==================================================================
    # ABA: NOVA REMESSA
    # ==================================================================
    def _montar_tab_remessa(self):
        frame = self.tab_remessa

        ctk.CTkLabel(
            frame,
            text="Montar Remessa de Leads",
            font=ctk.CTkFont(size=22, weight="bold"),
        ).pack(pady=(15, 5))

        # --- linha de seleção ---
        linha = ctk.CTkFrame(frame, fg_color="transparent")
        linha.pack(pady=10, padx=20, fill="x")

        col1 = ctk.CTkFrame(linha, fg_color="transparent")
        col1.pack(side="left", padx=5, fill="x", expand=True)
        ctk.CTkLabel(col1, text="Produto").pack(anchor="w")
        self.cbo_produto = CampoAutocomplete(
            col1,
            values=core.listar_produtos_ativos(self.config_data),
            command=self._on_produto_change,
            width=180,
        )
        self.cbo_produto.set(PH_PRODUTO)
        self.cbo_produto.pack(fill="x")

        col2 = ctk.CTkFrame(linha, fg_color="transparent")
        col2.pack(side="left", padx=5, fill="x", expand=True)
        ctk.CTkLabel(col2, text="Modelo").pack(anchor="w")
        self.cbo_modelo = CampoAutocomplete(col2, values=[], width=180)
        self.cbo_modelo.set(PH_MODELO)
        self.cbo_modelo.pack(fill="x")

        col3 = ctk.CTkFrame(linha, fg_color="transparent")
        col3.pack(side="left", padx=5, fill="x", expand=True)
        ctk.CTkLabel(col3, text="Vendedor").pack(anchor="w")
        self.cbo_vendedor = CampoAutocomplete(
            col3, values=self._nomes_vendedores(), width=180
        )
        self.cbo_vendedor.set(PH_VENDEDOR)
        self.cbo_vendedor.pack(fill="x")

        col4 = ctk.CTkFrame(linha, fg_color="transparent")
        col4.pack(side="left", padx=5)
        ctk.CTkLabel(col4, text="Qtd").pack(anchor="w")
        self.txt_quantidade = ctk.CTkEntry(col4, width=70, placeholder_text="QTD")
        self.txt_quantidade.pack(fill="x")

        # --- botões adicionar/remover ---
        linha_botoes = ctk.CTkFrame(frame, fg_color="transparent")
        linha_botoes.pack(pady=5, padx=20, fill="x")

        ctk.CTkButton(
            linha_botoes,
            text="ADICIONAR À FILA",
            fg_color=AZUL_ESCURO,
            command=self._adicionar_item,
        ).pack(side="left", padx=5)
        ctk.CTkButton(
            linha_botoes,
            text="REMOVER ITEM SELECIONADO",
            fg_color=VERMELHO,
            command=self._remover_item,
        ).pack(side="left", padx=5)

        ctk.CTkButton(
            linha_botoes, text="VOLTAR PRODUTO / MODELO", fg_color="#555555",
            command=self._restaurar_ultimo_bloco,
        ).pack(side="left", padx=5)

        # --- tabela da fila ---
        tabela_frame = ctk.CTkFrame(frame)
        tabela_frame.pack(padx=20, pady=10, fill="both", expand=True)

        colunas = ("produto", "modelo", "vendedor", "quantidade")
        self.tree = ttk.Treeview(
            tabela_frame, columns=colunas, show="headings", height=6
        )
        for c, w in zip(colunas, (220, 160, 200, 90)):
            self.tree.heading(c, text=c.upper())
            self.tree.column(c, width=w, anchor="center")
        self.tree.pack(fill="both", expand=True, padx=5, pady=5)

        # --- botão gerar ---
        self.btn_gerar = ctk.CTkButton(
            frame,
            text="GERAR IMPORTAÇÃO",
            fg_color=AZUL,
            height=45,
            font=ctk.CTkFont(size=15, weight="bold"),
            command=self._gerar_importacao_thread,
        )
        self.btn_gerar.pack(padx=20, pady=(5, 5), fill="x")
        controles = ctk.CTkFrame(frame, fg_color="transparent")
        controles.pack(fill="x", padx=20, pady=3)
        self.travar_repeticao = ctk.BooleanVar(value=True)
        ctk.CTkSwitch(controles, text="TRAVAR IMPORTAÇÃO REPETIDA", variable=self.travar_repeticao).pack(side="left", padx=5)
        ctk.CTkButton(controles, text="NOVA REMESSA", command=self._nova_remessa, width=150).pack(side="right", padx=5)

        self.lbl_status_geracao = ctk.CTkLabel(
            frame, text="Monte a fila e clique em Gerar Importação.", text_color="gray"
        )
        self.lbl_status_geracao.pack(pady=(0, 5))

        # --- ações pós-geração: Goalfy / E-mail / Imprimir ---
        acoes_frame = ctk.CTkFrame(frame)
        acoes_frame.pack(padx=20, pady=(5, 15), fill="x")

        self.lbl_resumo_envio = ctk.CTkLabel(
            acoes_frame, text="Nenhuma importação gerada ainda."
        )
        self.lbl_resumo_envio.pack(pady=(10, 5))

        self.progress_bar = ctk.CTkProgressBar(acoes_frame)
        self.progress_bar.set(0)
        self.progress_bar.pack(padx=20, pady=5, fill="x")

        botoes_acao = ctk.CTkFrame(acoes_frame, fg_color="transparent")
        botoes_acao.pack(padx=20, pady=(0, 10), fill="x")
        botoes_acao.grid_columnconfigure((0, 1, 2), weight=1)

        self.btn_enviar_goalfy = ctk.CTkButton(
            botoes_acao,
            text="ENVIAR PARA O GOALFY",
            fg_color=VERDE,
            hover_color="#166536",
            state="disabled",
            command=self._enviar_goalfy_thread,
        )
        self.btn_enviar_goalfy.grid(row=0, column=0, padx=5, sticky="ew")

        self.btn_enviar_email = ctk.CTkButton(
            botoes_acao,
            text="ENVIAR E-MAIL P/ VENDEDORES",
            fg_color=AZUL_ESCURO,
            state="disabled",
            command=lambda: self._abrir_popup_selecao_vendedores("email"),
        )
        self.btn_enviar_email.grid(row=0, column=1, padx=5, sticky="ew")

        self.btn_imprimir = ctk.CTkButton(
            botoes_acao,
            text="IMPRIMIR P/ VENDEDORES",
            fg_color="#555555",
            state="disabled",
            command=lambda: self._abrir_popup_selecao_vendedores("imprimir"),
        )
        self.btn_imprimir.grid(row=0, column=2, padx=5, sticky="ew")

        self.btn_enviar_gestores = ctk.CTkButton(
            botoes_acao, text="ENVIAR E-MAIL P/ GESTORES", fg_color=AZUL_ESCURO,
            state="disabled", command=self._selecionar_gestores,
        )
        self.btn_enviar_gestores.grid(row=1, column=0, columnspan=3, padx=5, pady=(8, 0), sticky="ew")
        self.btn_reenviar_goalfy = ctk.CTkButton(
            botoes_acao, text="ENVIAR PLANILHA SALVA AO GOALFY", fg_color=VERDE,
            command=self._abrir_planilha_goalfy,
        )
        self.btn_reenviar_goalfy.grid(row=2, column=0, columnspan=3, padx=5, pady=(8, 0), sticky="ew")


    def _selecionar_gestores(self):
        if not self.ultimos_registros:
            return
        try:
            configuracoes = gestores.obter_configuracoes_gestores(self.config_data)
            produtos = self.ultimos_registros
            selecionados = gestores.abrir_interface_selecao_gestores(produtos, configuracoes, parent=self)
        except Exception as exc:
            messagebox.showerror("Gestores", f"Não foi possível ler a configuração: {exc}", parent=self)
            return
        if not selecionados:
            return
        registros = [dict(r) for r in self.ultimos_registros]
        self._set_botoes_pos_geracao(False)
        threading.Thread(target=self._executar_envio_gestores,
                         args=(selecionados, registros), daemon=True).start()

    def _executar_envio_gestores(self, selecionados, registros):
        import tempfile
        def progresso(i, total, ok):
            self.after(0, lambda i=i, total=total: self.lbl_resumo_envio.configure(text=f"Gestores: envio {i} de {total}..."))
            self.after(0, lambda i=i, total=total: self.progress_bar.set(i / total))
        try:
            produtos = {gestores._chave(r["PRODUTO"]) for r in selecionados}
            registros = [r for r in registros if gestores.produtos_do_lead(r).intersection(produtos)]
            with tempfile.TemporaryDirectory(prefix="relatorios_gestores_") as pasta:
                relatorios = gestores.gerar_relatorios_gerais(registros, selecionados, pasta)
                sucesso, falha, avisos = gestores.enviar_emails_gestores(selecionados, relatorios, progresso)
            mensagem = f"Enviados: {sucesso}\nFalhas/ignorados: {falha}"
            if avisos:
                mensagem += "\n\nDetalhes:\n" + "\n".join(avisos)
            self.after(0, lambda mensagem=mensagem: messagebox.showinfo("Envio para Gestores", mensagem, parent=self))
        except Exception as exc:
            mensagem = str(exc)
            self.after(0, lambda mensagem=mensagem: messagebox.showerror("Gestores", mensagem, parent=self))
        finally:
            self.after(0, lambda: self._set_botoes_pos_geracao(True))

    def _nomes_vendedores(self):
        return [v.get("nome", "") for v in self.config_data.get("vendedores", [])]

    def _on_produto_change(self, *_):
        produto = self.cbo_produto.get().strip()
        if not produto or produto == PH_PRODUTO:
            return
        self._atualizar_modelos(produto)

    def _atualizar_modelos(self, produto: str):
        caminho = core.obter_caminho_matriz(self.config_data, produto)
        if not caminho or not os.path.isfile(caminho):
            messagebox.showwarning(
                "Erro de Arquivo",
                f"Não foi possível localizar o arquivo da matriz para: {produto}",
            )
            self.cbo_modelo.configure(values=[])
            self.cbo_modelo.set(PH_MODELO)
            return
        try:
            stat = os.stat(caminho)
            chave_cache = (caminho, stat.st_size, stat.st_mtime_ns)
            if chave_cache not in self._modelos_cache:
                self._modelos_cache[chave_cache] = core.obter_modelos_disponiveis(caminho, None, produto)
            saldo = dict(self._modelos_cache[chave_cache])
            for item in self.fila:
                if item.produto.strip().upper() == produto.strip().upper():
                    modelo = item.modelo.strip().upper()
                    saldo[modelo] = saldo.get(modelo, 0) - item.quantidade
            saldo = {modelo:qtd for modelo,qtd in saldo.items() if qtd > 0}
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao ler a matriz:\n{e}")
            return

        self._produto_modelos = produto
        if not saldo:
            self.cbo_modelo.configure(values=["(NENHUM MODELO COM ESTOQUE)"])
            self.cbo_modelo.set("(NENHUM MODELO COM ESTOQUE)")
        else:
            valores = [f"{modelo} - {qtd}" for modelo, qtd in saldo.items()]
            self.cbo_modelo.configure(values=valores)
            self.cbo_modelo.set(PH_MODELO)

    def _restaurar_ultimo_bloco(self):
        if self._ultimo_bloco is None:
            messagebox.showinfo("Última seleção", "Adicione um bloco à fila primeiro.", parent=self)
            return
        produto, modelo = self._ultimo_bloco
        self.cbo_produto.set(produto)
        self._atualizar_modelos(produto)
        opcao = next((v for v in self.cbo_modelo.cget("values") if v.rsplit(" - ",1)[0] == modelo), None)
        if opcao:
            self.cbo_modelo.set(opcao)
        else:
            messagebox.showinfo("Modelo", "O modelo anterior não possui saldo disponível para outro bloco.", parent=self)

    def _adicionar_item(self):
        produto = self.cbo_produto.get().strip()
        modelo_raw = self.cbo_modelo.get().strip()
        vendedor = self.cbo_vendedor.get().strip()
        qtd_raw = self.txt_quantidade.get().strip()

        if not produto or produto == PH_PRODUTO:
            messagebox.showwarning("Aviso", "Selecione um Produto!")
            return
        if (
            not modelo_raw
            or modelo_raw == PH_MODELO
            or modelo_raw == "(NENHUM MODELO COM ESTOQUE)"
        ):
            messagebox.showwarning("Aviso", "Selecione um Modelo válido!")
            return
        if not vendedor or vendedor == PH_VENDEDOR:
            messagebox.showwarning("Aviso", "Selecione um Vendedor!")
            return
        if not qtd_raw.isdigit() or int(qtd_raw) <= 0:
            messagebox.showwarning("Aviso", "Informe uma quantidade válida!")
            return

        produto_cadastrado = next((v for v in core.listar_produtos_ativos(self.config_data) if chave_campo(v.strip()) == chave_campo(produto)), None)
        vendedor_cadastrado = next((v for v in self._nomes_vendedores() if chave_campo(v.strip()) == chave_campo(vendedor)), None)
        if produto_cadastrado is None or vendedor_cadastrado is None:
            messagebox.showwarning("Seleção", "Escolha um produto e um vendedor cadastrados nas sugestões.", parent=self)
            return
        produto, vendedor = produto_cadastrado, vendedor_cadastrado.strip()
        if self._produto_modelos != produto:
            self._atualizar_modelos(produto)
        opcoes_modelo = self.cbo_modelo.cget("values")
        correspondencia = next((v for v in opcoes_modelo if chave_campo(v) == chave_campo(modelo_raw)
                                or chave_campo(v.rsplit(" - ",1)[0]) == chave_campo(modelo_raw)), None)
        if correspondencia is None:
            messagebox.showwarning("Modelo", "Selecione um modelo disponível nas sugestões.", parent=self)
            return
        modelo_raw = correspondencia
        modelo = (
            modelo_raw.rsplit(" - ",1)[0].strip() if " - " in modelo_raw else modelo_raw
        )

        item = core.ItemFila(
            produto=produto, modelo=modelo, vendedor=vendedor, quantidade=int(qtd_raw)
        )
        self._ultimo_bloco = (produto, modelo)
        self.fila.append(item)
        self.tree.insert(
            "",
            "end",
            values=(item.produto, item.modelo, item.vendedor, item.quantidade),
        )

        self.cbo_produto.set(PH_PRODUTO)
        self.cbo_modelo.configure(values=[])
        self.cbo_modelo.set(PH_MODELO)
        self.cbo_vendedor.set(PH_VENDEDOR)
        self.txt_quantidade.delete(0, "end")

    def _remover_item(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Aviso", "Selecione um item da lista para remover.")
            return
        idx = self.tree.index(sel[0])
        self.tree.delete(sel[0])
        del self.fila[idx]

    # ------------------------------------------------------------------
    # Geração da importação
    # ------------------------------------------------------------------
    def _nova_remessa(self):
        if self._gerando or self._goalfy_em_execucao:
            return
        if self._exportacao_pendente is not None:
            messagebox.showwarning("Salvamento pendente", "Conclua o salvamento da entrega atual antes de iniciar outra remessa.", parent=self)
            return
        self.fila.clear()
        self.tree.delete(*self.tree.get_children())
        self._assinaturas_geradas.clear()
        self.ultimos_registros = []
        self._set_botoes_pos_geracao(False)
        self.lbl_status_geracao.configure(text="Monte a nova fila de entrega.", text_color="gray")
        self.lbl_resumo_envio.configure(text="Nenhuma importação gerada nesta remessa.")
        self.progress_bar.set(0)

    def _gerar_importacao_thread(self):
        import copy
        if self._gerando or getattr(self, "_goalfy_em_execucao", False):
            return
        if not self.fila and self._exportacao_pendente is None:
            messagebox.showwarning("Aviso", "A fila está vazia. Adicione pelo menos um item.", parent=self)
            return
        assinatura = tuple((i.produto.strip(), i.modelo.strip(), i.vendedor.strip(), i.quantidade) for i in self.fila)
        if self._exportacao_pendente is None and self.travar_repeticao.get() and assinatura in self._assinaturas_geradas:
            messagebox.showwarning("Importação repetida bloqueada",
                                   "Esta fila já gerou uma importação. Não é possível gerar duas vezes a mesma entrega. Use NOVA REMESSA para iniciar outra entrega.", parent=self)
            return
        if self._exportacao_pendente is None:
            self._assinatura_em_geracao = assinatura
        nova_config = copy.deepcopy(self.config_data)
        nova_config["pasta_destino"] = self.entry_pasta_destino.get().strip()
        try:
            core.validar_pasta_destino(nova_config)
            core.save_config(nova_config)
        except Exception as exc:
            messagebox.showerror("Pasta de importação", str(exc), parent=self)
            return
        self.config_data = nova_config
        self._gerando = True
        self.btn_gerar.configure(state="disabled")
        self._set_botoes_pos_geracao(False)
        self.lbl_status_geracao.configure(text="Processando importação, aguarde...", text_color="gray")
        threading.Thread(target=self._gerar_importacao,
                         args=(copy.deepcopy(nova_config), copy.deepcopy(self.fila)), daemon=True).start()

    def _gerar_importacao(self, config, fila):
        try:
            if self._exportacao_pendente is None:
                def progresso(texto):
                    self.after(0, lambda texto=texto: self.lbl_status_geracao.configure(text=texto))
                registros, tem_bonif, avisos = core.consolidar_remessa(config, fila, progress_callback=progresso)
                if not registros:
                    self.after(0, self._finalizar_sem_leads)
                    return
                # Se a exportação falhar, repetir a geração salva esta mesma entrega.
                self._exportacao_pendente = (registros, tem_bonif, avisos)
            registros, tem_bonif, avisos = self._exportacao_pendente
            self.after(0, lambda: self.lbl_status_geracao.configure(text="Gravando arquivo de importação..."))
            caminho = core.salvar_planilha_importacao(config, registros, tem_bonif)
            self.after(0, lambda: self._finalizar_importacao(registros, tem_bonif, avisos, caminho))
        except Exception:
            erro = traceback.format_exc()
            self.after(0, lambda erro=erro: self._falha_importacao(erro))

    def _finalizar_sem_leads(self):
        self._gerando = False
        self.btn_gerar.configure(state="normal")
        self._set_botoes_pos_geracao(bool(self.ultimos_registros))
        self.lbl_status_geracao.configure(text="Nenhum lead elegível localizado.", text_color="orange")
        messagebox.showwarning("Aviso", "Nenhum lead elegível localizado para os critérios informados.", parent=self)

    def _finalizar_importacao(self, registros, tem_bonif, avisos, caminho):
        self._assinaturas_geradas.add(self._assinatura_em_geracao)
        self.ultimos_registros = registros
        self.tem_auto_bonif = tem_bonif
        self._exportacao_pendente = None
        self._gerando = False
        self.btn_gerar.configure(state="normal", text="GERAR IMPORTAÇÃO")
        self.lbl_status_geracao.configure(text=f"{len(registros)} leads importados.", text_color=VERDE)
        self.lbl_resumo_envio.configure(text=f"{len(registros)} leads prontos. Escolha uma ação abaixo.")
        self._set_botoes_pos_geracao(True)
        self.progress_bar.set(0)
        mensagem = f"{len(registros)} leads importados. Arquivo salvo em:\n{caminho}"
        if avisos:
            mensagem += "\n\nAvisos:\n" + "\n".join(avisos)
        messagebox.showinfo("Importação Concluída", mensagem, parent=self)

    def _falha_importacao(self, erro):
        self._gerando = False
        self.btn_gerar.configure(state="normal")
        if self._exportacao_pendente is not None:
            self.btn_gerar.configure(text="REPETIR SALVAMENTO DA IMPORTAÇÃO")
            erro = ("Os leads desta entrega foram preservados nesta sessão. "
                    "Corrija a pasta de saída e clique em REPETIR SALVAMENTO DA IMPORTAÇÃO. "
                    "Isso não selecionará novos leads. Mantenha o app aberto até concluir.\n\n" + erro)
        self.lbl_status_geracao.configure(text="Erro ao gerar importação.", text_color="red")
        self._set_botoes_pos_geracao(bool(self.ultimos_registros))
        messagebox.showerror("Erro", erro, parent=self)

    def _set_botoes_pos_geracao(self, habilitado: bool):
        estado = "normal" if habilitado else "disabled"
        self.btn_enviar_goalfy.configure(state=estado)
        self.btn_enviar_email.configure(state=estado)
        self.btn_imprimir.configure(state=estado)
        self.btn_enviar_gestores.configure(state=estado)

    # ------------------------------------------------------------------
    # Envio para o Goalfy
    # ------------------------------------------------------------------
    def _enviar_goalfy_thread(self):
        if not self.ultimos_registros or self._goalfy_em_execucao or self._gerando:
            return
        try:
            if not envio_goalfy.indices_pendentes(self.ultimos_registros):
                messagebox.showwarning("Envio repetido bloqueado", "Esta importação já foi enviada ao Goalfy. Não é possível enviar a mesma planilha duas vezes.", parent=self)
                return
        except Exception as exc:
            messagebox.showerror("Histórico de envios", str(exc), parent=self)
            return
        self._goalfy_em_execucao = True
        threading.Thread(target=self._enviar_goalfy, daemon=True).start()

    def _enviar_goalfy(self):
        self.after(0, lambda: self.btn_reenviar_goalfy.configure(state="disabled"))
        self._set_botoes_pos_geracao(habilitado=False)

        def progresso(i, total, ok):
            self.progress_bar.set(i / total)
            self.lbl_resumo_envio.configure(
                text=f"Enviando {i} de {total} leads ao Goalfy..."
            )

        try:
            sucessos, falhas, erros = envio_goalfy.enviar_protegido(
                self.config_data, self.ultimos_registros, progress_callback=progresso
            )
            self.lbl_resumo_envio.configure(
                text=f"✅ Goalfy: enviados {sucessos} | falhas {falhas}"
            )
            msg = f"Importação para o Goalfy finalizada!\n\nEnviados: {sucessos}\nFalhas: {falhas}"
            if erros:
                msg += "\n\nDetalhes das falhas:\n" + "\n".join(erros[:10])
            messagebox.showinfo("Envio Concluído", msg)
        except Exception:
            erro = traceback.format_exc()
            messagebox.showerror(
                "Erro", f"Ocorreu um erro ao enviar para o Goalfy:\n\n{erro}"
            )
        finally:
            self.after(0, self._finalizar_reenvio_goalfy)

    def _abrir_planilha_goalfy(self):
        if self._goalfy_em_execucao or self._gerando:
            return
        caminho = filedialog.askopenfilename(parent=self, title="Selecione uma importação salva",
                                             filetypes=[("Planilhas de importação", "*.xlsx")])
        if not caminho:
            return
        try:
            registros = core.ler_planilha_importacao(caminho)
            pendentes = list(range(len(registros)))
        except Exception as exc:
            messagebox.showerror("Planilha de importação", str(exc), parent=self)
            return
        popup = ctk.CTkToplevel(self)
        popup.title("Enviar importação salva ao Goalfy")
        popup.geometry("900x570")
        popup.transient(self)
        popup.grab_set()
        ctk.CTkLabel(popup, text="Enviar planilha salva ao Goalfy", font=ctk.CTkFont(size=20, weight="bold")).pack(anchor="w", padx=20, pady=(15, 5))
        ctk.CTkLabel(popup, text=f"{os.path.basename(caminho)} • {len(registros)} leads", wraplength=850).pack(anchor="w", padx=20)
        ctk.CTkLabel(popup, text="Selecione os leads que deseja reenviar ao Goalfy.", wraplength=850).pack(anchor="w", padx=20, pady=5)
        frame = ctk.CTkFrame(popup)
        frame.pack(fill="both", expand=True, padx=20, pady=5)
        colunas = ("linha", "nome", "vendedor", "produto", "empresa")
        tree = ttk.Treeview(frame, columns=colunas, show="headings", selectmode="extended")
        for c, w in zip(colunas, (55, 250, 180, 130, 100)):
            tree.heading(c, text=c.upper())
            tree.column(c, width=w)
        barra = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=barra.set)
        barra.pack(side="right", fill="y")
        tree.pack(fill="both", expand=True)
        for i in pendentes:
            r = registros[i]
            tree.insert("", "end", iid=str(i), values=(r["_linha_origem"], r["nome"], r["responsavel"], r["produto"], r["empresa"]))
        tree.selection_set(tree.get_children())
        botoes = ctk.CTkFrame(popup, fg_color="transparent")
        botoes.pack(fill="x", padx=20, pady=10)
        ctk.CTkButton(botoes, text="Selecionar Todos", command=lambda: tree.selection_set(tree.get_children())).pack(side="left", padx=5)
        ctk.CTkButton(botoes, text="Limpar Seleção", command=lambda: tree.selection_remove(*tree.selection()), fg_color="#555555").pack(side="left", padx=5)
        def confirmar():
            import copy
            indices = [int(i) for i in tree.selection()]
            selecionados = [registros[i] for i in indices]
            if not selecionados:
                messagebox.showwarning("Seleção", "Selecione pelo menos um lead.", parent=popup)
                return
            popup.destroy()
            self._goalfy_em_execucao = True
            self.btn_reenviar_goalfy.configure(state="disabled")
            self._set_botoes_pos_geracao(False)
            threading.Thread(target=self._executar_reenvio_goalfy,
                             args=(copy.deepcopy(self.config_data), registros, indices), daemon=True).start()
        rodape = ctk.CTkFrame(popup, fg_color="transparent")
        rodape.pack(fill="x", padx=20, pady=(0, 15))
        ctk.CTkButton(rodape, text="CONFIRMAR ENVIO", command=confirmar, fg_color=VERDE).pack(side="left", fill="x", expand=True, padx=5)
        ctk.CTkButton(rodape, text="CANCELAR", command=popup.destroy, fg_color=VERMELHO).pack(side="right", padx=5)

    def _executar_reenvio_goalfy(self, config, registros, indices=None):
        def progresso(i, total, ok):
            self.after(0, lambda i=i, total=total: self.progress_bar.set(i / total))
            self.after(0, lambda i=i, total=total: self.lbl_resumo_envio.configure(text=f"Planilha salva: enviando {i} de {total}..."))
        try:
            selecionados = registros if indices is None else [registros[i] for i in indices]
            sucesso, falha, erros = core.enviar_para_goalfy(config, selecionados, progress_callback=progresso)
            mensagem = f"Enviados: {sucesso}\nFalhas: {falha}"
            if erros:
                mensagem += "\n\n" + "\n".join(erros)
            self.after(0, lambda mensagem=mensagem: messagebox.showinfo("Envio da planilha salva", mensagem, parent=self))
        except Exception as exc:
            self.after(0, lambda mensagem=str(exc): messagebox.showerror("Goalfy", mensagem, parent=self))
        finally:
            self.after(0, self._finalizar_reenvio_goalfy)

    def _finalizar_reenvio_goalfy(self):
        self._goalfy_em_execucao = False
        self.btn_reenviar_goalfy.configure(state="normal")
        self._set_botoes_pos_geracao(bool(self.ultimos_registros))

    # ------------------------------------------------------------------
    # Cascata de seleção de vendedores (usada por e-mail e impressão)
    # ------------------------------------------------------------------
    def _abrir_popup_selecao_vendedores(self, acao: str):
        if not self.ultimos_registros:
            messagebox.showinfo("Aviso", "Gere uma importação primeiro.")
            return

        grupos = core.agrupar_leads(self.ultimos_registros)
        if not grupos:
            messagebox.showinfo(
                "Aviso", "Nenhum lead com vendedor responsável foi encontrado."
            )
            return

        popup = ctk.CTkToplevel(self)
        popup.title("Selecionar Vendedores")
        popup.geometry("520x520")
        popup.transient(self)
        popup.grab_set()

        titulo_acao = "enviar por e-mail" if acao == "email" else "imprimir"
        ctk.CTkLabel(
            popup,
            text=f"Selecione os grupos que deseja {titulo_acao}:",
            font=ctk.CTkFont(weight="bold"),
            wraplength=480,
        ).pack(pady=(15, 5), padx=15, anchor="w")

        scroll = ctk.CTkScrollableFrame(popup, width=480, height=340)
        scroll.pack(padx=15, pady=5, fill="both", expand=True)

        vars_checkbox: dict[tuple, ctk.BooleanVar] = {}
        for chave, regs in grupos.items():
            vend, prod, emp = chave
            texto = f"{vend} — {core.titulo_grupo(chave)}  ({len(regs)} leads)"
            if acao == "email" and not core.obter_email_vendedor(
                self.config_data, vend
            ):
                texto += "  [SEM E-MAIL CADASTRADO]"
            var = ctk.BooleanVar(value=True)
            ctk.CTkCheckBox(scroll, text=texto, variable=var).pack(
                anchor="w", pady=3, padx=5
            )
            vars_checkbox[chave] = var

        botoes = ctk.CTkFrame(popup, fg_color="transparent")
        botoes.pack(pady=10, fill="x", padx=15)

        def marcar_todos(valor):
            for v in vars_checkbox.values():
                v.set(valor)

        ctk.CTkButton(
            botoes,
            text="Selecionar Todos",
            width=140,
            command=lambda: marcar_todos(True),
        ).pack(side="left", padx=5)
        ctk.CTkButton(
            botoes,
            text="Limpar Seleção",
            width=140,
            command=lambda: marcar_todos(False),
        ).pack(side="left", padx=5)

        def confirmar():
            selecionados = {
                chave: grupos[chave] for chave, v in vars_checkbox.items() if v.get()
            }
            popup.destroy()
            if not selecionados:
                return
            if acao == "email":
                threading.Thread(
                    target=self._executar_envio_email, args=(selecionados,), daemon=True
                ).start()
            else:
                threading.Thread(
                    target=self._executar_impressao, args=(selecionados,), daemon=True
                ).start()

        ctk.CTkButton(
            popup, text="CONFIRMAR", fg_color=AZUL, height=38, command=confirmar
        ).pack(pady=(0, 15), padx=15, fill="x")

    def _executar_envio_email(self, selecionados: dict):
        self._set_botoes_pos_geracao(habilitado=False)

        def progresso(i, total, ok):
            self.progress_bar.set(i / total)
            self.lbl_resumo_envio.configure(text=f"Enviando e-mail {i} de {total}...")

        try:
            sucesso, falha, avisos = core.enviar_emails_para_grupos(
                self.config_data, selecionados, progress_callback=progresso
            )
            self.lbl_resumo_envio.configure(
                text=f"✅ E-mails: enviados {sucesso} | falhas {falha}"
            )
            msg = (
                f"Envio de e-mails finalizado!\n\nEnviados: {sucesso}\nFalhas: {falha}"
            )
            if avisos:
                msg += "\n\nDetalhes:\n" + "\n".join(avisos[:10])
            messagebox.showinfo("Envio de E-mails", msg)
        except Exception:
            erro = traceback.format_exc()
            messagebox.showerror(
                "Erro", f"Ocorreu um erro ao enviar os e-mails:\n\n{erro}"
            )
        finally:
            self._set_botoes_pos_geracao(habilitado=True)

    def _executar_impressao(self, selecionados: dict):
        self._set_botoes_pos_geracao(habilitado=False)

        def progresso(i, total, ok):
            self.progress_bar.set(i / total)
            self.lbl_resumo_envio.configure(
                text=f"Enviando para impressão {i} de {total}..."
            )

        try:
            sucesso, falha, avisos = core.imprimir_grupos(
                selecionados, progress_callback=progresso
            )
            self.lbl_resumo_envio.configure(
                text=f"✅ Impressão: enviados {sucesso} | falhas {falha}"
            )
            msg = f"Impressão finalizada!\n\nDocumentos enviados: {sucesso}\nFalhas: {falha}"
            if avisos:
                msg += "\n\nDetalhes:\n" + "\n".join(avisos[:10])
            messagebox.showinfo("Impressão", msg)
        except Exception:
            erro = traceback.format_exc()
            messagebox.showerror("Erro", f"Ocorreu um erro ao imprimir:\n\n{erro}")
        finally:
            self._set_botoes_pos_geracao(habilitado=True)

    # ==================================================================
    # ABA: CONFIGURAÇÕES
    # ==================================================================
    def _montar_tab_config(self):
        frame = self.tab_config

        sub_tabs = ctk.CTkTabview(frame, width=1000, height=660)
        sub_tabs.pack(padx=10, pady=10, fill="both", expand=True)

        tab_geral = sub_tabs.add("Geral")
        tab_matrizes = sub_tabs.add("Matrizes")
        tab_vendedores = sub_tabs.add("Vendedores")
        tab_gestores = sub_tabs.add("Gestores")

        self._montar_subtab_geral(tab_geral)
        self._montar_subtab_matrizes(tab_matrizes)
        self._montar_subtab_vendedores(tab_vendedores)
        self._montar_subtab_gestores(tab_gestores)

    # --- Geral ---
    def _montar_subtab_geral(self, frame):
        geral = ctk.CTkFrame(frame)
        geral.pack(padx=10, pady=15, fill="x")

        ctk.CTkLabel(geral, text="Pasta destino (backup das planilhas):").pack(
            anchor="w", padx=10, pady=(10, 0)
        )
        self.entry_pasta_destino = ctk.CTkEntry(geral)
        self.entry_pasta_destino.insert(0, self.config_data.get("pasta_destino", ""))
        self.entry_pasta_destino.pack(fill="x", padx=10, pady=(0, 10))
        def escolher_pasta():
            pasta = filedialog.askdirectory(parent=self, title="Selecione a pasta de importação")
            if pasta:
                self.entry_pasta_destino.delete(0, "end")
                self.entry_pasta_destino.insert(0, pasta)
        ctk.CTkButton(geral, text="Selecionar pasta de importação", command=escolher_pasta).pack(padx=10, pady=(0, 10))


        ctk.CTkLabel(geral, text="URL do Webhook Goalfy:").pack(
            anchor="w", padx=10, pady=(0, 0)
        )
        self.entry_webhook = ctk.CTkEntry(geral)
        self.entry_webhook.insert(0, self.config_data.get("webhook_url", ""))
        self.entry_webhook.pack(fill="x", padx=10, pady=(0, 10))

        ctk.CTkButton(
            geral, text="Salvar Configurações Gerais", command=self._salvar_config_geral
        ).pack(padx=10, pady=(0, 10))

        ctk.CTkLabel(
            frame,
            text="O envio de e-mail usa o Outlook instalado na máquina (igual à planilha original) "
            "e a impressão usa a impressora padrão do Windows — ambos só funcionam rodando o "
            "app em um Windows com Outlook/impressora configurados.",
            text_color="gray",
            wraplength=920,
            justify="left",
        ).pack(padx=10, pady=(10, 10), anchor="w")

    def _salvar_config_geral(self):
        import copy
        nova = copy.deepcopy(self.config_data)
        nova["pasta_destino"] = self.entry_pasta_destino.get().strip()
        nova["webhook_url"] = self.entry_webhook.get().strip()
        try:
            core.validar_pasta_destino(nova)
            core.save_config(nova)
        except Exception as exc:
            messagebox.showerror("Configurações", str(exc), parent=self)
            return
        self.config_data = nova
        messagebox.showinfo("Configurações", "Configurações gerais salvas com sucesso!", parent=self)

    # --- Matrizes ---
    def _montar_subtab_matrizes(self, frame):
        ctk.CTkLabel(
            frame, text="Matrizes cadastradas", font=ctk.CTkFont(size=16, weight="bold")
        ).pack(pady=(15, 5), anchor="w", padx=10)

        matrizes_frame = ctk.CTkFrame(frame)
        matrizes_frame.pack(padx=10, pady=5, fill="both", expand=True)

        cols = ("id", "exibicao", "ativado", "caminho")
        self.tree_matrizes = ttk.Treeview(
            matrizes_frame, columns=cols, show="headings", height=10
        )
        for c, w in zip(cols, (90, 130, 80, 560)):
            self.tree_matrizes.heading(c, text=c.upper())
            self.tree_matrizes.column(c, width=w, anchor="w")
        self.tree_matrizes.pack(fill="both", expand=True, padx=5, pady=5)
        self._recarregar_tabela_matrizes()

        botoes = ctk.CTkFrame(frame, fg_color="transparent")
        botoes.pack(padx=10, pady=(5, 15), fill="x")
        ctk.CTkButton(
            botoes,
            text="Adicionar Matriz",
            fg_color=AZUL_ESCURO,
            command=lambda: self._abrir_dialog_matriz(),
        ).pack(side="left", padx=5)
        ctk.CTkButton(
            botoes,
            text="Editar Selecionada",
            fg_color=AZUL,
            command=self._editar_matriz_selecionada,
        ).pack(side="left", padx=5)
        ctk.CTkButton(
            botoes,
            text="Remover Selecionada",
            fg_color=VERMELHO,
            command=self._remover_matriz_selecionada,
        ).pack(side="left", padx=5)

    def _recarregar_tabela_matrizes(self):
        for item in self.tree_matrizes.get_children():
            self.tree_matrizes.delete(item)
        for m in self.config_data.get("matrizes", []):
            self.tree_matrizes.insert(
                "",
                "end",
                values=(
                    m.get("id", ""),
                    m.get("exibicao", ""),
                    "SIM" if m.get("ativado") else "NÃO",
                    m.get("caminho", ""),
                ),
            )

    def _matriz_selecionada(self):
        sel = self.tree_matrizes.selection()
        if not sel:
            messagebox.showinfo("Aviso", "Selecione uma matriz na lista.")
            return None
        idx = self.tree_matrizes.index(sel[0])
        return self.config_data["matrizes"][idx]

    def _editar_matriz_selecionada(self):
        matriz = self._matriz_selecionada()
        if matriz:
            self._abrir_dialog_matriz(matriz)

    def _remover_matriz_selecionada(self):
        matriz = self._matriz_selecionada()
        if not matriz:
            return
        if messagebox.askyesno(
            "Confirmar", f"Remover a matriz '{matriz.get('exibicao')}'?"
        ):
            core.remover_matriz(self.config_data, matriz.get("id"))
            core.save_config(self.config_data)
            self._recarregar_tabela_matrizes()
            self.cbo_produto.configure(
                values=core.listar_produtos_ativos(self.config_data)
            )

    def _abrir_dialog_matriz(self, matriz: dict | None = None):
        popup = ctk.CTkToplevel(self)
        popup.title("Editar Matriz" if matriz else "Nova Matriz")
        popup.geometry("600x360")
        popup.transient(self)
        popup.grab_set()

        ctk.CTkLabel(popup, text="ID (identificador único, ex: MATRIZ_19)").pack(
            anchor="w", padx=15, pady=(15, 0)
        )
        entry_id = ctk.CTkEntry(popup)
        entry_id.insert(
            0,
            matriz.get("id", "")
            if matriz
            else f"MATRIZ_{len(self.config_data.get('matrizes', [])) + 1}",
        )
        if matriz:
            entry_id.configure(state="disabled")
        entry_id.pack(fill="x", padx=15)

        ctk.CTkLabel(popup, text="Nome de exibição (aparece no combo de Produto)").pack(
            anchor="w", padx=15, pady=(10, 0)
        )
        entry_exib = ctk.CTkEntry(popup)
        entry_exib.insert(0, matriz.get("exibicao", "") if matriz else "")
        entry_exib.pack(fill="x", padx=15)

        ctk.CTkLabel(popup, text="Caminho do arquivo da matriz (.xlsx)").pack(
            anchor="w", padx=15, pady=(10, 0)
        )
        frame_caminho = ctk.CTkFrame(popup, fg_color="transparent")
        frame_caminho.pack(fill="x", padx=15)
        entry_caminho = ctk.CTkEntry(frame_caminho)
        entry_caminho.insert(0, matriz.get("caminho", "") if matriz else "")
        entry_caminho.pack(side="left", fill="x", expand=True)

        def escolher_arquivo():
            caminho = filedialog.askopenfilename(
                title="Selecione a planilha da matriz",
                filetypes=[("Excel", "*.xlsx *.xlsm"), ("Todos os arquivos", "*.*")],
            )
            if caminho:
                entry_caminho.delete(0, "end")
                entry_caminho.insert(0, caminho)

        ctk.CTkButton(
            frame_caminho, text="Procurar...", width=90, command=escolher_arquivo
        ).pack(side="left", padx=(5, 0))

        var_ativado = ctk.BooleanVar(
            value=matriz.get("ativado", True) if matriz else True
        )
        ctk.CTkCheckBox(
            popup,
            text="Ativado (aparece na tela de Nova Remessa)",
            variable=var_ativado,
        ).pack(anchor="w", padx=15, pady=15)

        def salvar():
            novo_id = entry_id.get().strip()
            novo_exib = entry_exib.get().strip()
            novo_caminho = entry_caminho.get().strip()
            if not novo_id or not novo_exib or not novo_caminho:
                messagebox.showwarning(
                    "Aviso", "Preencha ID, Nome de exibição e Caminho."
                )
                return

            if matriz is None:
                if any(
                    m.get("id") == novo_id for m in self.config_data.get("matrizes", [])
                ):
                    messagebox.showwarning(
                        "Aviso", f"Já existe uma matriz com o ID '{novo_id}'."
                    )
                    return
                core.adicionar_matriz(
                    self.config_data,
                    novo_id,
                    novo_exib,
                    novo_caminho,
                    var_ativado.get(),
                )
            else:
                matriz["exibicao"] = novo_exib
                matriz["caminho"] = novo_caminho
                matriz["ativado"] = var_ativado.get()

            core.save_config(self.config_data)
            self._recarregar_tabela_matrizes()
            self.cbo_produto.configure(
                values=core.listar_produtos_ativos(self.config_data)
            )
            popup.destroy()

        ctk.CTkButton(
            popup, text="SALVAR", fg_color=AZUL, height=38, command=salvar
        ).pack(pady=10, padx=15, fill="x")

    # --- Gestores e subgestores ---
    def _montar_subtab_gestores(self, frame):
        ctk.CTkLabel(frame, text="Gestores e subgestores cadastrados",
                     font=ctk.CTkFont(size=16, weight="bold")).pack(pady=(15, 5), anchor="w", padx=10)
        ctk.CTkLabel(frame, text="Cadastre cada pessoa uma vez e associe seus produtos em Editar Selecionado.",
                     wraplength=850).pack(anchor="w", padx=10)
        tabela = ctk.CTkFrame(frame)
        tabela.pack(padx=10, pady=5, fill="both", expand=True)
        colunas = ("PESSOA", "EMAIL", "PRODUTOS")
        self.tree_gestores = ttk.Treeview(tabela, columns=colunas, show="headings", selectmode="browse")
        for c, largura in zip(colunas, (220, 250, 400)):
            self.tree_gestores.heading(c, text="E-MAIL" if c == "EMAIL" else c)
            self.tree_gestores.column(c, width=largura, minwidth=50, anchor="w")
        barra = ttk.Scrollbar(tabela, orient="vertical", command=self.tree_gestores.yview)
        self.tree_gestores.configure(yscrollcommand=barra.set)
        barra.pack(side="right", fill="y")
        self.tree_gestores.pack(fill="both", expand=True, padx=5, pady=5)
        self._recarregar_tabela_gestores()
        botoes = ctk.CTkFrame(frame, fg_color="transparent")
        botoes.pack(padx=10, pady=(5, 15), fill="x")
        ctk.CTkButton(botoes, text="Adicionar Gestor", fg_color=AZUL_ESCURO,
                     command=self._abrir_dialog_gestor).pack(side="left", padx=5)
        ctk.CTkButton(botoes, text="Editar Selecionado", fg_color=AZUL,
                     command=self._editar_gestor_selecionado).pack(side="left", padx=5)
        ctk.CTkButton(botoes, text="Remover Selecionado", fg_color=VERMELHO,
                     command=self._remover_gestor_selecionado).pack(side="left", padx=5)

    def _recarregar_tabela_gestores(self):
        self.tree_gestores.delete(*self.tree_gestores.get_children())
        for i, r in enumerate(self.config_data.get("gestores", [])):
            self.tree_gestores.insert("", "end", iid=str(i), values=(r.get("PESSOA", ""), r.get("EMAIL", ""), ", ".join(v["PRODUTO"] for v in r.get("PRODUTOS", []))))

    def _gestor_selecionado(self):
        selecionados = self.tree_gestores.selection()
        if not selecionados:
            messagebox.showinfo("Aviso", "Selecione um gestor ou subgestor na lista.", parent=self)
            return None
        return int(selecionados[0])

    def _editar_gestor_selecionado(self):
        indice = self._gestor_selecionado()
        if indice is not None:
            self._abrir_dialog_gestor(indice)

    def _remover_gestor_selecionado(self):
        indice = self._gestor_selecionado()
        if indice is None:
            return
        r = self.config_data["gestores"][indice]
        if messagebox.askyesno("Confirmar", f"Remover {r['PESSOA']} e todos os seus vínculos com produtos?", parent=self):
            import copy
            nova_config = copy.deepcopy(self.config_data)
            gestores.excluir_cadastro_gestor(nova_config, indice)
            try:
                core.save_config(nova_config)
            except Exception as exc:
                messagebox.showerror("Gestores", f"Não foi possível salvar: {exc}", parent=self)
                return
            self.config_data = nova_config
            self._recarregar_tabela_gestores()

    def _abrir_dialog_gestor(self, indice=None):
        import copy
        registro = self.config_data.get("gestores", [])[indice] if indice is not None else {}
        vinculos = copy.deepcopy(registro.get("PRODUTOS", []))
        popup = ctk.CTkToplevel(self)
        popup.title("Editar Gestor/Subgestor" if indice is not None else "Novo Gestor/Subgestor")
        popup.geometry("850x650")
        popup.transient(self)
        popup.grab_set()
        entradas = {}
        for coluna, nome in (("PESSOA", "Nome"), ("EMAIL", "E-mail")):
            ctk.CTkLabel(popup, text=nome).pack(anchor="w", padx=15, pady=(8, 0))
            campo = ctk.CTkEntry(popup)
            campo.insert(0, registro.get(coluna, ""))
            campo.pack(fill="x", padx=15)
            entradas[coluna] = campo
        ctk.CTkLabel(popup, text="Produtos atribuídos — selecione uma linha para editar suas opções.").pack(anchor="w", padx=15, pady=(10, 0))
        tree = ttk.Treeview(popup, columns=("PRODUTO", "TIPO", "PADRÃO", "ATIVO"), show="headings", height=8, selectmode="browse")
        for c in tree["columns"]:
            tree.heading(c, text=c)
            tree.column(c, width=180)
        tree.pack(fill="both", expand=True, padx=15, pady=5)
        produtos = list(dict.fromkeys(
            [m.get("exibicao", "") for m in self.config_data.get("matrizes", []) if m.get("exibicao")]
            + [str(r.get("produto", "")).strip() for r in self.ultimos_registros if r.get("produto")]
            + [v["PRODUTO"] for pessoa in self.config_data.get("gestores", []) for v in pessoa.get("PRODUTOS", [])]
        ))
        campos = {}
        controles = ctk.CTkFrame(popup)
        controles.pack(fill="x", padx=15, pady=5)
        opcoes = {"PRODUTO": produtos, "TIPO": ["GESTOR", "SUBGESTOR"], "PADRÃO": ["SIM", "NÃO"], "ATIVO": ["SIM", "NÃO"]}
        for i, c in enumerate(opcoes):
            controles.grid_columnconfigure(i, weight=1)
            ctk.CTkLabel(controles, text=c).grid(row=0, column=i, padx=5)
            campo = ctk.CTkComboBox(controles, values=opcoes[c], width=175,
                                   state="normal" if c == "PRODUTO" else "readonly")
            campo.set({"PRODUTO": "", "TIPO": "GESTOR", "PADRÃO": "NÃO", "ATIVO": "SIM"}[c])
            campo.grid(row=1, column=i, padx=5, pady=5, sticky="ew")
            campos[c] = campo
        editando = [None]
        def recarregar():
            tree.delete(*tree.get_children())
            for i, r in enumerate(vinculos):
                tree.insert("", "end", iid=str(i), values=[r[c] for c in tree["columns"]])
        def selecionar(event):
            if tree.selection():
                editando[0] = int(tree.selection()[0])
                for c, campo in campos.items():
                    campo.set(vinculos[editando[0]][c])
        tree.bind("<<TreeviewSelect>>", selecionar)
        def novo_produto():
            editando[0] = None
            tree.selection_remove(*tree.selection())
            campos["PRODUTO"].set("")
        def atribuir():
            r = {c: campo.get().strip() for c, campo in campos.items()}
            if not r["PRODUTO"]:
                messagebox.showwarning("Produto", "Selecione ou digite um produto.", parent=popup)
                return
            if any(i != editando[0] and gestores._chave(v["PRODUTO"]) == gestores._chave(r["PRODUTO"]) for i, v in enumerate(vinculos)):
                messagebox.showwarning("Produto", "Produto já atribuído. Selecione sua linha para editar.", parent=popup)
                return
            if editando[0] is None:
                vinculos.append(r)
            else:
                vinculos[editando[0]] = r
            novo_produto()
            recarregar()
        def remover_produto():
            if tree.selection():
                del vinculos[int(tree.selection()[0])]
                novo_produto()
                recarregar()
        botoes_produto = ctk.CTkFrame(popup, fg_color="transparent")
        botoes_produto.pack(fill="x", padx=15, pady=5)
        ctk.CTkButton(botoes_produto, text="Novo produto", command=novo_produto).pack(side="left", padx=5)
        ctk.CTkButton(botoes_produto, text="Atribuir / Atualizar produto", command=atribuir, width=210).pack(side="left", padx=5)
        ctk.CTkButton(botoes_produto, text="Remover produto", command=remover_produto, fg_color=VERMELHO).pack(side="left", padx=5)
        def salvar():
            # Aplica também o produto preenchido, evitando perder uma edição pendente.
            if campos["PRODUTO"].get().strip():
                r = {c: campo.get().strip() for c, campo in campos.items()}
                if editando[0] is not None:
                    vinculos[editando[0]] = r
                elif not any(gestores._chave(v["PRODUTO"]) == gestores._chave(r["PRODUTO"]) for v in vinculos):
                    vinculos.append(r)
            nova_config = copy.deepcopy(self.config_data)
            cadastro = {c: campo.get() for c, campo in entradas.items()}
            cadastro["PRODUTOS"] = vinculos
            try:
                gestores.salvar_cadastro_gestor(nova_config, cadastro, indice)
                core.save_config(nova_config)
            except ValueError as exc:
                messagebox.showwarning("Gestores", str(exc), parent=popup)
                return
            except Exception as exc:
                messagebox.showerror("Gestores", f"Não foi possível salvar: {exc}", parent=popup)
                return
            self.config_data = nova_config
            self._recarregar_tabela_gestores()
            popup.destroy()
        botoes = ctk.CTkFrame(popup, fg_color="transparent")
        botoes.pack(fill="x", padx=15, pady=10)
        ctk.CTkButton(botoes, text="SALVAR", command=salvar).pack(side="left", padx=5)
        ctk.CTkButton(botoes, text="CANCELAR", command=popup.destroy).pack(side="right", padx=5)
        recarregar()

    # --- Vendedores ---
    def _montar_subtab_vendedores(self, frame):
        ctk.CTkLabel(
            frame,
            text="Vendedores cadastrados",
            font=ctk.CTkFont(size=16, weight="bold"),
        ).pack(pady=(15, 5), anchor="w", padx=10)

        vendedores_frame = ctk.CTkFrame(frame)
        vendedores_frame.pack(padx=10, pady=5, fill="both", expand=True)

        cols = ("nome", "email")
        self.tree_vendedores = ttk.Treeview(
            vendedores_frame, columns=cols, show="headings", height=10
        )
        for c, w in zip(cols, (300, 400)):
            self.tree_vendedores.heading(c, text=c.upper())
            self.tree_vendedores.column(c, width=w, anchor="w")
        self.tree_vendedores.pack(fill="both", expand=True, padx=5, pady=5)
        self._recarregar_tabela_vendedores()

        botoes = ctk.CTkFrame(frame, fg_color="transparent")
        botoes.pack(padx=10, pady=(5, 15), fill="x")
        ctk.CTkButton(
            botoes,
            text="Adicionar Vendedor",
            fg_color=AZUL_ESCURO,
            command=lambda: self._abrir_dialog_vendedor(),
        ).pack(side="left", padx=5)
        ctk.CTkButton(
            botoes,
            text="Editar Selecionado",
            fg_color=AZUL,
            command=self._editar_vendedor_selecionado,
        ).pack(side="left", padx=5)
        ctk.CTkButton(
            botoes,
            text="Remover Selecionado",
            fg_color=VERMELHO,
            command=self._remover_vendedor_selecionado,
        ).pack(side="left", padx=5)

    def _recarregar_tabela_vendedores(self):
        for item in self.tree_vendedores.get_children():
            self.tree_vendedores.delete(item)
        for v in self.config_data.get("vendedores", []):
            self.tree_vendedores.insert(
                "", "end", values=(v.get("nome", ""), v.get("email", ""))
            )

    def _vendedor_selecionado(self):
        sel = self.tree_vendedores.selection()
        if not sel:
            messagebox.showinfo("Aviso", "Selecione um vendedor na lista.")
            return None
        idx = self.tree_vendedores.index(sel[0])
        return self.config_data["vendedores"][idx]

    def _editar_vendedor_selecionado(self):
        vendedor = self._vendedor_selecionado()
        if vendedor:
            self._abrir_dialog_vendedor(vendedor)

    def _remover_vendedor_selecionado(self):
        vendedor = self._vendedor_selecionado()
        if not vendedor:
            return
        if messagebox.askyesno(
            "Confirmar", f"Remover o vendedor '{vendedor.get('nome')}'?"
        ):
            core.remover_vendedor(self.config_data, vendedor.get("nome"))
            core.save_config(self.config_data)
            self._recarregar_tabela_vendedores()
            self.cbo_vendedor.configure(values=self._nomes_vendedores())

    def _abrir_dialog_vendedor(self, vendedor: dict | None = None):
        popup = ctk.CTkToplevel(self)
        popup.title("Editar Vendedor" if vendedor else "Novo Vendedor")
        popup.geometry("480x240")
        popup.transient(self)
        popup.grab_set()

        ctk.CTkLabel(popup, text="Nome do vendedor").pack(
            anchor="w", padx=15, pady=(15, 0)
        )
        entry_nome = ctk.CTkEntry(popup)
        entry_nome.insert(0, vendedor.get("nome", "") if vendedor else "")
        entry_nome.pack(fill="x", padx=15)

        ctk.CTkLabel(popup, text="E-mail").pack(anchor="w", padx=15, pady=(10, 0))
        entry_email = ctk.CTkEntry(popup)
        entry_email.insert(0, vendedor.get("email", "") if vendedor else "")
        entry_email.pack(fill="x", padx=15)

        def salvar():
            novo_nome = entry_nome.get().strip()
            novo_email = entry_email.get().strip()
            if not novo_nome or not novo_email:
                messagebox.showwarning("Aviso", "Preencha nome e e-mail.")
                return
            if vendedor is None:
                if any(
                    (v.get("nome") or "").strip().upper() == novo_nome.upper()
                    for v in self.config_data.get("vendedores", [])
                ):
                    messagebox.showwarning(
                        "Aviso", f"Já existe um vendedor chamado '{novo_nome}'."
                    )
                    return
                core.adicionar_vendedor(self.config_data, novo_nome, novo_email)
            else:
                vendedor["nome"] = novo_nome
                vendedor["email"] = novo_email

            core.save_config(self.config_data)
            self._recarregar_tabela_vendedores()
            self.cbo_vendedor.configure(values=self._nomes_vendedores())
            popup.destroy()

        ctk.CTkButton(
            popup, text="SALVAR", fg_color=AZUL, height=38, command=salvar
        ).pack(pady=15, padx=15, fill="x")


if __name__ == "__main__":
    app = AppImportador()
    app.mainloop()
