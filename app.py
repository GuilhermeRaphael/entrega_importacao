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

        self.config_data = core.load_config()
        self.fila: list[core.ItemFila] = []
        self.ultimos_registros: list[dict] = []
        self.tem_auto_bonif = False

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
        self.cbo_produto = ctk.CTkComboBox(
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
        self.cbo_modelo = ctk.CTkComboBox(col2, values=[], width=180)
        self.cbo_modelo.set(PH_MODELO)
        self.cbo_modelo.pack(fill="x")

        col3 = ctk.CTkFrame(linha, fg_color="transparent")
        col3.pack(side="left", padx=5, fill="x", expand=True)
        ctk.CTkLabel(col3, text="Vendedor").pack(anchor="w")
        self.cbo_vendedor = ctk.CTkComboBox(
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
        ctk.CTkButton(
            frame,
            text="GERAR IMPORTAÇÃO",
            fg_color=AZUL,
            height=45,
            font=ctk.CTkFont(size=15, weight="bold"),
            command=self._gerar_importacao_thread,
        ).pack(padx=20, pady=(5, 5), fill="x")

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
            saldo = core.obter_modelos_disponiveis(caminho, self.fila, produto)
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao ler a matriz:\n{e}")
            return

        if not saldo:
            self.cbo_modelo.configure(values=["(NENHUM MODELO COM ESTOQUE)"])
            self.cbo_modelo.set("(NENHUM MODELO COM ESTOQUE)")
        else:
            valores = [f"{modelo} - {qtd}" for modelo, qtd in saldo.items()]
            self.cbo_modelo.configure(values=valores)
            self.cbo_modelo.set(PH_MODELO)

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

        modelo = (
            modelo_raw.split(" - ")[0].strip() if " - " in modelo_raw else modelo_raw
        )

        item = core.ItemFila(
            produto=produto, modelo=modelo, vendedor=vendedor, quantidade=int(qtd_raw)
        )
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
    def _gerar_importacao_thread(self):
        if not self.fila:
            messagebox.showwarning(
                "Aviso", "A fila de entregas está vazia! Adicione pelo menos um item."
            )
            return
        threading.Thread(target=self._gerar_importacao, daemon=True).start()

    def _gerar_importacao(self):
        self.lbl_status_geracao.configure(
            text="Processando matrizes, aguarde...", text_color="gray"
        )
        self._set_botoes_pos_geracao(habilitado=False)
        try:
            registros, tem_bonif, avisos = core.consolidar_remessa(
                self.config_data, self.fila
            )

            if not registros:
                self.lbl_status_geracao.configure(
                    text="Nenhum lead elegível localizado nas matrizes selecionadas.",
                    text_color="orange",
                )
                messagebox.showwarning(
                    "Aviso",
                    "Nenhum lead elegível localizado para os critérios informados!",
                )
                return

            caminho_saida = core.salvar_planilha_importacao(
                self.config_data, registros, tem_bonif
            )

            self.ultimos_registros = registros
            self.tem_auto_bonif = tem_bonif

            texto = f"✅ {len(registros)} leads importados. Backup salvo em:\n{caminho_saida}"
            if avisos:
                texto += "\n\nAvisos:\n" + "\n".join(avisos)
            self.lbl_status_geracao.configure(
                text=f"✅ {len(registros)} leads importados.", text_color=VERDE
            )
            self.lbl_resumo_envio.configure(
                text=f"{len(registros)} leads prontos. Escolha uma ação abaixo."
            )
            self._set_botoes_pos_geracao(habilitado=True)
            self.progress_bar.set(0)

            messagebox.showinfo("Importação Concluída", texto)
        except Exception:
            erro = traceback.format_exc()
            self.lbl_status_geracao.configure(
                text="Erro ao gerar importação.", text_color="red"
            )
            messagebox.showerror(
                "Erro", f"Ocorreu um erro ao gerar a importação:\n\n{erro}"
            )

    def _set_botoes_pos_geracao(self, habilitado: bool):
        estado = "normal" if habilitado else "disabled"
        self.btn_enviar_goalfy.configure(state=estado)
        self.btn_enviar_email.configure(state=estado)
        self.btn_imprimir.configure(state=estado)

    # ------------------------------------------------------------------
    # Envio para o Goalfy
    # ------------------------------------------------------------------
    def _enviar_goalfy_thread(self):
        if not self.ultimos_registros:
            return
        threading.Thread(target=self._enviar_goalfy, daemon=True).start()

    def _enviar_goalfy(self):
        self._set_botoes_pos_geracao(habilitado=False)

        def progresso(i, total, ok):
            self.progress_bar.set(i / total)
            self.lbl_resumo_envio.configure(
                text=f"Enviando {i} de {total} leads ao Goalfy..."
            )

        try:
            sucessos, falhas, erros = core.enviar_para_goalfy(
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
            self._set_botoes_pos_geracao(habilitado=True)

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

        self._montar_subtab_geral(tab_geral)
        self._montar_subtab_matrizes(tab_matrizes)
        self._montar_subtab_vendedores(tab_vendedores)

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
        self.config_data["pasta_destino"] = self.entry_pasta_destino.get().strip()
        self.config_data["webhook_url"] = self.entry_webhook.get().strip()
        core.save_config(self.config_data)
        messagebox.showinfo("Configurações", "Configurações gerais salvas com sucesso!")

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
