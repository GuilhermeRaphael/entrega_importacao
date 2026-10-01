"""Campo editável com sugestões durante a digitação, integrado ao tema do app."""
import tkinter as tk
import unicodedata
import customtkinter as ctk


def chave(texto):
    return "".join(c for c in unicodedata.normalize("NFD", str(texto).upper()) if unicodedata.category(c) != "Mn")

class CampoAutocomplete(ctk.CTkComboBox):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._opcoes = list(kwargs.get("values", []))
        self._popup_sugestoes = None
        self._sugestoes = []
        self._entry.bind("<KeyRelease>", self._digitar, add="+")
        self._entry.bind("<Down>", self._baixo, add="+")
        self._entry.bind("<Up>", self._cima, add="+")
        self._entry.bind("<Return>", self._aceitar, add="+")
        self._entry.bind("<Escape>", lambda e: self._fechar(), add="+")
        self._entry.bind("<FocusOut>", lambda e: self.after(120, self._checar_foco), add="+")
        self.bind("<Destroy>", lambda e: self._fechar(), add="+")

    def configure(self, require_redraw=False, **kwargs):
        if "values" in kwargs:
            self._opcoes = list(kwargs["values"])
        return super().configure(require_redraw=require_redraw, **kwargs)

    def _fechar(self):
        if self._popup_sugestoes is not None:
            try:
                self._popup_sugestoes.destroy()
            except tk.TclError:
                pass
            self._popup_sugestoes = None

    def _checar_foco(self):
        if self.focus_get() is not self._entry:
            self._fechar()

    def _digitar(self, event=None):
        if event is not None and event.keysym in ("Up", "Down", "Return", "Escape", "Tab", "Shift_L", "Shift_R"):
            return
        if self.cget("state") != "normal":
            return
        texto = chave(self.get().strip())
        self._sugestoes = [v for v in self._opcoes if texto and texto in chave(v)][:8]
        self._fechar()
        if not self._sugestoes:
            return
        popup = tk.Toplevel(self.winfo_toplevel())
        popup.overrideredirect(True)
        popup.geometry(f"{max(self.winfo_width(),240)}x{len(self._sugestoes)*24+4}+{self.winfo_rootx()}+{self.winfo_rooty()+self.winfo_height()}")
        escuro = ctk.get_appearance_mode() == "Dark"
        lista = tk.Listbox(popup, font=("Segoe UI", 11), activestyle="none", relief="solid", borderwidth=1,
                           bg="#2b2b2b" if escuro else "#ffffff", fg="#ffffff" if escuro else "#202020",
                           selectbackground="#0066CC", selectforeground="white", exportselection=False)
        lista.pack(fill="both", expand=True)
        for v in self._sugestoes:
            lista.insert("end", v)
        lista.selection_set(0)
        lista.bind("<ButtonRelease-1>", self._aceitar)
        self._lista_sugestoes = lista
        self._popup_sugestoes = popup
        popup.lift()
        self._entry.focus_set()

    def _baixo(self, event):
        if self._popup_sugestoes is None:
            self._digitar()
            return "break"
        indice = min(int(self._lista_sugestoes.curselection()[0])+1, len(self._sugestoes)-1)
        self._lista_sugestoes.selection_clear(0,"end")
        self._lista_sugestoes.selection_set(indice)
        return "break"

    def _cima(self, event):
        if self._popup_sugestoes is not None:
            indice = max(int(self._lista_sugestoes.curselection()[0])-1,0)
            self._lista_sugestoes.selection_clear(0,"end")
            self._lista_sugestoes.selection_set(indice)
        return "break"

    def _aceitar(self, event=None):
        if self._popup_sugestoes is not None:
            selecao = self._lista_sugestoes.curselection()
            valor = self._sugestoes[int(selecao[0])] if selecao else None
        else:
            valor = next((v for v in self._opcoes if chave(v.strip()) == chave(self.get().strip())), None)
        if valor is not None:
            self.set(valor)
            self._fechar()
            self._entry.focus_set()
            if self._command:
                self._command(valor)
        return "break"
