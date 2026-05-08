#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Dénombrement UFC/g  v3.0  —  NF ISO 7218
"""

import tkinter as tk
import tkinter.font as tkfont
from tkinter import ttk, messagebox, filedialog
import json, os, sys, csv
from datetime import datetime

# ── Palette raffinée ─────────────────────────────────────────────────────────
C = {
    # Fonds
    "app_bg":       "#F4F6F8",
    "panel":        "#FFFFFF",
    "panel_alt":    "#F8FAFC",

    # Header
    "header":       "#1E293B",
    "header_sub":   "#334155",
    "header_text":  "#F8FAFC",
    "header_muted": "#94A3B8",

    # Texte
    "text":         "#0F172A",
    "text_sec":     "#475569",
    "text_muted":   "#94A3B8",

    # Bordures
    "border":       "#E2E8F0",
    "border_dark":  "#CBD5E1",

    # Accent principal
    "accent":       "#3B5BDB",
    "accent_hover": "#2F4AC0",

    # Boutons
    "btn_primary":  "#3B5BDB",
    "btn_neutral":  "#64748B",
    "btn_danger":   "#DC2626",
    "btn_text":     "#FFFFFF",

    # Statut champs (subtil)
    "ok_bg":        "#F0FDF4",
    "ok_border":    "#86EFAC",
    "ok_text":      "#166534",
    "warn_bg":      "#FFFBEB",
    "warn_border":  "#FCD34D",
    "warn_text":    "#92400E",
    "err_bg":       "#FFF1F2",
    "err_border":   "#FECDD3",
    "err_text":     "#9F1239",
    "nc_bg":        "#EFF6FF",
    "nc_border":    "#BFDBFE",
    "nc_text":      "#1E40AF",
    "empty":        "#FFFFFF",

    # Zone résultat
    "res_bg":       "#1E293B",
    "res_num":      "#F8FAFC",
    "res_sat":      "#4ADE80",
    "res_lim":      "#FCD34D",
    "res_insat":    "#F87171",
    "res_detail":   "#94A3B8",

    # Interprétation badge
    "sat_bg":       "#DCFCE7",
    "sat_fg":       "#166534",
    "lim_bg":       "#FEF9C3",
    "lim_fg":       "#713F12",
    "ins_bg":       "#FFE4E6",
    "ins_fg":       "#9F1239",
}

FF   = "Segoe UI"
MONO = "Consolas"

NB_DIL = 6

# ── Seuils par défaut (germe → matrice → {m, M}) ─────────────────────────────
DEFAULT_SEUILS = {
    "Flore aérobie mésophile": {
        "Viandes hachées crues":    {"m": 5e5,  "M": 5e6},
        "Viandes cuites":           {"m": 1e4,  "M": 1e5},
        "Produits laitiers":        {"m": 3e4,  "M": 1e5},
        "Fromages affinés":         {"m": 1e5,  "M": 1e6},
        "Plats cuisinés":           {"m": 1e5,  "M": 5e5},
        "Fruits et légumes frais":  {"m": 1e6,  "M": 1e7},
        "Poissons frais":           {"m": 1e6,  "M": 1e7},
    },
    "Entérobactéries": {
        "Viandes hachées crues":    {"m": 5e2,  "M": 5e3},
        "Viandes cuites":           {"m": 10,   "M": 1e2},
        "Produits laitiers":        {"m": 10,   "M": 1e2},
        "Fromages affinés":         {"m": 1e2,  "M": 1e3},
        "Plats cuisinés":           {"m": 10,   "M": 1e2},
        "Poissons frais":           {"m": 1e2,  "M": 1e3},
    },
    "Staphylocoques coag. +": {
        "Viandes hachées crues":    {"m": 5e2,  "M": 5e3},
        "Viandes cuites":           {"m": 1e2,  "M": 1e3},
        "Fromages affinés":         {"m": 1e2,  "M": 1e3},
        "Plats cuisinés":           {"m": 1e2,  "M": 1e3},
        "Crèmes et desserts":       {"m": 10,   "M": 1e2},
    },
    "Levures / Moisissures": {
        "Produits laitiers":        {"m": 1e2,  "M": 1e3},
        "Fromages affinés":         {"m": 5e2,  "M": 5e3},
        "Fruits et légumes frais":  {"m": 1e3,  "M": 1e4},
        "Produits de boulangerie":  {"m": 1e2,  "M": 1e3},
    },
    "Coliformes totaux": {
        "Eau potable":              {"m": 0,    "M": 0},
        "Produits laitiers":        {"m": 10,   "M": 1e2},
        "Glaces et sorbets":        {"m": 10,   "M": 1e2},
    },
    "Coliformes thermo.": {
        "Eau potable":              {"m": 0,    "M": 0},
        "Viandes hachées crues":    {"m": 5e2,  "M": 5e3},
        "Fromages affinés":         {"m": 1e2,  "M": 1e3},
    },
    "Listeria spp.": {
        "Viandes cuites":           {"m": 0,    "M": 1e2},
        "Produits laitiers":        {"m": 0,    "M": 1e2},
        "Fromages affinés":         {"m": 0,    "M": 1e2},
        "Plats cuisinés":           {"m": 0,    "M": 1e2},
        "Poissons fumés":           {"m": 0,    "M": 1e2},
    },
    "Salmonella": {
        "Viandes hachées crues":    {"m": 0,    "M": 0},
        "Viandes cuites":           {"m": 0,    "M": 0},
        "Produits laitiers":        {"m": 0,    "M": 0},
        "Fromages affinés":         {"m": 0,    "M": 0},
        "Plats cuisinés":           {"m": 0,    "M": 0},
        "Fruits et légumes frais":  {"m": 0,    "M": 0},
    },
}


# ── Configuration ─────────────────────────────────────────────────────────────

class Config:
    _DEFAULTS = {
        "instance":      "False",
        "ensemencement": "profondeur",
        "operateur":     "",
        "seuils":        DEFAULT_SEUILS,
    }

    def __init__(self):
        self.path = os.path.join(
            os.environ.get("USERPROFILE", os.path.expanduser("~")),
            "AppData", "Local", "denombrement", "conf.json",
        )
        try:
            with open(self.path) as f:
                saved = json.load(f)
                self.data = {**self._DEFAULTS, **saved}
                if "seuils" not in saved:
                    self.data["seuils"] = DEFAULT_SEUILS
        except (FileNotFoundError, json.JSONDecodeError):
            os.makedirs(os.path.dirname(self.path), exist_ok=True)
            self.data = {k: v for k, v in self._DEFAULTS.items()}
            self._write()

    def _write(self):
        with open(self.path, "w") as f:
            json.dump(self.data, f, indent=2, ensure_ascii=False)

    def set(self, key, value):
        self.data[key] = value
        self._write()

    def get(self, key, default=None):
        return self.data.get(key, default)


# ── Boîte de dialogue : gestion des seuils ───────────────────────────────────

class SeuilsDialog(tk.Toplevel):

    def __init__(self, parent, conf, on_save):
        super().__init__(parent)
        self.conf    = conf
        self.on_save = on_save
        self.seuils  = {g: {m: dict(v) for m, v in ms.items()}
                        for g, ms in conf.get("seuils", DEFAULT_SEUILS).items()}

        self.title("Paramètres — Seuils de conformité")
        self.resizable(True, True)
        self.geometry("820x560")
        self.configure(bg=C["panel"])
        self.grab_set()

        self._build_fonts()
        self._build_ui()
        self._refresh_germe_list()

    def _build_fonts(self):
        self.f_title  = tkfont.Font(family=FF, size=13, weight="bold")
        self.f_label  = tkfont.Font(family=FF, size=9)
        self.f_btn    = tkfont.Font(family=FF, size=9, weight="bold")
        self.f_small  = tkfont.Font(family=FF, size=8)
        self.f_mono   = tkfont.Font(family=MONO, size=9)

    def _build_ui(self):
        # ── En-tête ──────────────────────────────────────────────────────
        hdr = tk.Frame(self, bg=C["header"], pady=12)
        hdr.pack(fill="x")
        tk.Label(hdr, text="Seuils de conformité",
                 font=self.f_title, bg=C["header"], fg=C["header_text"]).pack(side="left", padx=16)
        tk.Label(hdr, text="Paramétrez m et M par germe et matrice produit",
                 font=self.f_small, bg=C["header"], fg=C["header_muted"]).pack(side="left", padx=4)

        body = tk.Frame(self, bg=C["panel"])
        body.pack(fill="both", expand=True, padx=16, pady=12)

        # ── Colonne gauche : liste des germes ────────────────────────────
        left = tk.Frame(body, bg=C["panel"], width=200)
        left.pack(side="left", fill="y", padx=(0, 12))
        left.pack_propagate(False)

        tk.Label(left, text="GERMES", font=self.f_small,
                 bg=C["panel"], fg=C["text_muted"]).pack(anchor="w", pady=(0, 4))

        lbf = tk.Frame(left, bg=C["border"], bd=1, relief="flat")
        lbf.pack(fill="both", expand=True)
        self.lb_germes = tk.Listbox(lbf, font=self.f_label, bd=0, relief="flat",
                                     selectbackground=C["accent"], selectforeground="white",
                                     activestyle="none", bg=C["panel"])
        sb_g = ttk.Scrollbar(lbf, orient="vertical", command=self.lb_germes.yview)
        self.lb_germes.configure(yscrollcommand=sb_g.set)
        self.lb_germes.pack(side="left", fill="both", expand=True)
        sb_g.pack(side="right", fill="y")
        self.lb_germes.bind("<<ListboxSelect>>", self._on_germe_select)

        btn_g = tk.Frame(left, bg=C["panel"], pady=4)
        btn_g.pack(fill="x")
        self._btn(btn_g, "+ Germe",   self._add_germe,    C["btn_primary"]).pack(side="left", padx=(0,4))
        self._btn(btn_g, "- Germe",   self._del_germe,    C["btn_danger"]).pack(side="left")

        # ── Colonne droite : table des matrices ──────────────────────────
        right = tk.Frame(body, bg=C["panel"])
        right.pack(side="left", fill="both", expand=True)

        self.lbl_germe_title = tk.Label(right, text="— sélectionnez un germe —",
                                         font=self.f_label, bg=C["panel"], fg=C["text_muted"])
        self.lbl_germe_title.pack(anchor="w", pady=(0, 4))

        # Table
        tf = tk.Frame(right, bg=C["panel"])
        tf.pack(fill="both", expand=True)

        cols = ("matrice", "m", "M")
        self.tree = ttk.Treeview(tf, columns=cols, show="headings", height=12)
        headers = {"matrice": ("Matrice produit", 280), "m": ("m  (ufc/g)", 130), "M": ("M  (ufc/g)", 130)}
        for col, (hdg, w) in headers.items():
            self.tree.heading(col, text=hdg)
            self.tree.column(col, width=w, anchor="center" if col != "matrice" else "w")

        style = ttk.Style()
        style.configure("Seuils.Treeview",         font=(FF, 9), rowheight=24)
        style.configure("Seuils.Treeview.Heading", font=(FF, 9, "bold"))
        self.tree.configure(style="Seuils.Treeview")

        sb_t = ttk.Scrollbar(tf, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb_t.set)

        self.tree.pack(side="left", fill="both", expand=True)
        sb_t.pack(side="right", fill="y")
        self.tree.bind("<Double-1>", self._edit_row)

        btn_t = tk.Frame(right, bg=C["panel"], pady=6)
        btn_t.pack(fill="x")
        self._btn(btn_t, "+ Matrice",  self._add_matrice,  C["btn_primary"]).pack(side="left", padx=(0,4))
        self._btn(btn_t, "Modifier",   self._edit_row,     C["btn_neutral"]).pack(side="left", padx=(0,4))
        self._btn(btn_t, "Supprimer",  self._del_matrice,  C["btn_danger"]).pack(side="left")

        # ── Pied de page ────────────────────────────────────────────────
        foot = tk.Frame(self, bg=C["border"], height=1)
        foot.pack(fill="x")
        btns = tk.Frame(self, bg=C["panel"], pady=10)
        btns.pack(fill="x", padx=16)
        self._btn(btns, "Annuler",     self.destroy,       C["btn_neutral"]).pack(side="right", padx=(6,0))
        self._btn(btns, "Enregistrer", self._save,         C["btn_primary"]).pack(side="right")
        self._btn(btns, "Restaurer défauts", self._restore_defaults, "#64748B").pack(side="left")

    def _btn(self, parent, text, cmd, bg):
        return tk.Button(parent, text=text, command=cmd,
                         font=self.f_btn, bg=bg, fg="white",
                         relief="flat", cursor="hand2", padx=10, pady=4,
                         activebackground=bg, activeforeground="white")

    # ── Actions germes ───────────────────────────────────────────────────────

    def _refresh_germe_list(self):
        self.lb_germes.delete(0, "end")
        for g in sorted(self.seuils):
            self.lb_germes.insert("end", f"  {g}")

    def _on_germe_select(self, _=None):
        sel = self.lb_germes.curselection()
        if not sel:
            return
        germe = self.lb_germes.get(sel[0]).strip()
        self.current_germe = germe
        self.lbl_germe_title.config(text=germe, fg=C["text"])
        self._refresh_matrix_tree()

    def _add_germe(self):
        d = _InputDialog(self, "Nouveau germe", "Nom du germe :")
        name = d.result
        if name and name.strip() and name.strip() not in self.seuils:
            self.seuils[name.strip()] = {}
            self._refresh_germe_list()

    def _del_germe(self):
        sel = self.lb_germes.curselection()
        if not sel:
            return
        germe = self.lb_germes.get(sel[0]).strip()
        if messagebox.askyesno("Supprimer", f"Supprimer le germe « {germe} » et toutes ses matrices ?", parent=self):
            del self.seuils[germe]
            self._refresh_germe_list()
            self.tree.delete(*self.tree.get_children())
            self.lbl_germe_title.config(text="— sélectionnez un germe —", fg=C["text_muted"])

    # ── Actions matrices ──────────────────────────────────────────────────────

    def _refresh_matrix_tree(self):
        self.tree.delete(*self.tree.get_children())
        germe = getattr(self, "current_germe", None)
        if not germe or germe not in self.seuils:
            return
        for mat, vals in sorted(self.seuils[germe].items()):
            m_str = _fmt_seuil(vals["m"])
            M_str = _fmt_seuil(vals["M"])
            self.tree.insert("", "end", iid=mat, values=(mat, m_str, M_str))

    def _add_matrice(self):
        germe = getattr(self, "current_germe", None)
        if not germe:
            messagebox.showinfo("Info", "Sélectionnez d'abord un germe.", parent=self)
            return
        d = _SeuilEditDialog(self, "Nouvelle matrice", "", None, None)
        if d.result:
            mat, m, M = d.result
            self.seuils[germe][mat] = {"m": m, "M": M}
            self._refresh_matrix_tree()

    def _edit_row(self, _=None):
        sel = self.tree.selection()
        if not sel:
            return
        germe = getattr(self, "current_germe", None)
        mat   = sel[0]
        vals  = self.seuils[germe][mat]
        d = _SeuilEditDialog(self, "Modifier", mat, vals["m"], vals["M"])
        if d.result:
            new_mat, m, M = d.result
            del self.seuils[germe][mat]
            self.seuils[germe][new_mat] = {"m": m, "M": M}
            self._refresh_matrix_tree()

    def _del_matrice(self):
        sel = self.tree.selection()
        if not sel:
            return
        germe = getattr(self, "current_germe", None)
        mat   = sel[0]
        if messagebox.askyesno("Supprimer", f"Supprimer la matrice « {mat} » ?", parent=self):
            del self.seuils[germe][mat]
            self._refresh_matrix_tree()

    def _save(self):
        self.conf.set("seuils", self.seuils)
        self.on_save()
        self.destroy()

    def _restore_defaults(self):
        if messagebox.askyesno("Restaurer", "Restaurer tous les seuils par défaut ?", parent=self):
            self.seuils = {g: {m: dict(v) for m, v in ms.items()}
                           for g, ms in DEFAULT_SEUILS.items()}
            self._refresh_germe_list()
            self.tree.delete(*self.tree.get_children())


# ── Dialogues utilitaires ─────────────────────────────────────────────────────

class _InputDialog(tk.Toplevel):
    def __init__(self, parent, title, label):
        super().__init__(parent)
        self.result = None
        self.title(title)
        self.resizable(False, False)
        self.grab_set()
        self.configure(bg=C["panel"])
        f = tkfont.Font(family=FF, size=9)
        tk.Label(self, text=label, font=f, bg=C["panel"], fg=C["text"]).pack(padx=20, pady=(16,4))
        self.entry = tk.Entry(self, font=f, width=32, bd=1, relief="solid")
        self.entry.pack(padx=20, pady=(0,10))
        self.entry.focus_set()
        btns = tk.Frame(self, bg=C["panel"])
        btns.pack(pady=(0,12))
        tk.Button(btns, text="OK", font=f, bg=C["btn_primary"], fg="white",
                  relief="flat", padx=14, command=self._ok).pack(side="left", padx=4)
        tk.Button(btns, text="Annuler", font=f, bg=C["btn_neutral"], fg="white",
                  relief="flat", padx=10, command=self.destroy).pack(side="left")
        self.entry.bind("<Return>", lambda e: self._ok())
        self.wait_window()

    def _ok(self):
        self.result = self.entry.get().strip()
        self.destroy()


class _SeuilEditDialog(tk.Toplevel):
    def __init__(self, parent, title, matrice, m, M):
        super().__init__(parent)
        self.result = None
        self.title(title)
        self.resizable(False, False)
        self.grab_set()
        self.configure(bg=C["panel"])
        f  = tkfont.Font(family=FF, size=9)
        fb = tkfont.Font(family=FF, size=9, weight="bold")

        def row(lbl, default, r):
            tk.Label(self, text=lbl, font=f, bg=C["panel"],
                     fg=C["text_sec"]).grid(row=r, column=0, sticky="e", padx=(16,6), pady=5)
            e = tk.Entry(self, font=tkfont.Font(family=MONO, size=9),
                         width=18, bd=1, relief="solid")
            e.insert(0, _fmt_seuil(default) if default is not None else "0")
            e.grid(row=r, column=1, padx=(0,16), pady=5)
            return e

        self.e_mat = row("Matrice produit :", None, 0)
        self.e_mat.delete(0, "end")
        self.e_mat.insert(0, matrice)

        self.e_m  = row("m  (seuil de qualité, ufc/g) :", m, 1)
        self.e_M  = row("M  (limite d'action, ufc/g) :", M, 2)

        tk.Label(self, text="Exemples : 100, 1e3, 50000", font=tkfont.Font(family=FF, size=8),
                 bg=C["panel"], fg=C["text_muted"]).grid(row=3, columnspan=2, pady=(0,6))

        btns = tk.Frame(self, bg=C["panel"])
        btns.grid(row=4, columnspan=2, pady=(0,12))
        tk.Button(btns, text="Valider", font=fb, bg=C["btn_primary"], fg="white",
                  relief="flat", padx=14, command=self._ok).pack(side="left", padx=4)
        tk.Button(btns, text="Annuler", font=f, bg=C["btn_neutral"], fg="white",
                  relief="flat", padx=10, command=self.destroy).pack(side="left")
        self.e_mat.focus_set()
        self.wait_window()

    def _ok(self):
        try:
            mat = self.e_mat.get().strip()
            m   = float(self.e_m.get().strip())
            M   = float(self.e_M.get().strip())
            if not mat:
                raise ValueError
            self.result = (mat, m, M)
            self.destroy()
        except ValueError:
            messagebox.showerror("Erreur", "Valeurs invalides.", parent=self)


def _fmt_seuil(v):
    if v is None:
        return "—"
    if v == 0:
        return "0  (absent)"
    if v >= 1e6:
        return f"{v:.0e}"
    return f"{int(v):,}".replace(",", " ")


# ── Application principale ────────────────────────────────────────────────────

class DenombrementApp(tk.Frame):

    def __init__(self, root, conf):
        super().__init__(root, bg=C["app_bg"])
        self.root    = root
        self.conf    = conf
        self.history = []

        self._build_fonts()
        self._build_vars()
        self._build_ui()
        self._apply_prefs()

    # ── Polices ──────────────────────────────────────────────────────────────

    def _build_fonts(self):
        self.f_header = tkfont.Font(family=FF, size=14, weight="bold")
        self.f_sub    = tkfont.Font(family=FF, size=8)
        self.f_label  = tkfont.Font(family=FF, size=9)
        self.f_label_b= tkfont.Font(family=FF, size=9, weight="bold")
        self.f_entry  = tkfont.Font(family=MONO, size=14)
        self.f_btn    = tkfont.Font(family=FF, size=10, weight="bold")
        self.f_result = tkfont.Font(family=FF, size=26, weight="bold")
        self.f_interp = tkfont.Font(family=FF, size=10, weight="bold")
        self.f_detail = tkfont.Font(family=FF, size=8)
        self.f_dil    = tkfont.Font(family=FF, size=7, slant="italic")

    # ── Variables ─────────────────────────────────────────────────────────────

    def _build_vars(self):
        self.varfields   = [[tk.StringVar() for _ in range(2)] for _ in range(NB_DIL)]
        self.volume      = tk.DoubleVar(value=1.0)
        self.echantillon = tk.StringVar()
        self.operateur   = tk.StringVar(value=self.conf.get("operateur", ""))
        self.germe       = tk.StringVar()
        self.matrice     = tk.StringVar()

        for i in range(NB_DIL):
            for j in range(2):
                self.varfields[i][j].trace_add("write",
                    lambda *a, i=i, j=j: self._on_field_change(i, j))

        self.operateur.trace_add("write",
            lambda *a: self.conf.set("operateur", self.operateur.get()))

    def _apply_prefs(self):
        if self.conf.get("ensemencement") == "surface":
            self.volume.set(0.1)

    # ── Construction UI ───────────────────────────────────────────────────────

    def _build_ui(self):
        self._build_header()
        self._build_sample_bar()
        self._build_dilution_grid()
        self._build_legend()
        self._build_buttons()
        self._build_result_panel()
        self._build_history_panel()

    # ── Header ────────────────────────────────────────────────────────────────

    def _build_header(self):
        hdr = tk.Frame(self, bg=C["header"])
        hdr.pack(fill="x")

        left = tk.Frame(hdr, bg=C["header"], padx=16, pady=12)
        left.pack(side="left")
        tk.Label(left, text="DÉNOMBREMENT  UFC/g",
                 font=self.f_header, bg=C["header"], fg=C["header_text"]).pack(anchor="w")
        tk.Label(left, text="Méthode NF ISO 7218  •  Version 3.0",
                 font=self.f_sub, bg=C["header"], fg=C["header_muted"]).pack(anchor="w")

        right = tk.Frame(hdr, bg=C["header"], padx=14)
        right.pack(side="right", fill="y")

        # Badge mode ensemencement — cliquable
        self.badge_mode = tk.Label(right, font=self.f_label_b,
                                    bg=C["header_sub"], fg=C["header_text"],
                                    padx=12, pady=5, cursor="hand2")
        self.badge_mode.pack(side="right", pady=12)
        self.badge_mode.bind("<Button-1>", self._toggle_mode)
        self.volume.trace_add("write", lambda *a: self._refresh_mode_badge())
        self._refresh_mode_badge()

        # Horloge
        self.lbl_clock = tk.Label(right, font=self.f_sub,
                                   bg=C["header"], fg=C["header_muted"])
        self.lbl_clock.pack(side="right", padx=16)
        self._tick_clock()

    # ── Barre d'identification ─────────────────────────────────────────────────

    def _build_sample_bar(self):
        bar = tk.Frame(self, bg=C["panel"],
                       highlightbackground=C["border"], highlightthickness=1)
        bar.pack(fill="x", padx=10, pady=(8, 0))

        inner = tk.Frame(bar, bg=C["panel"], pady=7, padx=12)
        inner.pack(fill="x")

        def field_group(parent, label, var, w, col):
            tk.Label(parent, text=label, font=self.f_sub,
                     bg=C["panel"], fg=C["text_muted"]).grid(row=0, column=col*2, sticky="w", padx=(0 if col==0 else 16, 0))
            e = tk.Entry(parent, textvariable=var, font=self.f_label,
                         width=w, bd=0, bg=C["panel_alt"],
                         highlightthickness=1, highlightbackground=C["border"],
                         highlightcolor=C["accent"], relief="flat")
            e.grid(row=1, column=col*2, padx=(0 if col==0 else 16, 0), sticky="w", ipady=3)
            return e

        field_group(inner, "ÉCHANTILLON", self.echantillon, 16, 0)
        field_group(inner, "OPÉRATEUR",   self.operateur,   14, 1)

        # Germe
        tk.Label(inner, text="GERME", font=self.f_sub,
                 bg=C["panel"], fg=C["text_muted"]).grid(row=0, column=4, sticky="w", padx=(16,0))
        seuils = self.conf.get("seuils", DEFAULT_SEUILS)
        self.cb_germe = ttk.Combobox(inner, textvariable=self.germe,
                                      values=["(Aucun seuil)"] + sorted(seuils.keys()),
                                      width=24, state="readonly", font=self.f_label)
        self.cb_germe.grid(row=1, column=4, padx=(16,0), sticky="w")
        self.germe.set("(Aucun seuil)")

        # Matrice
        tk.Label(inner, text="MATRICE PRODUIT", font=self.f_sub,
                 bg=C["panel"], fg=C["text_muted"]).grid(row=0, column=5, sticky="w", padx=(12,0))
        self.cb_matrice = ttk.Combobox(inner, textvariable=self.matrice,
                                        values=[], width=22,
                                        state="readonly", font=self.f_label)
        self.cb_matrice.grid(row=1, column=5, padx=(12,0), sticky="w")

        # Connecter le trace après création de cb_matrice
        self.germe.trace_add("write", lambda *a: self._on_germe_change())

        # Date
        self.lbl_date = tk.Label(inner, font=self.f_sub,
                                  bg=C["panel"], fg=C["text_muted"])
        self.lbl_date.grid(row=0, column=6, rowspan=2, sticky="e", padx=(12,0))
        inner.columnconfigure(6, weight=1)
        self._tick_date()

    # ── Grille de dilution ────────────────────────────────────────────────────

    def _build_dilution_grid(self):
        wrapper = tk.Frame(self, bg=C["app_bg"])
        wrapper.pack(fill="x", padx=10, pady=(10, 0))

        self.entry_widgets = []
        self.entry_frames  = []

        for i in range(NB_DIL):
            card = tk.Frame(wrapper, bg=C["panel"],
                            highlightbackground=C["border"], highlightthickness=1)
            card.grid(row=0, column=i, padx=4, sticky="nsew")
            wrapper.columnconfigure(i, weight=1)

            # Bande colorée en haut
            stripe = tk.Frame(card, bg=C["header_sub"], height=4)
            stripe.pack(fill="x")

            # Label dilution
            tk.Label(card, text=f"Dil. {i+1}",
                     font=self.f_label_b, bg=C["panel"], fg=C["text"]).pack(pady=(6,0))
            tk.Label(card, text=f"10⁻{i+1}",
                     font=self.f_dil, bg=C["panel"], fg=C["text_muted"]).pack()

            # Séparateur
            tk.Frame(card, bg=C["border"], height=1).pack(fill="x", padx=8, pady=4)

            row_w, row_f = [], []
            for j in range(2):
                inner = tk.Frame(card, bg=C["panel"])
                inner.pack(fill="x", padx=8, pady=(0,6))
                tk.Label(inner, text=f"B{j+1}", font=self.f_dil,
                         bg=C["panel"], fg=C["text_muted"]).pack(anchor="w")
                ef = tk.Frame(inner, bg=C["empty"],
                              highlightbackground=C["border_dark"],
                              highlightthickness=1)
                ef.pack(fill="x")
                e = tk.Entry(ef, textvariable=self.varfields[i][j],
                             font=self.f_entry, width=6, bd=0,
                             bg=C["empty"], justify="center",
                             insertbackground=C["text"])
                e.pack(pady=5, padx=4)
                row_w.append(e)
                row_f.append(ef)

            self.entry_widgets.append(row_w)
            self.entry_frames.append(row_f)

        self._setup_kb_nav()

    def _setup_kb_nav(self):
        flat = [self.entry_widgets[i][j] for i in range(NB_DIL) for j in range(2)]
        for k, e in enumerate(flat):
            e.bind("<Return>", lambda ev, n=flat[(k+1) % len(flat)]: n.focus_set())
        flat[-1].bind("<Return>", lambda ev: self.calculate())

    # ── Légende couleurs ──────────────────────────────────────────────────────

    def _build_legend(self):
        bar = tk.Frame(self, bg=C["app_bg"])
        bar.pack(fill="x", padx=14, pady=(4,0))

        items = [
            (C["ok_bg"],   C["ok_border"],   "30 – 300  (retenu)"),
            (C["warn_bg"], C["warn_border"],  "Hors plage"),
            (C["err_bg"],  C["err_border"],   "Valeur invalide"),
            (C["nc_bg"],   C["nc_border"],    "NC"),
        ]
        for bg, bd, txt in items:
            dot = tk.Frame(bar, bg=bg, width=12, height=12,
                           highlightbackground=bd, highlightthickness=1)
            dot.pack(side="left", padx=(0,3))
            dot.pack_propagate(False)
            tk.Label(bar, text=txt, font=self.f_dil,
                     bg=C["app_bg"], fg=C["text_sec"]).pack(side="left", padx=(0,14))

    # ── Boutons ───────────────────────────────────────────────────────────────

    def _build_buttons(self):
        frm = tk.Frame(self, bg=C["app_bg"])
        frm.pack(pady=10)

        def btn(text, cmd, bg, w=16):
            return tk.Button(frm, text=text, command=cmd,
                             font=self.f_btn, bg=bg, fg=C["btn_text"],
                             relief="flat", cursor="hand2", width=w,
                             activebackground=bg, activeforeground="white",
                             pady=7)

        btn("Réinitialiser",    self.reset,     C["btn_neutral"]).pack(side="left", padx=6)
        btn("Calculer  [F5]",   self.calculate, C["btn_primary"]).pack(side="left", padx=6)

        self.root.bind("<F5>",     lambda e: self.calculate())
        self.root.bind("<Escape>", lambda e: self.reset())

    # ── Panneau résultat ──────────────────────────────────────────────────────

    def _build_result_panel(self):
        outer = tk.Frame(self, bg=C["app_bg"])
        outer.pack(fill="x", padx=10, pady=(0,8))

        res = tk.Frame(outer, bg=C["res_bg"])
        res.pack(fill="x")

        # Nombre
        tk.Label(res, text="RÉSULTAT", font=self.f_sub,
                 bg=C["res_bg"], fg=C["header_muted"]).pack(pady=(14,0))
        self.lbl_result = tk.Label(res, text="— — —",
                                    font=self.f_result,
                                    bg=C["res_bg"], fg=C["res_num"])
        self.lbl_result.pack()

        # Badge interprétation
        self.lbl_interp = tk.Label(res, text="",
                                    font=self.f_interp,
                                    bg=C["res_bg"], fg=C["res_num"],
                                    padx=14, pady=3)
        self.lbl_interp.pack(pady=(4,0))

        # Détail
        self.lbl_detail = tk.Label(res, text="",
                                    font=self.f_detail,
                                    bg=C["res_bg"], fg=C["res_detail"],
                                    wraplength=920)
        self.lbl_detail.pack(pady=(4,14))

    # ── Historique ────────────────────────────────────────────────────────────

    def _build_history_panel(self):
        panel = tk.Frame(self, bg=C["panel"],
                         highlightbackground=C["border"], highlightthickness=1)
        panel.pack(fill="both", expand=True, padx=10, pady=(0,10))

        top = tk.Frame(panel, bg=C["panel"])
        top.pack(fill="x", padx=10, pady=(6,2))
        tk.Label(top, text="Historique de session",
                 font=self.f_label_b, bg=C["panel"], fg=C["text"]).pack(side="left")
        tk.Button(top, text="Exporter CSV",
                  font=self.f_sub, bg=C["btn_primary"], fg="white",
                  relief="flat", cursor="hand2", padx=10, pady=3,
                  activebackground=C["accent_hover"],
                  command=self._export_csv).pack(side="right")

        cols     = ("date","echantillon","operateur","germe","matrice","volume","resultat","interpretation")
        headings = ("Date/Heure","Échantillon","Opérateur","Germe","Matrice","V ml","N ufc/g","Résultat")
        widths   = (120, 100, 90, 150, 140, 50, 90, 110)

        style = ttk.Style()
        style.configure("Hist.Treeview",         font=(FF, 8), rowheight=22)
        style.configure("Hist.Treeview.Heading", font=(FF, 8, "bold"))
        style.map("Hist.Treeview", background=[("selected", C["accent"])])

        self.tree = ttk.Treeview(panel, columns=cols, show="headings",
                                  height=4, style="Hist.Treeview")
        for col, hdg, w in zip(cols, headings, widths):
            self.tree.heading(col, text=hdg)
            self.tree.column(col, width=w, anchor="center")

        sb = ttk.Scrollbar(panel, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True, padx=(10,0), pady=(0,8))
        sb.pack(side="right", fill="y", padx=(0,8), pady=(0,8))

    # ── Logique métier ────────────────────────────────────────────────────────

    def _parse_fields(self):
        fields, errors = {}, []
        for i in range(NB_DIL):
            vals = []
            for j in range(2):
                raw = self.varfields[i][j].get().strip()
                if raw == "":
                    vals.append(None)
                elif raw.upper() == "NC":
                    vals.append("NC")
                elif raw.lstrip("0").isnumeric() or raw == "0":
                    vals.append(int(raw))
                elif raw.isnumeric():
                    vals.append(int(raw))
                else:
                    errors.append(f"Dil.{i+1} / B{j+1} : « {raw} » invalide")
            fields[i] = vals
        return fields, errors

    def _compute(self, fields):
        d = 0
        for key in fields:
            for v in fields[key]:
                if isinstance(v, int) and 30 <= v <= 300:
                    d = 10 ** (-(key + 1))
                    break
            if d:
                break
        if not d:
            return None, 0, 0, {}

        nb, somme = {}, 0
        for key in fields:
            cnt = sum(1 for v in fields[key] if isinstance(v, int) and 30 <= v <= 300)
            if cnt >= 1 and len(nb) < 2:
                nb[key] = cnt
                somme  += sum(v for v in fields[key] if isinstance(v, int) and 30 <= v <= 300)

        if not nb:
            return None, d, 0, {}

        vals  = list(nb.values())
        n1    = vals[0] if len(vals) > 0 else 0
        n2    = vals[1] if len(vals) > 1 else 0
        denom = d * self.volume.get() * (n1 + 0.1 * n2)
        if denom == 0:
            return None, d, somme, nb

        return somme / denom, d, somme, nb

    def _interpret(self, N):
        germe   = self.germe.get()
        matrice = self.matrice.get()
        seuils  = self.conf.get("seuils", {})
        vals    = seuils.get(germe, {}).get(matrice)
        if not vals:
            return "", None, None

        m, M = vals["m"], vals["M"]
        if N <= m:
            return "✔  Satisfaisant", C["res_sat"], (C["sat_bg"], C["sat_fg"])
        if N <= M:
            return "⚠  Limite", C["res_lim"], (C["lim_bg"], C["lim_fg"])
        return "✘  Insatisfaisant", C["res_insat"], (C["ins_bg"], C["ins_fg"])

    # ── Actions ───────────────────────────────────────────────────────────────

    def calculate(self):
        fields, errors = self._parse_fields()
        if errors:
            messagebox.showerror("Erreur de saisie", "\n".join(errors))
            return

        N, d, somme, nb = self._compute(fields)
        if N is None:
            messagebox.showwarning("Aucune boîte retenue",
                "Aucune boîte ne contient entre 30 et 300 colonies.\n"
                "Indiquez NC pour les non-comptables.")
            return

        self.lbl_result.config(text=f"N = {N:.2e}   ufc/g", fg=C["res_num"])

        interp_text, interp_color, _ = self._interpret(N)
        self.lbl_interp.config(text=interp_text,
                                fg=interp_color if interp_color else C["header_muted"])

        # Détail
        parts = []
        for key, cnt in nb.items():
            vs = " + ".join(str(v) for v in fields[key]
                            if isinstance(v, int) and 30 <= v <= 300)
            parts.append(f"Dil.{key+1} [{vs}]  {cnt} boîte{'s' if cnt>1 else ''}")
        self.lbl_detail.config(
            text=f"Retenues : {'  |  '.join(parts)}    ΣC = {somme}    d = {d:.0e}    V = {self.volume.get()} ml"
        )

        self._update_field_colors(fields)
        self._push_history(N, interp_text)

    def reset(self):
        for i in range(NB_DIL):
            for j in range(2):
                self.varfields[i][j].set("")
                self._set_color(i, j, C["empty"], C["border_dark"])
        self.echantillon.set("")
        self.lbl_result.config(text="— — —", fg=C["res_num"])
        self.lbl_interp.config(text="")
        self.lbl_detail.config(text="")
        self.entry_widgets[0][0].focus_set()

    def _push_history(self, N, interp_text):
        clean = interp_text.replace("✔  ","").replace("⚠  ","").replace("✘  ","") if interp_text else "—"
        row = {
            "date":           datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
            "echantillon":    self.echantillon.get() or "—",
            "operateur":      self.operateur.get() or "—",
            "germe":          self.germe.get(),
            "matrice":        self.matrice.get() or "—",
            "volume":         f"{self.volume.get():.1f}",
            "resultat":       f"{N:.2e}",
            "interpretation": clean,
        }
        self.history.insert(0, row)
        self.tree.insert("", 0, values=tuple(row.values()))

    def _export_csv(self):
        if not self.history:
            messagebox.showinfo("Export CSV", "Aucun résultat à exporter.")
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV", "*.csv")],
            initialfile=f"denombrement_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
        )
        if not path:
            return
        with open(path, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.DictWriter(f, fieldnames=list(self.history[0].keys()), delimiter=";")
            w.writeheader()
            w.writerows(self.history)
        messagebox.showinfo("Export CSV", f"Exporté :\n{path}")

    # ── Validation temps réel ─────────────────────────────────────────────────

    def _on_field_change(self, i, j):
        raw = self.varfields[i][j].get().strip()
        if raw == "":
            bg, bd = C["empty"], C["border_dark"]
        elif raw.upper() == "NC":
            bg, bd = C["nc_bg"], C["nc_border"]
        elif raw.isnumeric():
            n = int(raw)
            if 30 <= n <= 300:
                bg, bd = C["ok_bg"],   C["ok_border"]
            else:
                bg, bd = C["warn_bg"], C["warn_border"]
        else:
            bg, bd = C["err_bg"], C["err_border"]
        self._set_color(i, j, bg, bd)

    def _update_field_colors(self, fields):
        for i in range(NB_DIL):
            for j in range(2):
                v = fields[i][j]
                if v is None:
                    bg, bd = C["empty"],   C["border_dark"]
                elif v == "NC":
                    bg, bd = C["nc_bg"],   C["nc_border"]
                elif isinstance(v, int):
                    if 30 <= v <= 300:
                        bg, bd = C["ok_bg"],   C["ok_border"]
                    else:
                        bg, bd = C["warn_bg"], C["warn_border"]
                else:
                    bg, bd = C["err_bg"], C["err_border"]
                self._set_color(i, j, bg, bd)

    def _set_color(self, i, j, bg, bd):
        self.entry_widgets[i][j].config(bg=bg)
        self.entry_frames[i][j].config(bg=bg, highlightbackground=bd)

    # ── Combobox germe / matrice ──────────────────────────────────────────────

    def _on_germe_change(self):
        germe   = self.germe.get()
        seuils  = self.conf.get("seuils", {})
        matrices = sorted(seuils.get(germe, {}).keys())
        self.cb_matrice.config(values=matrices)
        if matrices:
            self.matrice.set(matrices[0])
        else:
            self.matrice.set("")

    def refresh_germe_list(self):
        seuils = self.conf.get("seuils", DEFAULT_SEUILS)
        self.cb_germe.config(values=["(Aucun seuil)"] + sorted(seuils.keys()))
        self._on_germe_change()

    # ── Utilitaires ───────────────────────────────────────────────────────────

    def _toggle_mode(self, _=None):
        if self.volume.get() == 1.0:
            self.volume.set(0.1)
            self.conf.set("ensemencement", "surface")
        else:
            self.volume.set(1.0)
            self.conf.set("ensemencement", "profondeur")

    def _refresh_mode_badge(self):
        if self.volume.get() == 1.0:
            self.badge_mode.config(text="Profondeur  ·  1 ml")
        else:
            self.badge_mode.config(text="Surface  ·  0.1 ml")

    def _tick_clock(self):
        self.lbl_clock.config(text=datetime.now().strftime("%H:%M"))
        self.after(30_000, self._tick_clock)

    def _tick_date(self):
        self.lbl_date.config(text=datetime.now().strftime("%d %b %Y"))
        self.after(60_000, self._tick_date)


# ── Menu ──────────────────────────────────────────────────────────────────────

def build_menu(root, app, conf):
    bar = tk.Menu(root)

    m_file = tk.Menu(bar, tearoff=0)
    m_file.add_command(label="Exporter CSV…", command=app._export_csv)
    m_file.add_separator()
    m_file.add_command(label="Quitter",
                       command=lambda: _quit(root, conf))
    bar.add_cascade(label="Fichier", menu=m_file)

    m_pref = tk.Menu(bar, tearoff=0)
    m_pref.add_radiobutton(label="Profondeur  (1 ml)",
                            variable=app.volume, value=1.0,
                            command=lambda: conf.set("ensemencement","profondeur"))
    m_pref.add_radiobutton(label="Surface  (0.1 ml)",
                            variable=app.volume, value=0.1,
                            command=lambda: conf.set("ensemencement","surface"))
    m_pref.add_separator()
    m_pref.add_command(label="Seuils de conformité…",
                        command=lambda: SeuilsDialog(root, conf, app.refresh_germe_list))
    bar.add_cascade(label="Préférences", menu=m_pref)

    m_aide = tk.Menu(bar, tearoff=0)
    m_aide.add_command(label="À propos…", command=lambda: _about(root))
    bar.add_cascade(label="Aide", menu=m_aide)

    root.config(menu=bar)


def _quit(root, conf):
    if messagebox.askokcancel("Quitter", "Quitter l'application ?"):
        conf.set("instance", "False")
        root.destroy()


def _about(root):
    w = tk.Toplevel(root)
    w.title("À propos")
    w.resizable(False, False)
    w.geometry("340x200")
    w.configure(bg=C["panel"])
    f_b = tkfont.Font(family=FF, size=13, weight="bold")
    f   = tkfont.Font(family=FF, size=9)
    f_s = tkfont.Font(family=FF, size=8)
    tk.Frame(w, bg=C["header"], height=6).pack(fill="x")
    tk.Label(w, text="DÉNOMBREMENT UFC/g", font=f_b,
             bg=C["panel"], fg=C["text"]).pack(pady=(20,4))
    tk.Label(w, text="Version 3.0  —  NF ISO 7218", font=f,
             bg=C["panel"], fg=C["text_sec"]).pack()
    tk.Label(w, text="Ibrahima Gaye  —  IbsomTech", font=f_s,
             bg=C["panel"], fg=C["text_muted"]).pack(pady=(12,0))


# ── Point d'entrée ────────────────────────────────────────────────────────────

if __name__ == "__main__":
    conf = Config()

    if conf.get("instance") == "True":
        _r = tk.Tk()
        _r.withdraw()
        messagebox.showwarning("Dénombrement", "L'application est déjà ouverte.")
        _r.destroy()
        sys.exit(0)

    conf.set("instance", "True")
    try:
        root = tk.Tk()
        root.title("Dénombrement UFC/g")
        root.configure(bg=C["app_bg"])
        root.minsize(980, 640)

        app_dir   = os.path.dirname(os.path.abspath(sys.argv[0]))
        icon_path = os.path.join(app_dir, "icon.ico")
        if sys.platform == "win32" and os.path.exists(icon_path):
            root.iconbitmap(icon_path)

        app = DenombrementApp(root, conf)
        app.pack(fill="both", expand=True)
        build_menu(root, app, conf)
        root.mainloop()
    finally:
        conf.set("instance", "False")
