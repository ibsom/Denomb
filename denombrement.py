#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Dénombrement UFC/g  v2.0  —  NF ISO 7218
Calcul du nombre de colonies formant unités par gramme (méthode milieu solide).
"""

import tkinter as tk
import tkinter.font as tkfont
from tkinter import ttk, messagebox, filedialog
import json
import os
import sys
import csv
from datetime import datetime

# ── Palette de couleurs ──────────────────────────────────────────────────────
C = {
    "bg":          "#F0F2F5",
    "header":      "#1B2A4A",
    "panel":       "#FFFFFF",
    "border":      "#D1D9E0",
    "accent":      "#2E86C1",
    "green":       "#1E8449",
    "green_bg":    "#D5F5E3",
    "orange_bg":   "#FEF9E7",
    "red":         "#C0392B",
    "red_bg":      "#FDEDEC",
    "nc_bg":       "#EBF5FB",
    "white":       "#FFFFFF",
    "text":        "#1B2A4A",
    "muted":       "#7F8C8D",
    "result_bg":   "#1B2A4A",
    "result_val":  "#2ECC71",
}
FF = "Segoe UI"

# ── Seuils réglementaires (ufc/g) par germe ──────────────────────────────────
# m = seuil de qualité acceptable, M = seuil d'action (au-delà = insatisfaisant)
SEUILS = {
    "(Aucun seuil)":           {"m": None, "M": None},
    "Flore mésophile totale":  {"m": 1e5,  "M": 1e7},
    "Entérobactéries":         {"m": 1e2,  "M": 1e4},
    "Staphylocoques coag. +":  {"m": 1e2,  "M": 1e3},
    "Levures / Moisissures":   {"m": 1e3,  "M": 1e5},
    "Coliformes totaux":       {"m": 1e2,  "M": 1e4},
    "Coliformes thermo.":      {"m": 10,   "M": 1e2},
    "Listeria spp.":           {"m": 0,    "M": 1e2},
    "Salmonella":              {"m": 0,    "M": 0},
}

NB_DIL = 6


# ── Configuration persistante ────────────────────────────────────────────────

class Config:
    _DEFAULTS = {
        "instance":      "False",
        "ensemencement": "profondeur",
        "operateur":     "",
    }

    def __init__(self):
        self.path = os.path.join(
            os.environ.get("USERPROFILE", os.path.expanduser("~")),
            "AppData", "Local", "denombrement", "conf.json",
        )
        try:
            with open(self.path) as f:
                self.data = {**self._DEFAULTS, **json.load(f)}
        except (FileNotFoundError, json.JSONDecodeError):
            os.makedirs(os.path.dirname(self.path), exist_ok=True)
            self.data = self._DEFAULTS.copy()
            self._write()

    def _write(self):
        with open(self.path, "w") as f:
            json.dump(self.data, f)

    def set(self, key, value):
        self.data[key] = value
        self._write()

    def get(self, key, default=None):
        return self.data.get(key, default)


# ── Application principale ───────────────────────────────────────────────────

class DenombrementApp(tk.Frame):

    def __init__(self, root, conf):
        super().__init__(root, bg=C["bg"])
        self.root    = root
        self.conf    = conf
        self.history = []

        self._build_fonts()
        self._build_vars()
        self._build_ui()
        self._apply_prefs()

    # ── Polices ─────────────────────────────────────────────────────────────

    def _build_fonts(self):
        self.f_title  = tkfont.Font(family=FF, size=15, weight="bold")
        self.f_label  = tkfont.Font(family=FF, size=9)
        self.f_entry  = tkfont.Font(family="Consolas", size=13)
        self.f_btn    = tkfont.Font(family=FF, size=11, weight="bold")
        self.f_result = tkfont.Font(family=FF, size=24, weight="bold")
        self.f_interp = tkfont.Font(family=FF, size=11, weight="bold")
        self.f_detail = tkfont.Font(family=FF, size=9)
        self.f_small  = tkfont.Font(family=FF, size=8)
        self.f_dil    = tkfont.Font(family=FF, size=8, slant="italic")

    # ── Variables ────────────────────────────────────────────────────────────

    def _build_vars(self):
        self.varfields   = [[tk.StringVar() for _ in range(2)] for _ in range(NB_DIL)]
        self.volume      = tk.DoubleVar(value=1.0)
        self.echantillon = tk.StringVar()
        self.operateur   = tk.StringVar(value=self.conf.get("operateur", ""))
        self.germe       = tk.StringVar(value="(Aucun seuil)")

        for i in range(NB_DIL):
            for j in range(2):
                self.varfields[i][j].trace_add(
                    "write", lambda *a, i=i, j=j: self._on_field_change(i, j)
                )
        self.operateur.trace_add(
            "write", lambda *a: self.conf.set("operateur", self.operateur.get())
        )

    def _apply_prefs(self):
        if self.conf.get("ensemencement") == "surface":
            self.volume.set(0.1)

    # ── Construction de l'interface ──────────────────────────────────────────

    def _build_ui(self):
        self._build_header()
        self._build_sample_bar()
        self._build_dilution_grid()
        self._build_buttons()
        self._build_result_panel()
        self._build_history_panel()

    def _build_header(self):
        hdr = tk.Frame(self, bg=C["header"], pady=10)
        hdr.pack(fill="x")

        tk.Label(hdr, text="DÉNOMBREMENT UFC/g",
                 font=self.f_title, bg=C["header"], fg=C["white"]).pack(side="left", padx=16)
        tk.Label(hdr, text="NF ISO 7218  •  v2.0",
                 font=self.f_small, bg=C["header"], fg="#7FB3D3").pack(side="left")

        self.lbl_mode = tk.Label(hdr, text="", font=self.f_small,
                                  bg=C["accent"], fg="white", padx=10, pady=4, cursor="hand2")
        self.lbl_mode.pack(side="right", padx=14)
        self.lbl_mode.bind("<Button-1>", self._toggle_mode)
        self.volume.trace_add("write", lambda *a: self._refresh_mode_label())
        self._refresh_mode_label()

    def _build_sample_bar(self):
        bar = tk.Frame(self, bg=C["panel"], pady=8, padx=14,
                       highlightbackground=C["border"], highlightthickness=1)
        bar.pack(fill="x", padx=10, pady=(8, 0))

        def lbl(text, col):
            tk.Label(bar, text=text, font=self.f_label,
                     bg=C["panel"], fg=C["muted"]).grid(row=0, column=col, sticky="w", padx=(8 if col else 0, 0))

        def field(var, col, w=18):
            e = tk.Entry(bar, textvariable=var, font=self.f_label,
                         width=w, bd=1, relief="solid", bg=C["white"])
            e.grid(row=0, column=col, padx=(4, 12), sticky="w")
            return e

        lbl("Échantillon :", 0)
        field(self.echantillon, 1, w=16)
        lbl("Opérateur :", 2)
        field(self.operateur, 3, w=14)
        lbl("Germe :", 4)

        cb = ttk.Combobox(bar, textvariable=self.germe,
                          values=list(SEUILS.keys()), width=22,
                          state="readonly", font=self.f_label)
        cb.grid(row=0, column=5, padx=(4, 12), sticky="w")

        self.lbl_date = tk.Label(bar, font=self.f_label, bg=C["panel"], fg=C["muted"])
        self.lbl_date.grid(row=0, column=6, sticky="e")
        bar.columnconfigure(6, weight=1)
        self._tick_clock()

    def _build_dilution_grid(self):
        wrapper = tk.Frame(self, bg=C["bg"])
        wrapper.pack(fill="x", padx=10, pady=(10, 0))

        self.entry_widgets = []
        self.entry_frames  = []

        for i in range(NB_DIL):
            col = tk.Frame(wrapper, bg=C["panel"],
                           highlightbackground=C["border"], highlightthickness=1)
            col.grid(row=0, column=i, padx=5, sticky="nsew")
            wrapper.columnconfigure(i, weight=1)

            # En-tête
            hdr = tk.Frame(col, bg=C["header"], pady=5)
            hdr.pack(fill="x")
            tk.Label(hdr, text=f"Dilution {i+1}", font=self.f_label,
                     bg=C["header"], fg="white").pack()
            tk.Label(hdr, text=f"10⁻{i+1}", font=self.f_dil,
                     bg=C["header"], fg="#7FB3D3").pack()

            row_w, row_f = [], []
            for j in range(2):
                body = tk.Frame(col, bg=C["panel"])
                body.pack(fill="x", padx=6, pady=(5, 4))
                tk.Label(body, text=f"Boîte {j+1}", font=self.f_small,
                         bg=C["panel"], fg=C["muted"]).pack()
                ef = tk.Frame(body, bg=C["white"], bd=1, relief="solid")
                ef.pack(fill="x")
                e = tk.Entry(ef, textvariable=self.varfields[i][j],
                             font=self.f_entry, width=7, bd=0,
                             bg=C["white"], justify="center",
                             insertbackground=C["text"])
                e.pack(padx=2, pady=4)
                row_w.append(e)
                row_f.append(ef)

            self.entry_widgets.append(row_w)
            self.entry_frames.append(row_f)

        self._setup_keyboard_nav()

    def _setup_keyboard_nav(self):
        flat = [self.entry_widgets[i][j] for i in range(NB_DIL) for j in range(2)]
        for k, e in enumerate(flat):
            nxt = flat[(k + 1) % len(flat)]
            e.bind("<Return>", lambda ev, n=nxt: n.focus_set())
        flat[-1].bind("<Return>", lambda ev: self.calculate())

    def _build_buttons(self):
        frm = tk.Frame(self, bg=C["bg"])
        frm.pack(pady=12)

        tk.Button(frm, text="⟳  RÉINITIALISER",
                  font=self.f_btn, bg=C["red"], fg="white",
                  width=17, relief="flat", cursor="hand2",
                  activebackground="#A93226", activeforeground="white",
                  command=self.reset).pack(side="left", padx=10)

        tk.Button(frm, text="▶  CALCULER  [F5]",
                  font=self.f_btn, bg=C["green"], fg="white",
                  width=17, relief="flat", cursor="hand2",
                  activebackground="#186A3B", activeforeground="white",
                  command=self.calculate).pack(side="left", padx=10)

        self.root.bind("<F5>",     lambda e: self.calculate())
        self.root.bind("<Escape>", lambda e: self.reset())

    def _build_result_panel(self):
        outer = tk.Frame(self, bg=C["bg"])
        outer.pack(fill="x", padx=10, pady=(0, 8))

        panel = tk.Frame(outer, bg=C["result_bg"], pady=16)
        panel.pack(fill="x")

        tk.Label(panel, text="RÉSULTAT", font=self.f_small,
                 bg=C["result_bg"], fg="#7FB3D3").pack()

        self.lbl_result = tk.Label(panel, text="— — —",
                                    font=self.f_result,
                                    bg=C["result_bg"], fg=C["white"])
        self.lbl_result.pack()

        self.lbl_interp = tk.Label(panel, text="",
                                    font=self.f_interp,
                                    bg=C["result_bg"], fg=C["white"])
        self.lbl_interp.pack(pady=(4, 0))

        self.lbl_detail = tk.Label(panel, text="",
                                    font=self.f_detail,
                                    bg=C["result_bg"], fg="#7FB3D3",
                                    wraplength=860)
        self.lbl_detail.pack(pady=(4, 0))

    def _build_history_panel(self):
        panel = tk.Frame(self, bg=C["panel"],
                         highlightbackground=C["border"], highlightthickness=1)
        panel.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        top = tk.Frame(panel, bg=C["panel"])
        top.pack(fill="x", padx=10, pady=(6, 2))
        tk.Label(top, text="Historique de session",
                 font=self.f_label, bg=C["panel"], fg=C["text"]).pack(side="left")
        tk.Button(top, text="Exporter CSV",
                  font=self.f_small, bg=C["accent"], fg="white",
                  relief="flat", cursor="hand2",
                  activebackground="#1A5276", activeforeground="white",
                  command=self._export_csv).pack(side="right", padx=4)

        cols = ("date", "echantillon", "operateur", "germe", "volume", "resultat", "interpretation")
        widths = (120, 110, 100, 150, 75, 100, 120)
        headings = ("Date/Heure", "Échantillon", "Opérateur", "Germe",
                    "Volume ml", "N (ufc/g)", "Interprétation")

        self.tree = ttk.Treeview(panel, columns=cols, show="headings", height=4)
        for col, hdg, w in zip(cols, headings, widths):
            self.tree.heading(col, text=hdg)
            self.tree.column(col, width=w, anchor="center")

        style = ttk.Style()
        style.configure("Treeview",         font=(FF, 9),         rowheight=22)
        style.configure("Treeview.Heading", font=(FF, 9, "bold"))

        sb = ttk.Scrollbar(panel, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True, padx=(10, 0), pady=(0, 8))
        sb.pack(side="right", fill="y", padx=(0, 8), pady=(0, 8))

    # ── Logique métier ──────────────────────────────────────────────────────

    def _parse_fields(self):
        """Lit et valide tous les champs. Retourne (fields_values, errors)."""
        fields_values = {}
        errors = []
        for i in range(NB_DIL):
            values = []
            for j in range(2):
                raw = self.varfields[i][j].get().strip()
                if raw == "":
                    values.append(None)
                elif raw.upper() == "NC":
                    values.append("NC")
                elif raw.lstrip("0").isnumeric() or raw == "0":
                    values.append(int(raw))
                elif raw.isnumeric():
                    values.append(int(raw))
                else:
                    errors.append(f"Dilution {i+1} / Boîte {j+1} : « {raw} » invalide")
            fields_values[i] = values
        return fields_values, errors

    def _compute(self, fields_values):
        """
        Calcule N selon NF ISO 7218 :
          N = ΣC / (d × V × (n1 + 0.1·n2))
        Retourne (N, d, somme, nb_retenues_dict) ou (None, ...) si pas de boîtes valides.
        """
        # Trouver le taux de dilution de la première dilution avec boîtes dans [30,300]
        d = 0
        for key in fields_values:
            for v in fields_values[key]:
                if isinstance(v, int) and 30 <= v <= 300:
                    d = 10 ** (-(key + 1))
                    break
            if d:
                break

        if not d:
            return None, 0, 0, {}

        # Collecter les 2 niveaux de dilution avec boîtes retenues (30–300)
        nb_retenues = {}
        somme = 0
        for key in fields_values:
            count = sum(1 for v in fields_values[key] if isinstance(v, int) and 30 <= v <= 300)
            if count >= 1 and len(nb_retenues) < 2:
                nb_retenues[key] = count
                somme += sum(v for v in fields_values[key] if isinstance(v, int) and 30 <= v <= 300)

        if not nb_retenues:
            return None, d, 0, {}

        vals = list(nb_retenues.values())
        n1   = vals[0] if len(vals) > 0 else 0
        n2   = vals[1] if len(vals) > 1 else 0

        denom = d * self.volume.get() * (n1 + 0.1 * n2)
        if denom == 0:
            return None, d, somme, nb_retenues

        return somme / denom, d, somme, nb_retenues

    def _interpret(self, N):
        """Retourne (texte, couleur_fg) selon seuils m/M du germe sélectionné."""
        seuil = SEUILS.get(self.germe.get(), {"m": None, "M": None})
        m, M = seuil["m"], seuil["M"]
        if m is None:
            return "", C["white"]
        if N <= m:
            return "✔  Satisfaisant  (N ≤ m)", C["green_bg"]
        if N <= M:
            return "⚠  Limite  (m < N ≤ M)", "#F9E79F"
        return "✘  Insatisfaisant  (N > M)", C["red_bg"]

    # ── Actions utilisateur ──────────────────────────────────────────────────

    def calculate(self):
        fields_values, errors = self._parse_fields()
        if errors:
            messagebox.showerror("Erreur de saisie", "\n".join(errors))
            return

        N, d, somme, nb_retenues = self._compute(fields_values)

        if N is None:
            messagebox.showwarning(
                "Aucune boîte retenue",
                "Aucune boîte ne contient entre 30 et 300 colonies.\n"
                "Vérifiez vos valeurs ou indiquez NC pour les non comptables.",
            )
            return

        # Résultat
        self.lbl_result.config(text=f"N = {N:.2e}   ufc/g")

        # Interprétation réglementaire
        interp_text, interp_color = self._interpret(N)
        self.lbl_interp.config(text=interp_text, fg=interp_color)

        # Détail des boîtes retenues
        parts = []
        for key, count in nb_retenues.items():
            vals_str = " + ".join(
                str(v) for v in fields_values[key]
                if isinstance(v, int) and 30 <= v <= 300
            )
            parts.append(f"Dil.{key+1} [{vals_str}] ({count} boîte{'s' if count > 1 else ''})")
        detail = (
            f"Boîtes retenues : {'  |  '.join(parts)}"
            f"     ΣC = {somme}   d = {d:.0e}   V = {self.volume.get()} ml"
        )
        self.lbl_detail.config(text=detail)

        # Coloriser les champs selon leur statut
        self._update_field_colors(fields_values)

        # Historique
        self._push_history(N, interp_text)

    def reset(self):
        for i in range(NB_DIL):
            for j in range(2):
                self.varfields[i][j].set("")
                self._set_entry_color(i, j, C["white"])
        self.echantillon.set("")
        self.lbl_result.config(text="— — —")
        self.lbl_interp.config(text="")
        self.lbl_detail.config(text="")
        self.entry_widgets[0][0].focus_set()

    # ── Historique & export ──────────────────────────────────────────────────

    def _push_history(self, N, interp_text):
        clean = interp_text.replace("✔  ", "").replace("⚠  ", "").replace("✘  ", "").split("  (")[0]
        row = {
            "date":           datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
            "echantillon":    self.echantillon.get() or "—",
            "operateur":      self.operateur.get() or "—",
            "germe":          self.germe.get(),
            "volume":         f"{self.volume.get():.1f}",
            "resultat":       f"{N:.2e}",
            "interpretation": clean or "—",
        }
        self.history.insert(0, row)
        self.tree.insert("", 0, values=tuple(row.values()))

    def _export_csv(self):
        if not self.history:
            messagebox.showinfo("Export CSV", "Aucun résultat à exporter.")
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("Fichier CSV", "*.csv")],
            initialfile=f"denombrement_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
        )
        if not path:
            return
        with open(path, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.DictWriter(f, fieldnames=list(self.history[0].keys()), delimiter=";")
            w.writeheader()
            w.writerows(self.history)
        messagebox.showinfo("Export CSV", f"Exporté :\n{path}")

    # ── Validation temps réel ────────────────────────────────────────────────

    def _on_field_change(self, i, j):
        raw = self.varfields[i][j].get().strip()
        if raw == "":
            color = C["white"]
        elif raw.upper() == "NC":
            color = C["nc_bg"]
        elif raw.isnumeric():
            color = C["green_bg"] if 30 <= int(raw) <= 300 else C["orange_bg"]
        else:
            color = C["red_bg"]
        self._set_entry_color(i, j, color)

    def _update_field_colors(self, fields_values):
        for i in range(NB_DIL):
            for j in range(2):
                v = fields_values[i][j]
                if v is None:
                    color = C["white"]
                elif v == "NC":
                    color = C["nc_bg"]
                elif isinstance(v, int):
                    color = C["green_bg"] if 30 <= v <= 300 else C["orange_bg"]
                else:
                    color = C["red_bg"]
                self._set_entry_color(i, j, color)

    def _set_entry_color(self, i, j, color):
        self.entry_widgets[i][j].config(bg=color)
        self.entry_frames[i][j].config(bg=color)

    # ── Utilitaires ──────────────────────────────────────────────────────────

    def _toggle_mode(self, _event=None):
        if self.volume.get() == 1.0:
            self.volume.set(0.1)
            self.conf.set("ensemencement", "surface")
        else:
            self.volume.set(1.0)
            self.conf.set("ensemencement", "profondeur")

    def _refresh_mode_label(self):
        if self.volume.get() == 1.0:
            self.lbl_mode.config(text="⬤  Profondeur  1 ml")
        else:
            self.lbl_mode.config(text="⬤  Surface  0.1 ml")

    def _tick_clock(self):
        self.lbl_date.config(text=datetime.now().strftime("%d/%m/%Y  %H:%M"))
        self.after(30_000, self._tick_clock)


# ── Menu ─────────────────────────────────────────────────────────────────────

def build_menu(root, app, conf):
    bar = tk.Menu(root)

    m_file = tk.Menu(bar, tearoff=0)
    m_file.add_command(label="Exporter CSV…", command=app._export_csv)
    m_file.add_separator()
    m_file.add_command(label="Quitter", command=lambda: _quit(root, conf))
    bar.add_cascade(label="Fichier", menu=m_file)

    m_pref = tk.Menu(bar, tearoff=0)
    m_pref.add_radiobutton(
        label="Ensemencement en profondeur (1 ml)",
        variable=app.volume, value=1.0,
        command=lambda: conf.set("ensemencement", "profondeur"),
    )
    m_pref.add_radiobutton(
        label="Ensemencement en surface (0.1 ml)",
        variable=app.volume, value=0.1,
        command=lambda: conf.set("ensemencement", "surface"),
    )
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
    w.geometry("340x190")
    w.configure(bg=C["panel"])
    tk.Label(w, text="DÉNOMBREMENT UFC/g",
             font=tkfont.Font(family=FF, size=14, weight="bold"),
             bg=C["panel"], fg=C["text"]).pack(pady=(24, 4))
    tk.Label(w, text="Version 2.0  —  NF ISO 7218",
             font=tkfont.Font(family=FF, size=10),
             bg=C["panel"], fg=C["text"]).pack()
    tk.Label(w, text="Ibrahima Gaye  —  IbsomTech",
             font=tkfont.Font(family=FF, size=9),
             bg=C["panel"], fg=C["muted"]).pack(pady=(14, 0))


# ── Point d'entrée ───────────────────────────────────────────────────────────

if __name__ == "__main__":
    conf = Config()

    # Protection multi-instance
    if conf.get("instance") == "True":
        root_tmp = tk.Tk()
        root_tmp.withdraw()
        messagebox.showwarning("Dénombrement", "L'application est déjà ouverte.")
        root_tmp.destroy()
        sys.exit(0)

    conf.set("instance", "True")
    try:
        root = tk.Tk()
        root.title("Dénombrement UFC/g")
        root.configure(bg=C["bg"])
        root.minsize(950, 620)

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
