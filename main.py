#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel Cleaner (beginner-friendly)
---------------------------------
- Ouvre un fichier Excel (.xlsx/.xls) ou CSV
- Permet de supprimer des colonnes inutiles
- Corrige automatiquement les dates (détection + normalisation YYYY-MM-DD)
- Sauvegarde un fichier propre à côté de l'original: *_clean.xlsx
- Interface simple Tkinter avec sélection de fichier et (si dispo) glisser-déposer

Dépendances:
- pandas
- openpyxl
- (optionnel) tkinterdnd2 pour le drag-and-drop

Auteur: Vous ✨
Licence: MIT
"""
import os
import sys
import traceback
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# Drag & Drop (optionnel)
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD  # pip install tkinterdnd2
    DND_AVAILABLE = True
except Exception:
    DND_AVAILABLE = False

import pandas as pd
from datetime import datetime

APP_TITLE = "Excel Cleaner (Pandas + Tkinter)"
SUPPORTED_EXT = (".xlsx", ".xls", ".csv")

def is_probable_date_series(s, min_ratio=0.6):
    """
    Essaie de détecter si une série ressemble à des dates.
    On tente une conversion 'to_datetime' et on compte le % de valeurs convertibles.
    """
    try:
        converted = pd.to_datetime(s, errors="coerce", dayfirst=True, infer_datetime_format=True)
        ratio = converted.notna().mean()
        return ratio >= min_ratio
    except Exception:
        return False

def normalize_dates(df, selected_date_cols=None):
    """
    Normalise les colonnes de dates au format YYYY-MM-DD.
    Si 'selected_date_cols' est None, on tente de détecter automatiquement.
    Retourne la liste des colonnes traitées.
    """
    handled = []
    if selected_date_cols is None:
        # Détection auto
        candidate_cols = []
        for col in df.columns:
            # On ignore les colonnes numériques pures
            if pd.api.types.is_numeric_dtype(df[col]):
                continue
            # On teste si ça ressemble à des dates
            if is_probable_date_series(df[col]):
                candidate_cols.append(col)
        selected_date_cols = candidate_cols

    for col in selected_date_cols:
        try:
            converted = pd.to_datetime(df[col], errors="coerce", dayfirst=True, infer_datetime_format=True)
            df[col] = converted.dt.strftime("%Y-%m-%d")
            handled.append(col)
        except Exception:
            # On n'arrête pas tout si une colonne pose souci
            pass
    return handled

def setup_theme(root):
    """
    Configure a modern dark theme for the application.
    """
    style = ttk.Style(root)
    
    # Colors
    bg_color = "#2b2b2b"
    fg_color = "#ffffff"
    accent_color = "#4ade80"
    secondary_bg = "#3c3f41"
    select_bg = "#4ade80"
    select_fg = "#000000"
    
    # Configure standard styles
    style.theme_use('clam')
    
    style.configure(".", 
        background=bg_color, 
        foreground=fg_color, 
        fieldbackground=secondary_bg,
        troughcolor=bg_color,
        font=("Segoe UI", 10)
    )
    
    style.configure("TFrame", background=bg_color)
    style.configure("TLabel", background=bg_color, foreground=fg_color)
    style.configure("TButton", 
        background=secondary_bg, 
        foreground=fg_color, 
        borderwidth=0, 
        focuscolor=accent_color
    )
    style.map("TButton", 
        background=[('active', accent_color)], 
        foreground=[('active', '#000000')]
    )
    
    style.configure("TEntry", 
        fieldbackground=secondary_bg, 
        foreground=fg_color,
        insertcolor=fg_color
    )
    
    # Custom styles
    style.configure("Header.TLabel", font=("Segoe UI", 12, "bold"), foreground=accent_color)
    style.configure("Status.TLabel", font=("Segoe UI", 9), foreground="#aaaaaa")

    # Configure root background
    root.configure(bg=bg_color)
    
    return {
        "bg": bg_color,
        "fg": fg_color,
        "secondary": secondary_bg,
        "accent": accent_color
    }

class ExcelCleanerApp:
    def __init__(self, root, colors=None):
        self.root = root
        self.root.title(APP_TITLE)
        self.filepath = None
        self.df = None
        self.colors = colors if colors else {"bg": "#2b2b2b", "fg": "#ffffff", "secondary": "#3c3f41", "accent": "#4ade80"}

        # Style de base
        self.root.geometry("900x650")
        self.root.minsize(800, 600)

        # Frame top: sélection fichier + drag&drop
        top_frame = ttk.Frame(root, padding=20)
        top_frame.pack(fill="x")

        ttk.Label(top_frame, text="Fichier à nettoyer (.xlsx/.xls/.csv) :", style="Header.TLabel").pack(anchor="w", pady=(0, 5))
        self.entry_path = ttk.Entry(top_frame)
        self.entry_path.pack(side="left", fill="x", expand=True, padx=(0,10), ipady=5)

        ttk.Button(top_frame, text="Parcourir...", command=self.select_file).pack(side="left")

        if DND_AVAILABLE:
            # Zone de drop si tkinterdnd2 dispo
            self.drop_label = ttk.Label(top_frame, text="Glissez-déposez un fichier ici", relief="flat", padding=10, background=self.colors["secondary"])
            self.drop_label.pack(side="left", padx=(10,0))
            self.drop_label.drop_target_register(DND_FILES)
            self.drop_label.dnd_bind('<<Drop>>', self.on_drop)
        else:
            self.drop_label = ttk.Label(top_frame, text="(Drag&Drop indisponible)", foreground="#666")
            self.drop_label.pack(side="left", padx=(10,0))

        # Frame milieu: colonnes + actions
        mid_frame = ttk.Frame(root, padding=20)
        mid_frame.pack(fill="both", expand=True)

        # Colonne gauche: liste des colonnes à supprimer
        left = ttk.Frame(mid_frame)
        left.pack(side="left", fill="both", expand=True)

        ttk.Label(left, text="Colonnes détectées (sélectionnez celles à SUPPRIMER):", style="Header.TLabel").pack(anchor="w", pady=(0, 5))
        
        # Custom Listbox style
        self.listbox_cols = tk.Listbox(left, selectmode="multiple", exportselection=False, 
                                       bg=self.colors["secondary"], fg=self.colors["fg"], 
                                       selectbackground=self.colors["accent"], selectforeground=self.colors["select_fg"] if "select_fg" in self.colors else "#000",
                                       borderwidth=0, highlightthickness=1, highlightbackground=self.colors["bg"])
        self.listbox_cols.pack(fill="both", expand=True, pady=(4,10))

        btns_left = ttk.Frame(left)
        btns_left.pack(anchor="w", pady=(0,8))
        ttk.Button(btns_left, text="Tout sélectionner", command=self.select_all_cols).pack(side="left", padx=(0,6))
        ttk.Button(btns_left, text="Tout désélectionner", command=self.clear_selection).pack(side="left")

        # Colonne droite: options dates + prévisualisation
        right = ttk.Frame(mid_frame)
        right.pack(side="left", fill="both", expand=True, padx=(20,0))

        ttk.Label(right, text="Colonnes de dates (optionnel) :", style="Header.TLabel").pack(anchor="w", pady=(0, 5))
        self.listbox_dates = tk.Listbox(right, selectmode="multiple", exportselection=False,
                                        bg=self.colors["secondary"], fg=self.colors["fg"], 
                                        selectbackground=self.colors["accent"], selectforeground=self.colors["select_fg"] if "select_fg" in self.colors else "#000",
                                        borderwidth=0, highlightthickness=1, highlightbackground=self.colors["bg"])
        self.listbox_dates.pack(fill="both", expand=True, pady=(4,10))

        ttk.Label(right, text="Aperçu (5 premières lignes) :", style="Header.TLabel").pack(anchor="w", pady=(10, 5))
        self.text_preview = tk.Text(right, height=12, bg=self.colors["secondary"], fg=self.colors["fg"], borderwidth=0, highlightthickness=0)
        self.text_preview.pack(fill="both", expand=True)

        # Frame bas: actions
        bottom = ttk.Frame(root, padding=20)
        bottom.pack(fill="x")

        self.btn_load = ttk.Button(bottom, text="Charger", command=self.load_file)
        self.btn_load.pack(side="left")

        self.btn_clean = ttk.Button(bottom, text="Nettoyer & Enregistrer", command=self.clean_and_save, state="disabled")
        self.btn_clean.pack(side="left", padx=(10,0))

        self.status = ttk.Label(bottom, text="Prêt.", style="Status.TLabel")

    def on_drop(self, event):
        # Peut contenir des guillemets si le chemin a des espaces
        path = event.data
        if path.startswith("{") and path.endswith("}"):
            path = path[1:-1]
        self.entry_path.delete(0, tk.END)
        self.entry_path.insert(0, path)
        self.load_file()

    def select_file(self):
        path = filedialog.askopenfilename(
            title="Choisir un fichier",
            filetypes=[("Excel/CSV", "*.xlsx *.xls *.csv"), ("Tous les fichiers", "*.*")],
        )
        if path:
            self.entry_path.delete(0, tk.END)
            self.entry_path.insert(0, path)
            self.load_file()

    def load_file(self):
        self.status.config(text="Chargement...")
        self.root.update_idletasks()
        path = self.entry_path.get().strip()
        if not path or not os.path.isfile(path):
            messagebox.showerror("Erreur", "Sélectionnez un fichier valide.")
            self.status.config(text="Erreur: fichier invalide.")
            return
        if not path.lower().endswith(SUPPORTED_EXT):
            messagebox.showerror("Erreur", f"Extension non supportée. Extensions acceptées: {', '.join(SUPPORTED_EXT)}")
            self.status.config(text="Erreur: extension non supportée.")
            return
        try:
            if path.lower().endswith(".csv"):
                df = pd.read_csv(path, encoding="utf-8", sep=None, engine="python")  # autodétection du séparateur
            else:
                df = pd.read_excel(path, engine="openpyxl")
            self.df = df
            self.filepath = path
            self.populate_lists()
            self.show_preview(df)
            self.btn_clean.config(state="normal")
            self.status.config(text=f"Chargé: {os.path.basename(path)}")
        except Exception as e:
            traceback.print_exc()
            messagebox.showerror("Erreur", f"Impossible de charger le fichier:\n{e}")
            self.status.config(text="Erreur de chargement.")

    def populate_lists(self):
        if self.df is None:
            return
        cols = list(self.df.columns)
        self.listbox_cols.delete(0, tk.END)
        self.listbox_dates.delete(0, tk.END)
        for c in cols:
            self.listbox_cols.insert(tk.END, c)

        # Proposer en "dates" les colonnes qui ressemblent à des dates
        for c in cols:
            try:
                if is_probable_date_series(self.df[c]):
                    self.listbox_dates.insert(tk.END, c)
            except Exception:
                # on ignore en cas de problème
                pass

    def show_preview(self, df):
        self.text_preview.delete("1.0", tk.END)
        with pd.option_context("display.max_rows", 5, "display.max_columns", 20, "display.width", 1000):
            self.text_preview.insert(tk.END, str(df.head(5)))

    def get_selected(self, listbox):
        indices = listbox.curselection()
        return [listbox.get(i) for i in indices]

    def select_all_cols(self):
        self.listbox_cols.select_set(0, tk.END)

    def clear_selection(self):
        self.listbox_cols.selection_clear(0, tk.END)
        self.listbox_dates.selection_clear(0, tk.END)

    def clean_and_save(self):
        if self.df is None or not self.filepath:
            messagebox.showerror("Erreur", "Aucun fichier chargé.")
            return

        df = self.df.copy()

        # 1) Supprimer colonnes inutiles
        to_drop = self.get_selected(self.listbox_cols)
        if to_drop:
            for c in to_drop:
                if c in df.columns:
                    df.drop(columns=[c], inplace=True, errors="ignore")

        # 2) Dates: colonnes sélectionnées OU auto-détection si aucune sélection
        selected_dates = self.get_selected(self.listbox_dates)
        handled = normalize_dates(df, selected_date_cols=selected_dates if selected_dates else None)

        # 3) Sauvegarde
        base, ext = os.path.splitext(self.filepath)
        out_path = base + "_clean.xlsx"
        try:
            df.to_excel(out_path, index=False)
            msg = "Fichier propre enregistré:\n" + out_path
            if handled:
                msg += "\n\nColonnes dates normalisées: " + ", ".join(handled)
            messagebox.showinfo("Succès", msg)
            self.status.config(text="Enregistré: " + os.path.basename(out_path))
        except Exception as e:
            traceback.print_exc()
            messagebox.showerror("Erreur", f"Impossible d'enregistrer:\n{e}")
            self.status.config(text="Erreur d'enregistrement.")

def main():
    # Support optionnel TkinterDnD
    if DND_AVAILABLE:
        root = TkinterDnD.Tk()
    else:
        root = tk.Tk()
    
    # Apply Modern Dark Theme
    colors = setup_theme(root)

    app = ExcelCleanerApp(root, colors)
    root.mainloop()

if __name__ == "__main__":
    main()
