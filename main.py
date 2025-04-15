"""
Application de gestion des contrats - Point d'entrée principal
Ce module centralise l'accès à toutes les fonctionnalités de l'application
et fournit une interface utilisateur unifiée.
"""

import os
import sys
import tkinter as tk
from tkinter import ttk, filedialog
import threading
import pandas as pd  

# Ajouter le répertoire parent (racine du projet) au sys.path
PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.append(PROJECT_ROOT)

# Import des modules personnalisés
from core.ui_styles import (
    COLORS, FONTS, configure_ttk_style, create_button, 
    create_label, create_frame
)
from core.utils import (
    logger, show_error, show_info, show_success, show_warning,
    ask_question, load_settings, save_settings, get_file_path,
    update_file_path, ensure_directory_exists, center_window
)
from widgets.widgets import (
    HeaderFrame, FooterFrame, TabView, StatusBar
)

# Import des modules fonctionnels
from modules.invoices import analyse_facture
from modules.payslips import bulletins
from modules.contracts import contrat53_new as contrat53

class MainApplication(tk.Tk):
    """
    Application principale qui intègre tous les modules.
    """
    def __init__(self):
        """
        Initialise l'application principale.
        """
        super().__init__()
        
        # Configuration de la fenêtre principale
        self.title("Gestion des contrats - SELARL Anesthésistes Mathilde")
        self.geometry("1200x800")
        self.minsize(800, 600)
        
        # Configurer le style ttk
        self.style = configure_ttk_style()
        
        # Vérifier les chemins de fichiers
        self.check_file_paths()
        
        # Créer l'interface
        self.create_widgets()
        
        # Centrer la fenêtre
        center_window(self, 1200, 800)
        
        # Initialiser l'interface d'analyse des factures après la création complète
        self.after(100, self.initialize_interfaces)
        
        # Journalisation
        logger.info("Application démarrée")
    
    def initialize_interfaces(self):
        """Initialise les interfaces qui nécessitent que l'application soit complètement démarrée"""
        try:
            # Vérifiez d'abord que le dossier des factures existe
            factures_path = get_file_path("dossier_factures", verify_exists=True, create_if_missing=True)
            if factures_path:
                self.open_invoice_analysis()
            else:
                show_warning("Dossier manquant", "Le dossier des factures n'existe pas. Veuillez le configurer dans les paramètres.")
        except Exception as e:
            logger.error(f"Erreur lors de l'initialisation des interfaces: {str(e)}")
            show_error("Erreur", f"Impossible d'initialiser l'interface d'analyse des factures: {str(e)}")
                
    def check_file_paths(self):
        """
        Vérifie que tous les chemins de fichiers nécessaires sont définis.
        """
        settings = load_settings()
        
        # Liste des chemins requis
        required_paths = {
            "dossier_factures": "Dossier des factures",
            "bulletins_salaire": "Dossier des bulletins de salaire",
            "excel_mar": "Fichier Excel MAR",
            "excel_iade": "Fichier Excel IADE",
            "word_mar": "Modèle Word MAR",
            "word_iade": "Modèle Word IADE",
            "pdf_mar": "Dossier PDF Contrats MAR",
            "pdf_iade": "Dossier PDF Contrats IADE",
            "excel_salaries": "Fichier Excel Salariés"
        }
        
        # Vérifier chaque chemin
        missing_paths = []
        for key, description in required_paths.items():
            if key not in settings or not settings[key]:
                missing_paths.append(description)
        
        # Si des chemins sont manquants, afficher un avertissement
        if missing_paths:
            message = "Certains chemins de fichiers ne sont pas définis :\n"
            message += "\n".join([f"- {path}" for path in missing_paths])
            message += "\n\nVeuillez les configurer dans les paramètres."
            
            show_warning("Configuration incomplète", message)
    
    def create_widgets(self):
        """
        Crée les widgets de l'interface principale.
        """
        # Cadre principal
        self.main_frame = tk.Frame(self, bg=COLORS["secondary"])
        self.main_frame.pack(fill="both", expand=True)
        
        # En-tête
        self.header = HeaderFrame(
            self.main_frame, 
            "Gestion des contrats - SELARL Anesthésistes Mathilde",
            "Système de gestion intégré"
        )
        self.header.pack(fill="x")
        
        # Barre de statut (créée avant les onglets pour être disponible)
        self.status_bar = StatusBar(self.main_frame)
        self.status_bar.pack(fill="x", side="bottom")
        self.status_bar.set_message("Prêt")
        
        # Onglets
        self.tabs = TabView(self.main_frame)
        self.tabs.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Onglet Accueil
        self.home_tab = self.tabs.add_tab("Accueil", bg=COLORS["secondary"])
        self.create_home_tab()
        
        # Onglet Contrats
        self.contracts_tab = self.tabs.add_tab("Contrats", bg=COLORS["secondary"])
        self.create_contracts_tab()
        
        # Onglet Comptabilité (remplace "Factures" et "Bulletins")
        self.accounting_tab = self.tabs.add_tab("Comptabilité", bg=COLORS["secondary"])
        self.create_accounting_tab()
        
        # Onglet Paramètres
        self.settings_tab = self.tabs.add_tab("Paramètres", bg=COLORS["secondary"])
        self.create_settings_tab()
    
    def create_home_tab(self):
        """
        Crée le contenu de l'onglet Accueil.
        """
        # Titre
        create_label(
            self.home_tab, 
            "Bienvenue dans l'application de gestion des contrats",
            style="subtitle"
        ).pack(pady=20)
        
        # Description
        description = (
            "Cette application centralise la gestion des contrats, factures et bulletins de salaire "
            "pour la SELARL des Anesthésistes de la Clinique Mathilde.\n\n"
            "Utilisez les onglets ci-dessus pour accéder aux différentes fonctionnalités."
        )
        
        create_label(
            self.home_tab, 
            description,
            style="body"
        ).pack(pady=10, padx=50)
        
        # Cadre pour les raccourcis
        shortcuts_frame = create_frame(self.home_tab)
        shortcuts_frame.pack(pady=30, fill="x")
        
        # Titre des raccourcis
        create_label(
            shortcuts_frame, 
            "Accès rapide",
            style="subtitle"
        ).pack(pady=10)
        
        # Grille de boutons
        buttons_frame = create_frame(shortcuts_frame)
        buttons_frame.pack(pady=10)
        
        # Fonction locale pour créer des commandes sécurisées
        def make_tab_selector(tab_name):
            return lambda: self.tabs.select_tab(tab_name)
        
        # Boutons de raccourci
        shortcuts = [
            ("Nouveau contrat MAR", make_tab_selector("Contrats"), "primary"),
            ("Nouveau contrat IADE", make_tab_selector("Contrats"), "primary"),
            ("Analyser factures", make_tab_selector("Comptabilité"), "info"),
            ("Consulter bulletins", make_tab_selector("Comptabilité"), "info"),
            ("Paramètres", make_tab_selector("Paramètres"), "neutral")
        ]
        
        for i, (text, command, style) in enumerate(shortcuts):
            create_button(
                buttons_frame, 
                text=text, 
                command=command,
                style=style,
                width=20,
                height=2
            ).grid(row=i//3, column=i%3, padx=10, pady=10)
    
    def create_contracts_tab(self):
        """
        Crée le contenu de l'onglet Contrats.
        """
        # Supprimer les sous-onglets et utiliser directement un frame unique
        contracts_frame = create_frame(self.contracts_tab, bg=COLORS["secondary"])
        contracts_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # Titre de section élégant
        title_label = tk.Label(
            contracts_frame, 
            text="Gestion des contrats", 
            font=("Arial", 16, "bold"),
            bg=COLORS["primary"],
            fg="white",
            padx=10,
            pady=8
        )
        title_label.pack(fill="x", pady=(0, 20))
        
        # Section MAR avec un style amélioré
        mar_section = create_frame(contracts_frame, bg=COLORS["secondary"])
        mar_section.pack(fill="x", pady=10)
        
        tk.Label(
            mar_section, 
            text="Contrats MAR (Médecins Anesthésistes Réanimateurs)", 
            font=("Arial", 12, "bold"),
            bg=COLORS["info"],
            fg="white",
            padx=10,
            pady=5
        ).pack(fill="x", pady=(0, 10))
        
        # Boutons MAR avec style amélioré
        button_frame_mar = create_frame(mar_section, bg=COLORS["secondary"])
        button_frame_mar.pack(pady=10)
        
        create_button(
            button_frame_mar, 
            text="Nouveau contrat MAR", 
            command=self.open_contract_creation_mar,
            style="primary",
            width=25,
            height=2
        ).pack(side="left", padx=10, pady=10)
        
        create_button(
            button_frame_mar, 
            text="Gérer les contrats MAR existants", 
            command=self.manage_mar_contracts,
            style="info",
            width=25,
            height=2
        ).pack(side="left", padx=10, pady=10)
        
        # Séparateur visuel
        ttk.Separator(contracts_frame, orient="horizontal").pack(fill="x", pady=20)
        
        # Section IADE avec un style amélioré
        iade_section = create_frame(contracts_frame, bg=COLORS["secondary"])
        iade_section.pack(fill="x", pady=10)
        
        tk.Label(
            iade_section, 
            text="Contrats IADE (Infirmiers Anesthésistes Diplômés d'État)", 
            font=("Arial", 12, "bold"),
            bg=COLORS["accent"],
            fg="white",
            padx=10,
            pady=5
        ).pack(fill="x", pady=(0, 10))
        
        # Boutons IADE avec style amélioré
        button_frame_iade = create_frame(iade_section, bg=COLORS["secondary"])
        button_frame_iade.pack(pady=10)
        
        create_button(
            button_frame_iade, 
            text="Nouveau contrat IADE", 
            command=self.open_contract_creation_iade,
            style="accent",
            width=25,
            height=2
        ).pack(side="left", padx=10, pady=10)
        
        create_button(
            button_frame_iade, 
            text="Gérer les contrats IADE existants", 
            command=self.manage_iade_contracts,
            style="info",
            width=25,
            height=2
        ).pack(side="left", padx=10, pady=10)
    
    def create_accounting_tab(self):
        """
        Crée le contenu de l'onglet Comptabilité avec des sous-onglets.
        """
        # Créer les sous-onglets pour la comptabilité
        accounting_tabs = TabView(self.accounting_tab)
        accounting_tabs.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Sous-onglet Frais et factures
        invoices_tab = accounting_tabs.add_tab("Frais et factures", bg=COLORS["secondary"])
        
        # Cadre pour l'analyse des factures (occupe tout l'espace)
        self.invoices_frame = create_frame(invoices_tab)
        self.invoices_frame.pack(fill="both", expand=True, padx=0, pady=0)
        
        # N'initialise pas ici - maintenant fait dans initialize_interfaces
        # self.open_invoice_analysis()
        
        # Sous-onglet Bulletins de salaire
        bulletins_tab = accounting_tabs.add_tab("Bulletins de salaire", bg=COLORS["secondary"])
        
        # Cadre pour les bulletins
        self.bulletins_frame = create_frame(bulletins_tab)
        self.bulletins_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Boutons d'action pour les bulletins
        buttons_frame = create_frame(bulletins_tab)
        buttons_frame.pack(pady=10)
        
        create_button(
            buttons_frame, 
            text="Consulter les bulletins", 
            command=lambda: self.open_bulletins(),
            style="primary",
            width=25
        ).pack(side="left", padx=10)
        
        create_button(
            buttons_frame, 
            text="Scanner un nouveau bulletin", 
            command=lambda: self.scan_new_bulletin(),
            style="info",
            width=25
        ).pack(side="left", padx=10)
        
        # Sous-onglet Virements
        transfers_tab = accounting_tabs.add_tab("Virements", bg=COLORS["secondary"])
        
        # Boutons pour les virements
        create_button(
            transfers_tab, 
            text="Effectuer un virement", 
            command=lambda: self.open_transfer(),
            style="primary",
            width=25,
            height=2
        ).pack(pady=20)
        
        create_button(
            transfers_tab, 
            text="Virement rémunération MARS", 
            command=lambda: self.open_mars_transfer(),
            style="info",
            width=25,
            height=2
        ).pack(pady=10)
    
    def create_settings_tab(self):
        """
        Crée le contenu de l'onglet Paramètres.
        """
        # Sous-onglets pour les différents types de paramètres
        settings_tabs = TabView(self.settings_tab)
        settings_tabs.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Sous-onglet Chemins
        paths_tab = settings_tabs.add_tab("Chemins", bg=COLORS["secondary"])
        
        # Au lieu d'appeler directement, on utilise lambda
        self.paths_tab_setup(paths_tab)
        
        # Sous-onglet Médecins
        doctors_tab = settings_tabs.add_tab("Médecins", bg=COLORS["secondary"])
        self.doctors_tab_setup(doctors_tab)
        
        # Sous-onglet IADE
        iade_tab = settings_tabs.add_tab("IADE", bg=COLORS["secondary"])
        self.iade_tab_setup(iade_tab)
        
        # Sous-onglet Salariés
        employees_tab = settings_tabs.add_tab("Salariés", bg=COLORS["secondary"])
        self.employees_tab_setup(employees_tab)
        
        # Sous-onglet DocuSign
        docusign_tab = settings_tabs.add_tab("DocuSign", bg=COLORS["secondary"])
        self.docusign_tab_setup(docusign_tab)
    
    # Renommer les fonctions pour éviter les conflits
    def paths_tab_setup(self, parent):
        """
        Crée le contenu du sous-onglet Chemins.
        """
        # Titre
        create_label(
            parent, 
            "Configuration des chemins de fichiers",
            style="subtitle"
        ).pack(pady=10)
        
        # Cadre avec scrollbar
        canvas = tk.Canvas(parent, bg=COLORS["secondary"])
        scrollbar = ttk.Scrollbar(parent, orient="vertical", command=canvas.yview)
        scrollable_frame = create_frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True, padx=10, pady=10)
        scrollbar.pack(side="right", fill="y")
        
        # Charger les paramètres
        settings = load_settings()
        
        # Liste des chemins à configurer
        paths = [
            ("Chemin Excel MAR", "excel_mar", "file"),
            ("Chemin Excel IADE", "excel_iade", "file"),
            ("Chemin Word MAR", "word_mar", "file"),
            ("Chemin Word IADE", "word_iade", "file"),
            ("Dossier PDF Contrats MAR", "pdf_mar", "directory"),
            ("Dossier PDF Contrats IADE", "pdf_iade", "directory"),
            ("Dossier Bulletins de salaire", "bulletins_salaire", "directory"),
            ("Dossier Frais/Factures", "dossier_factures", "directory"),
            ("Chemin Excel Salariés", "excel_salaries", "file")
        ]
        
        # Variables pour les chemins
        self.path_vars = {}
        
        # Créer les champs pour chaque chemin
        for i, (label, key, file_type) in enumerate(paths):
            var = tk.StringVar(value=settings.get(key, ""))
            self.path_vars[key] = var
            
            # Frame pour chaque ligne
            row_frame = tk.Frame(scrollable_frame, bg=COLORS["secondary"])
            row_frame.pack(fill="x", pady=5)
            
            # Label
            tk.Label(
                row_frame, 
                text=label, 
                width=20, 
                anchor="w", 
                bg=COLORS["secondary"]
            ).pack(side="left")
            
            # Entrée
            entry = tk.Entry(row_frame, textvariable=var, width=40)
            entry.pack(side="left", padx=5, fill="x", expand=True)
            
            # Bouton parcourir
            browse_button = tk.Button(
                row_frame, 
                text="...", 
                command=lambda v=var, t=file_type: self.select_path(v, t),
                width=3
            )
            browse_button.pack(side="left")
        
        # Bouton de sauvegarde
        create_button(
            parent, 
            text="Enregistrer", 
            command=lambda: self.save_paths(),
            style="accent",
            width=15
        ).pack(pady=10)
    
    def select_path(self, variable, file_type):
        """
        Sélectionne un chemin de fichier ou de dossier.
        """
        if file_type == "directory":
            path = filedialog.askdirectory()
        else:
            path = filedialog.askopenfilename()
        
        if path:
            variable.set(path)
    
    def doctors_tab_setup(self, parent):
        """
        Crée le contenu du sous-onglet Médecins.
        """
        # Sous-onglets pour les différents types de médecins
        doctors_tabs = TabView(parent)
        doctors_tabs.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Sous-onglet MAR titulaires
        mar_tab = doctors_tabs.add_tab("MAR titulaires", bg=COLORS["secondary"])
        
        # Bouton pour gérer les MAR titulaires
        create_button(
            mar_tab, 
            text="Gérer les MAR titulaires", 
            command=lambda: self.manage_mar_titulaires(),
            style="primary",
            width=25
        ).pack(pady=20)
        
        # Sous-onglet MAR remplaçants
        mar_remp_tab = doctors_tabs.add_tab("MAR remplaçants", bg=COLORS["secondary"])
        
        # Bouton pour gérer les MAR remplaçants
        create_button(
            mar_remp_tab, 
            text="Gérer les MAR remplaçants", 
            command=lambda: self.manage_mar_remplacants(),
            style="primary",
            width=25
        ).pack(pady=20)
    
    def iade_tab_setup(self, parent):
        """
        Crée le contenu du sous-onglet IADE.
        """
        # Bouton pour gérer les IADE remplaçants
        create_button(
            parent, 
            text="Gérer les IADE remplaçants", 
            command=lambda: self.manage_iade_remplacants(),
            style="primary",
            width=25
        ).pack(pady=20)
    
    def employees_tab_setup(self, parent):
        """
        Crée le contenu du sous-onglet Salariés.
        """
        # Bouton pour gérer les salariés
        create_button(
            parent, 
            text="Gérer les salariés", 
            command=lambda: self.manage_salaries(),
            style="primary",
            width=25
        ).pack(pady=20)
    
    def docusign_tab_setup(self, parent):
        """
        Crée le contenu du sous-onglet DocuSign.
        """
        from widgets.widgets import FormField
        import json
        
        # Titre
        create_label(
            parent, 
            "Configuration DocuSign",
            style="subtitle"
        ).pack(pady=10)
        
        # Charger la configuration
        config_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")
        if os.path.exists(config_file):
            with open(config_file, "r") as f:
                config = json.load(f)
        else:
            config = {}
        
        # Variables pour les champs
        self.docusign_vars = {
            "login_page": tk.StringVar(value=config.get("docusign_login_page", "https://account.docusign.com")),
            "email": tk.StringVar(value=config.get("docusign_email", "")),
            "password": tk.StringVar(value=config.get("docusign_password", ""))
        }
        
        # Champs de formulaire
        FormField(
            parent, 
            "Page de login DocuSign :", 
            variable=self.docusign_vars["login_page"]
        ).pack(fill="x", padx=50, pady=5)
        
        FormField(
            parent, 
            "Email DocuSign :", 
            variable=self.docusign_vars["email"]
        ).pack(fill="x", padx=50, pady=5)
        
        # Champ de mot de passe modifié pour cacher les caractères
        password_frame = tk.Frame(parent, bg=parent["bg"])
        password_frame.pack(fill="x", padx=50, pady=5)
        
        tk.Label(
            password_frame, 
            text="Mot de passe (laisser vide si non stocké) :",
            font=("Arial", 10),
            bg=parent["bg"],
            anchor="w"
        ).pack(fill="x", pady=(5, 2))
        
        password_entry = tk.Entry(
            password_frame, 
            textvariable=self.docusign_vars["password"],
            font=("Arial", 10),
            show="*"  # Cette ligne est la clé - elle masque le mot de passe avec des *
        )
        password_entry.pack(fill="x", pady=(0, 5))
        
        # Bouton de sauvegarde
        create_button(
            parent, 
            text="Enregistrer", 
            command=lambda: self.save_docusign(),
            style="accent",
            width=15
        ).pack(pady=20)
    
    def save_paths(self):
        """
        Sauvegarde les chemins de fichiers.
        """
        settings = load_settings()
        
        # Mettre à jour les chemins
        for key, var in self.path_vars.items():
            settings[key] = var.get()
        
        # Sauvegarder les paramètres
        if save_settings(settings):
            show_success("Paramètres", "Les chemins ont été enregistrés avec succès.")
        else:
            show_error("Paramètres", "Erreur lors de l'enregistrement des chemins.")
    
    def save_docusign(self):
        """
        Sauvegarde les paramètres DocuSign.
        """
        import json
        
        config_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")
        
        # Charger la configuration existante
        if os.path.exists(config_file):
            with open(config_file, "r") as f:
                config = json.load(f)
        else:
            config = {}
        
        # Mettre à jour les paramètres DocuSign
        config["docusign_login_page"] = self.docusign_vars["login_page"].get()
        config["docusign_email"] = self.docusign_vars["email"].get()
        config["docusign_password"] = self.docusign_vars["password"].get()
        
        # Sauvegarder la configuration
        try:
            with open(config_file, "w") as f:
                json.dump(config, f, indent=4)
            
            show_success("Paramètres", "Les paramètres DocuSign ont été enregistrés avec succès.")
        except Exception as e:
            show_error("Paramètres", f"Erreur lors de l'enregistrement des paramètres DocuSign : {str(e)}")
    
    def open_contract_creation_mar(self):
        """
        Ouvre l'interface de création de contrat MAR.
        """
        self.status_bar.set_message("Ouverture de l'interface de création de contrat MAR...")
        
        # Nettoyer l'interface existante
        for widget in self.contracts_tab.winfo_children():
            if isinstance(widget, TabView):
                continue
            widget.destroy()
        
        # Créer un cadre pour l'interface de création de contrat
        contract_frame = create_frame(self.contracts_tab)
        contract_frame.pack(fill="both", expand=True, padx=10, pady=10)
        # Ajouter un bouton de retour en haut à gauche de l'interface
        return_button_frame = create_frame(contract_frame, bg=COLORS["secondary"])
        return_button_frame.pack(fill="x", pady=(0, 10))
        
        # Bouton de retour
        create_button(
            return_button_frame, 
            text="🔙 Retour aux contrats", 
            command=lambda: self.restore_contracts_tab(),
            style="neutral",
            width=15
        ).pack(side="left", padx=5, pady=5)
        
        try:
            # Appeler la fonction de création de contrat MAR avec le frame créé
            contrat53.create_contract_interface_mar(contract_frame)
            self.status_bar.set_message("Interface de création de contrat MAR ouverte.")
        except Exception as e:
            show_error("Erreur", f"Impossible d'ouvrir l'interface de création de contrat MAR: {str(e)}")
            self.status_bar.set_message("Erreur lors de l'ouverture de l'interface de création de contrat MAR.")
       
    def restore_contracts_tab(self):
        """
        Restaure l'interface principale des contrats après être revenu d'une sous-interface.
        """
        # Nettoyer l'interface existante
        for widget in self.contracts_tab.winfo_children():
            widget.destroy()
        
        # Recréer l'interface des contrats
        self.create_contracts_tab()
        
        # Mettre à jour la barre de statut
        self.status_bar.set_message("Interface des contrats restaurée.")       
            
    def open_contract_creation_iade(self):
        """
        Ouvre l'interface de création de contrat IADE.
        """
        self.status_bar.set_message("Ouverture de l'interface de création de contrat IADE...")
        
        # Nettoyer l'interface existante
        for widget in self.contracts_tab.winfo_children():
            widget.destroy()
        
        # Créer un cadre pour l'interface de création de contrat
        contract_frame = create_frame(self.contracts_tab)
        contract_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Ajouter un bouton de retour en haut à gauche de l'interface
        return_button_frame = create_frame(contract_frame, bg=COLORS["secondary"])
        return_button_frame.pack(fill="x", pady=(0, 10))
        
        # Bouton de retour
        create_button(
            return_button_frame, 
            text="🔙 Retour aux contrats", 
            command=lambda: self.restore_contracts_tab(),
            style="neutral",
            width=15
        ).pack(side="left", padx=5, pady=5)
        
        try:
            # Appeler la fonction de création de contrat IADE avec le frame créé
            contrat53.create_contract_interface_iade(contract_frame)
            self.status_bar.set_message("Interface de création de contrat IADE ouverte.")
        except Exception as e:
            show_error("Erreur", f"Impossible d'ouvrir l'interface de création de contrat IADE: {str(e)}")
            self.status_bar.set_message("Erreur lors de l'ouverture de l'interface de création de contrat IADE.")
            
            
    def manage_mar_contracts(self):
        """
        Ouvre l'interface de gestion des contrats MAR existants.
        """
        self.status_bar.set_message("Ouverture de l'interface de gestion des contrats MAR...")
        
        # À implémenter
        show_info("Fonctionnalité à venir", "La gestion des contrats MAR existants sera disponible prochainement.")
        
        self.status_bar.set_message("Prêt")
    
    def manage_iade_contracts(self):
        """
        Ouvre l'interface de gestion des contrats IADE existants.
        """
        self.status_bar.set_message("Ouverture de l'interface de gestion des contrats IADE...")
        
        # À implémenter
        show_info("Fonctionnalité à venir", "La gestion des contrats IADE existants sera disponible prochainement.")
        
        self.status_bar.set_message("Prêt")
    
    def open_invoice_analysis(self):
        """Ouvre l'interface d'analyse des factures."""
        self.status_bar.set_message("Ouverture de l'interface d'analyse des factures...")
        
        # Nettoyer l'interface existante
        for widget in self.invoices_frame.winfo_children():
            widget.destroy()
        
        # Créer un cadre pour l'interface d'analyse des factures
        analysis_frame = create_frame(self.invoices_frame)
        analysis_frame.pack(fill="both", expand=True)
        
        try:
            # Créer l'instance de l'analyseur
            factures_path = get_file_path("dossier_factures", verify_exists=True, create_if_missing=True)
            analyseur = analyse_facture.AnalyseFactures(factures_path)
            
            # Créer l'interface de l'analyseur dans le cadre
            analyseur.creer_interface(analysis_frame)
            
            self.status_bar.set_message("Interface d'analyse des factures ouverte.")
        except Exception as e:
            import traceback
            error_details = traceback.format_exc()
            print(f"❌ Erreur lors de l'ouverture de l'interface d'analyse des factures: {e}")
            print(f"Détails: {error_details}")
            
            # Afficher un message d'erreur convivial
            error_frame = create_frame(analysis_frame, bg="#fff0f0")
            error_frame.pack(fill="both", expand=True, padx=20, pady=20)
            
            tk.Label(error_frame, 
                    text="Une erreur s'est produite", 
                    font=("Arial", 16, "bold"),
                    bg="#fff0f0", fg="#c00000").pack(pady=(30, 10))
            
            tk.Label(error_frame, 
                    text=f"Détails de l'erreur: {str(e)}", 
                    font=("Arial", 12),
                    bg="#fff0f0", fg="#c00000").pack(pady=(0, 20))
            
            # Bouton pour afficher les détails techniques
            def show_details():
                details_window = tk.Toplevel()
                details_window.title("Détails techniques")
                details_window.geometry("800x600")
                
                details_text = tk.Text(details_window, wrap="word", font=("Courier", 10))
                details_text.pack(fill="both", expand=True, padx=10, pady=10)
                details_text.insert("1.0", error_details)
                
                tk.Button(details_window, text="Fermer", command=details_window.destroy).pack(pady=10)
            
            tk.Button(error_frame, text="Afficher les détails techniques", command=show_details).pack(pady=10)
            
            # Bouton pour réessayer
            tk.Button(error_frame, text="Réessayer", command=self.open_invoice_analysis).pack(pady=10)
            
            self.status_bar.set_message("Erreur lors de l'ouverture de l'interface d'analyse des factures.")
            
    
    def open_transfer(self):
        """
        Ouvre l'interface de virement.
        """
        self.status_bar.set_message("Ouverture de l'interface de virement...")
        
        # À implémenter
        show_info("Fonctionnalité à venir", "L'interface de virement sera disponible prochainement.")
        
        self.status_bar.set_message("Prêt")
    
    def open_mars_transfer(self):
        """
        Ouvre l'interface de virement des rémunérations MAR.
        """
        self.status_bar.set_message("Ouverture de l'interface de virement MAR...")
        
        # À implémenter
        show_info("Fonctionnalité à venir", "L'interface de virement MAR sera disponible prochainement.")
        
        self.status_bar.set_message("Prêt")
    
    def open_bulletins(self):
        """
        Ouvre l'interface de consultation des bulletins.
        """
        self.status_bar.set_message("Ouverture de l'interface de consultation des bulletins...")
        
        # Nettoyer l'interface existante
        for widget in self.bulletins_frame.winfo_children():
            widget.destroy()
        
        # Créer un cadre pour l'interface de consultation des bulletins
        bulletins_frame = create_frame(self.bulletins_frame)
        bulletins_frame.pack(fill="both", expand=True)
        
        # Appeler la fonction de consultation des bulletins
        bulletins.show_bulletins_in_frame(bulletins_frame)
        
        self.status_bar.set_message("Interface de consultation des bulletins ouverte.")
    
    def scan_new_bulletin(self):
        """
        Ouvre l'interface de scan d'un nouveau bulletin.
        """
        self.status_bar.set_message("Ouverture de l'interface de scan d'un nouveau bulletin...")
        
        # Appeler la fonction de scan d'un nouveau bulletin
        bulletins.scan_new_pdf()
        
        self.status_bar.set_message("Prêt")
    
    def manage_mar_titulaires(self):
        """
        Ouvre l'interface de gestion des MAR titulaires.
        """
        self.status_bar.set_message("Ouverture de l'interface de gestion des MAR titulaires...")
        
        # Nettoyage de l'interface existante si nécessaire
        for widget in self.settings_tab.winfo_children():
            if isinstance(widget, TabView):
                continue
            widget.destroy()
        
        # Créer un cadre pour l'interface de gestion des MAR titulaires
        management_frame = create_frame(self.settings_tab)
        management_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Bouton de retour
        return_frame = create_frame(management_frame, bg=COLORS["secondary"])
        return_frame.pack(fill="x", pady=(0, 10))
        
        create_button(
            return_frame,
            text="🔙 Retour aux paramètres",
            command=lambda: self.create_settings_tab(),
            style="neutral",
            width=20
        ).pack(side="left", padx=5, pady=5)
        
        # Titre
        create_label(
            management_frame,
            "Gestion des MAR titulaires",
            style="title"
        ).pack(pady=10)
        
        try:
            # Chargement des données depuis la bonne feuille
            excel_path = get_file_path("excel_mar", verify_exists=True)
            if not excel_path:
                raise ValueError("Le chemin du fichier Excel MAR n'est pas défini.")
            
            mars_titulaires = pd.read_excel(excel_path, sheet_name="MARS SELARL")
            
            # Cadre principal
            main_frame = create_frame(management_frame, bg=COLORS["secondary"])
            main_frame.pack(fill="both", expand=True, padx=10, pady=10)
            
            # Créer une listbox pour afficher les MARS titulaires
            listbox = tk.Listbox(main_frame, width=50, height=15, font=("Arial", 12))
            listbox.pack(side="left", fill="both", expand=True, padx=5, pady=10)
            
            # Ajouter une scrollbar
            scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=listbox.yview)
            scrollbar.pack(side="right", fill="y")
            listbox.config(yscrollcommand=scrollbar.set)
            
            # Variable pour stocker les données
            mars_data = []
            
            def refresh_listbox():
                """Met à jour la liste des MARS affichée dans la Listbox."""
                listbox.delete(0, tk.END)
                mars_data.clear()
                
                for _, row in mars_titulaires.iterrows():
                    nom = row["NOM"] if not pd.isna(row["NOM"]) else ""
                    prenom = row["PRENOM"] if not pd.isna(row["PRENOM"]) else ""
                    full_name = f"{nom} {prenom}".strip()
                    mars_data.append(row)
                    listbox.insert(tk.END, full_name)
            
            # Remplir la listbox initialement
            refresh_listbox()
            
            # Fonctions pour les boutons
            def on_modify():
                """Modifier un MAR titulaire."""
                selected_index = listbox.curselection()
                if not selected_index:
                    show_warning("Avertissement", "Veuillez sélectionner un médecin.")
                    return
                
                selected_index = selected_index[0]
                selected_row = mars_data[selected_index]
                
                # Créer une fenêtre de modification
                modify_window = tk.Toplevel(management_frame.winfo_toplevel())
                modify_window.title("Modifier un MAR titulaire")
                modify_window.geometry("400x400")
                modify_window.grab_set()  # Rendre la fenêtre modale
                
                # Variables pour les champs
                nom_var = StringVar(value=selected_row["NOM"] if not pd.isna(selected_row["NOM"]) else "")
                prenom_var = StringVar(value=selected_row["PRENOM"] if not pd.isna(selected_row["PRENOM"]) else "")
                ordre_var = StringVar(value=selected_row["N ORDRE"] if not pd.isna(selected_row["N ORDRE"]) else "")
                email_var = StringVar(value=selected_row["EMAIL"] if not pd.isna(selected_row["EMAIL"]) else "")
                iban_var = StringVar(value=selected_row.get("IBAN", "") if not pd.isna(selected_row.get("IBAN", "")) else "")
                
                # Création des champs
                padx, pady = 10, 5
                row = 0
                
                ttk.Label(modify_window, text="Nom:").grid(row=row, column=0, sticky="w", padx=padx, pady=pady)
                ttk.Entry(modify_window, textvariable=nom_var, width=30).grid(row=row, column=1, padx=padx, pady=pady)
                row += 1
                
                ttk.Label(modify_window, text="Prénom:").grid(row=row, column=0, sticky="w", padx=padx, pady=pady)
                ttk.Entry(modify_window, textvariable=prenom_var, width=30).grid(row=row, column=1, padx=padx, pady=pady)
                row += 1
                
                ttk.Label(modify_window, text="N° Ordre:").grid(row=row, column=0, sticky="w", padx=padx, pady=pady)
                ttk.Entry(modify_window, textvariable=ordre_var, width=30).grid(row=row, column=1, padx=padx, pady=pady)
                row += 1
                
                ttk.Label(modify_window, text="Email:").grid(row=row, column=0, sticky="w", padx=padx, pady=pady)
                ttk.Entry(modify_window, textvariable=email_var, width=30).grid(row=row, column=1, padx=padx, pady=pady)
                row += 1
                
                ttk.Label(modify_window, text="IBAN:").grid(row=row, column=0, sticky="w", padx=padx, pady=pady)
                ttk.Entry(modify_window, textvariable=iban_var, width=30).grid(row=row, column=1, padx=padx, pady=pady)
                row += 1
                
                def save_changes():
                    """Enregistre les modifications et met à jour Excel."""
                    # Mise à jour des données
                    mars_titulaires.loc[selected_index, "NOM"] = nom_var.get()
                    mars_titulaires.loc[selected_index, "PRENOM"] = prenom_var.get()
                    mars_titulaires.loc[selected_index, "N ORDRE"] = ordre_var.get()
                    mars_titulaires.loc[selected_index, "EMAIL"] = email_var.get()
                    
                    # Ajouter IBAN s'il n'existe pas
                    if "IBAN" not in mars_titulaires.columns:
                        mars_titulaires["IBAN"] = ""
                    mars_titulaires.loc[selected_index, "IBAN"] = iban_var.get()
                    
                    try:
                        # Sauvegarde dans Excel
                        from modules.contracts.contrat53_new import save_excel_with_updated_sheet
                        save_excel_with_updated_sheet(excel_path, "MARS SELARL", mars_titulaires)
                        show_success("Succès", "Modifications enregistrées avec succès.")
                        
                        # Mise à jour de l'affichage
                        refresh_listbox()
                        modify_window.destroy()
                    except Exception as e:
                        show_error("Erreur", f"Erreur lors de la sauvegarde : {e}")
                
                # Boutons de validation
                buttons_frame = tk.Frame(modify_window)
                buttons_frame.grid(row=row, column=0, columnspan=2, pady=15)
                
                ttk.Button(buttons_frame, text="Enregistrer", command=save_changes).pack(side="left", padx=10)
                ttk.Button(buttons_frame, text="Annuler", command=modify_window.destroy).pack(side="left", padx=10)
            
            def on_add():
                """Ajouter un nouveau MAR titulaire."""
                # Créer une fenêtre d'ajout
                add_window = tk.Toplevel(management_frame.winfo_toplevel())
                add_window.title("Ajouter un MAR titulaire")
                add_window.geometry("400x400")
                add_window.grab_set()  # Rendre la fenêtre modale
                
                # Variables pour les champs
                nom_var = StringVar()
                prenom_var = StringVar()
                ordre_var = StringVar()
                email_var = StringVar()
                iban_var = StringVar()
                
                # Création des champs
                padx, pady = 10, 5
                row = 0
                
                ttk.Label(add_window, text="Nom:").grid(row=row, column=0, sticky="w", padx=padx, pady=pady)
                ttk.Entry(add_window, textvariable=nom_var, width=30).grid(row=row, column=1, padx=padx, pady=pady)
                row += 1
                
                ttk.Label(add_window, text="Prénom:").grid(row=row, column=0, sticky="w", padx=padx, pady=pady)
                ttk.Entry(add_window, textvariable=prenom_var, width=30).grid(row=row, column=1, padx=padx, pady=pady)
                row += 1
                
                ttk.Label(add_window, text="N° Ordre:").grid(row=row, column=0, sticky="w", padx=padx, pady=pady)
                ttk.Entry(add_window, textvariable=ordre_var, width=30).grid(row=row, column=1, padx=padx, pady=pady)
                row += 1
                
                ttk.Label(add_window, text="Email:").grid(row=row, column=0, sticky="w", padx=padx, pady=pady)
                ttk.Entry(add_window, textvariable=email_var, width=30).grid(row=row, column=1, padx=padx, pady=pady)
                row += 1
                
                ttk.Label(add_window, text="IBAN:").grid(row=row, column=0, sticky="w", padx=padx, pady=pady)
                ttk.Entry(add_window, textvariable=iban_var, width=30).grid(row=row, column=1, padx=padx, pady=pady)
                row += 1
                
                def save_new():
                    """Enregistre le nouveau MAR et met à jour Excel."""
                    # Vérification des champs obligatoires
                    if not nom_var.get().strip() or not prenom_var.get().strip():
                        show_warning("Attention", "Veuillez renseigner au moins le nom et le prénom.")
                        return
                    
                    # Préparation des nouvelles données
                    new_row = {
                        "NOM": nom_var.get(),
                        "PRENOM": prenom_var.get(),
                        "N ORDRE": ordre_var.get(),
                        "EMAIL": email_var.get()
                    }
                    
                    # Ajouter IBAN s'il n'existe pas
                    if "IBAN" not in mars_titulaires.columns:
                        mars_titulaires["IBAN"] = ""
                    new_row["IBAN"] = iban_var.get()
                    
                    try:
                        # Ajout au DataFrame
                        mars_titulaires.loc[len(mars_titulaires)] = new_row
                        
                        # Sauvegarde dans Excel
                        from modules.contracts.contrat53_new import save_excel_with_updated_sheet
                        save_excel_with_updated_sheet(excel_path, "MARS SELARL", mars_titulaires)
                        show_success("Succès", "MAR titulaire ajouté avec succès.")
                        
                        # Mise à jour de l'affichage
                        refresh_listbox()
                        add_window.destroy()
                    except Exception as e:
                        show_error("Erreur", f"Erreur lors de l'ajout : {e}")
                
                # Boutons de validation
                buttons_frame = tk.Frame(add_window)
                buttons_frame.grid(row=row, column=0, columnspan=2, pady=15)
                
                ttk.Button(buttons_frame, text="Ajouter", command=save_new).pack(side="left", padx=10)
                ttk.Button(buttons_frame, text="Annuler", command=add_window.destroy).pack(side="left", padx=10)
            
            def on_delete():
                """Supprimer un MAR titulaire."""
                selected_index = listbox.curselection()
                if not selected_index:
                    show_warning("Avertissement", "Veuillez sélectionner un médecin.")
                    return
                
                selected_index = selected_index[0]
                selected_row = mars_data[selected_index]
                
                # Confirmation de suppression
                nom = selected_row["NOM"] if not pd.isna(selected_row["NOM"]) else ""
                prenom = selected_row["PRENOM"] if not pd.isna(selected_row["PRENOM"]) else ""
                full_name = f"{nom} {prenom}".strip()
                
                if ask_question("Confirmation", f"Voulez-vous vraiment supprimer {full_name} ?"):
                    try:
                        # Suppression de la ligne
                        mars_titulaires.drop(mars_titulaires.index[selected_index], inplace=True)
                        mars_titulaires.reset_index(drop=True, inplace=True)
                        
                        # Sauvegarde dans Excel
                        from modules.contracts.contrat53_new import save_excel_with_updated_sheet
                        save_excel_with_updated_sheet(excel_path, "MARS SELARL", mars_titulaires)
                        show_success("Succès", f"{full_name} supprimé avec succès.")
                        
                        # Mise à jour de l'affichage
                        refresh_listbox()
                    except Exception as e:
                        show_error("Erreur", f"Erreur lors de la suppression : {e}")
            
            # Boutons d'action
            buttons_frame = create_frame(management_frame)
            buttons_frame.pack(pady=10)
            
            create_button(
                buttons_frame, 
                text="➕ Ajouter", 
                command=on_add,
                style="primary",
                width=15
            ).pack(side="left", padx=10)
            
            create_button(
                buttons_frame, 
                text="✏️ Modifier", 
                command=on_modify,
                style="info",
                width=15
            ).pack(side="left", padx=10)
            
            create_button(
                buttons_frame, 
                text="🗑️ Supprimer", 
                command=on_delete,
                style="danger",
                width=15
            ).pack(side="left", padx=10)
            
            self.status_bar.set_message("Interface de gestion des MAR titulaires ouverte.")
        except Exception as e:
            import traceback
            traceback.print_exc()
            show_error("Erreur", f"Impossible d'ouvrir l'interface de gestion des MAR titulaires: {str(e)}")
            self.status_bar.set_message("Erreur lors de l'ouverture de l'interface de gestion des MAR titulaires.")

    
    def manage_iade_remplacants(self):
        """
        Ouvre l'interface de gestion des IADE remplaçants.
        """
        self.status_bar.set_message("Ouverture de l'interface de gestion des IADE remplaçants...")
        
        # Nettoyage de l'interface existante si nécessaire
        for widget in self.settings_tab.winfo_children():
            if isinstance(widget, TabView):
                continue
            widget.destroy()
        
        # Créer un cadre pour l'interface de gestion des IADE remplaçants
        management_frame = create_frame(self.settings_tab)
        management_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Bouton de retour
        return_frame = create_frame(management_frame, bg=COLORS["secondary"])
        return_frame.pack(fill="x", pady=(0, 10))
        
        create_button(
            return_frame,
            text="🔙 Retour aux paramètres",
            command=lambda: self.create_settings_tab(),
            style="neutral",
            width=20
        ).pack(side="left", padx=5, pady=5)
        
        # Titre
        create_label(
            management_frame,
            "Gestion des IADE remplaçants",
            style="title"
        ).pack(pady=10)
        
        try:
            # Chargement des données depuis la bonne feuille
            excel_path = get_file_path("excel_iade", verify_exists=True)
            if not excel_path:
                raise ValueError("Le chemin du fichier Excel IADE n'est pas défini.")
            
            iade_data = pd.read_excel(excel_path, sheet_name="Coordonnées IADEs")
            
            # Cadre principal
            main_frame = create_frame(management_frame, bg=COLORS["secondary"])
            main_frame.pack(fill="both", expand=True, padx=10, pady=10)
            
            # Créer une listbox pour afficher les IADE remplaçants
            listbox = tk.Listbox(main_frame, width=50, height=15, font=("Arial", 12))
            listbox.pack(side="left", fill="both", expand=True, padx=5, pady=10)
            
            # Ajouter une scrollbar
            scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=listbox.yview)
            scrollbar.pack(side="right", fill="y")
            listbox.config(yscrollcommand=scrollbar.set)
            
            # Variable pour stocker les données
            iade_data_list = []
            
            def refresh_listbox():
                """Met à jour la liste des IADE affichée dans la Listbox."""
                listbox.delete(0, tk.END)
                iade_data_list.clear()
                
                for _, row in iade_data.iterrows():
                    nom = row["NOMR"] if not pd.isna(row["NOMR"]) else ""
                    prenom = row["PRENOMR"] if not pd.isna(row["PRENOMR"]) else ""
                    full_name = f"{nom} {prenom}".strip()
                    iade_data_list.append(row)
                    listbox.insert(tk.END, full_name)
            
            # Remplir la listbox initialement
            refresh_listbox()
            
            # Fonctions pour les boutons
            def on_modify():
                """Modifier un IADE remplaçant."""
                selected_index = listbox.curselection()
                if not selected_index:
                    show_warning("Avertissement", "Veuillez sélectionner un IADE.")
                    return
                
                selected_index = selected_index[0]
                selected_row = iade_data_list[selected_index]
                
                # Créer une fenêtre de modification
                modify_window = tk.Toplevel(management_frame.winfo_toplevel())
                modify_window.title("Modifier un IADE remplaçant")
                modify_window.geometry("500x600")
                modify_window.grab_set()  # Rendre la fenêtre modale
                
                # Variables pour les champs
                nom_var = StringVar(value=selected_row["NOMR"] if not pd.isna(selected_row["NOMR"]) else "")
                prenom_var = StringVar(value=selected_row["PRENOMR"] if not pd.isna(selected_row["PRENOMR"]) else "")
                email_var = StringVar(value=selected_row["EMAIL"] if not pd.isna(selected_row["EMAIL"]) else "")
                ddn_var = StringVar(value=selected_row["DDNR"] if not pd.isna(selected_row["DDNR"]) else "")
                lieu_naissance_var = StringVar(value=selected_row["LIEUNR"] if not pd.isna(selected_row["LIEUNR"]) else "")
                dept_naissance_var = StringVar(value=selected_row["DPTN"] if not pd.isna(selected_row["DPTN"]) else "")
                adresse_var = StringVar(value=selected_row["ADRESSER"] if not pd.isna(selected_row["ADRESSER"]) else "")
                secu_var = StringVar(value=selected_row["NOSSR"] if not pd.isna(selected_row["NOSSR"]) else "")
                nationalite_var = StringVar(value=selected_row["NATR"] if not pd.isna(selected_row["NATR"]) else "")
                iban_var = StringVar(value=selected_row.get("IBAN", "") if not pd.isna(selected_row.get("IBAN", "")) else "")
                sexe_var = StringVar(value=selected_row.get("SEXE", "Monsieur") if not pd.isna(selected_row.get("SEXE", "")) else "Monsieur")
                
                # Fonction pour mettre à jour les champs dérivés du sexe
                def update_gender_fields(*args):
                    if sexe_var.get() == "Madame":
                        er_var.set("e")
                        ilr_var.set("elle")
                        salarier_var.set("à la salariée")
                    else:
                        er_var.set("")
                        ilr_var.set("il")
                        salarier_var.set("au salarié")
                
                # Variables supplémentaires dérivées du sexe
                er_var = StringVar(value=selected_row.get("ER", ""))
                ilr_var = StringVar(value=selected_row.get("ILR", ""))
                salarier_var = StringVar(value=selected_row.get("SALARIER", ""))
                
                # Lier la mise à jour automatique au changement de sexe
                sexe_var.trace_add("write", update_gender_fields)
                
                # Mise à jour initiale des champs dérivés
                update_gender_fields()
                
                # Création des champs
                padx, pady = 10, 5
                row = 0
                
                ttk.Label(modify_window, text="Nom:").grid(row=row, column=0, sticky="w", padx=padx, pady=pady)
                ttk.Entry(modify_window, textvariable=nom_var, width=30).grid(row=row, column=1, padx=padx, pady=pady)
                row += 1
                
                ttk.Label(modify_window, text="Prénom:").grid(row=row, column=0, sticky="w", padx=padx, pady=pady)
                ttk.Entry(modify_window, textvariable=prenom_var, width=30).grid(row=row, column=1, padx=padx, pady=pady)
                row += 1
                
                ttk.Label(modify_window, text="Date de naissance:").grid(row=row, column=0, sticky="w", padx=padx, pady=pady)
                ttk.Entry(modify_window, textvariable=ddn_var, width=30).grid(row=row, column=1, padx=padx, pady=pady)
                row += 1
                
                ttk.Label(modify_window, text="Lieu de naissance:").grid(row=row, column=0, sticky="w", padx=padx, pady=pady)
                ttk.Entry(modify_window, textvariable=lieu_naissance_var, width=30).grid(row=row, column=1, padx=padx, pady=pady)
                row += 1
                
                ttk.Label(modify_window, text="Département de naissance:").grid(row=row, column=0, sticky="w", padx=padx, pady=pady)
                ttk.Entry(modify_window, textvariable=dept_naissance_var, width=30).grid(row=row, column=1, padx=padx, pady=pady)
                row += 1
                
                ttk.Label(modify_window, text="Nationalité:").grid(row=row, column=0, sticky="w", padx=padx, pady=pady)
                ttk.Entry(modify_window, textvariable=nationalite_var, width=30).grid(row=row, column=1, padx=padx, pady=pady)
                row += 1
                
                ttk.Label(modify_window, text="Adresse:").grid(row=row, column=0, sticky="w", padx=padx, pady=pady)
                ttk.Entry(modify_window, textvariable=adresse_var, width=30).grid(row=row, column=1, padx=padx, pady=pady)
                row += 1
                
                ttk.Label(modify_window, text="Email:").grid(row=row, column=0, sticky="w", padx=padx, pady=pady)
                ttk.Entry(modify_window, textvariable=email_var, width=30).grid(row=row, column=1, padx=padx, pady=pady)
                row += 1
                
                ttk.Label(modify_window, text="N° Sécurité Sociale:").grid(row=row, column=0, sticky="w", padx=padx, pady=pady)
                ttk.Entry(modify_window, textvariable=secu_var, width=30).grid(row=row, column=1, padx=padx, pady=pady)
                row += 1
                
                ttk.Label(modify_window, text="IBAN:").grid(row=row, column=0, sticky="w", padx=padx, pady=pady)
                ttk.Entry(modify_window, textvariable=iban_var, width=30).grid(row=row, column=1, padx=padx, pady=pady)
                row += 1
                
                ttk.Label(modify_window, text="Sexe:").grid(row=row, column=0, sticky="w", padx=padx, pady=pady)
                ttk.OptionMenu(modify_window, sexe_var, "Monsieur", "Madame").grid(row=row, column=1, sticky="w", padx=padx, pady=pady)
                row += 1
                
                # Champs dérivés du sexe (non modifiables)
                ttk.Label(modify_window, text="ER:").grid(row=row, column=0, sticky="w", padx=padx, pady=pady)
                ttk.Entry(modify_window, textvariable=er_var, state="readonly", width=30).grid(row=row, column=1, padx=padx, pady=pady)
                row += 1
                
                ttk.Label(modify_window, text="ILR:").grid(row=row, column=0, sticky="w", padx=padx, pady=pady)
                ttk.Entry(modify_window, textvariable=ilr_var, state="readonly", width=30).grid(row=row, column=1, padx=padx, pady=pady)
                row += 1
                
                ttk.Label(modify_window, text="SALARIER:").grid(row=row, column=0, sticky="w", padx=padx, pady=pady)
                ttk.Entry(modify_window, textvariable=salarier_var, state="readonly", width=30).grid(row=row, column=1, padx=padx, pady=pady)
                row += 1
                
                def save_changes():
                    """Enregistre les modifications et met à jour Excel."""
                    # Mise à jour des données
                    iade_data.loc[selected_index, "NOMR"] = nom_var.get()
                    iade_data.loc[selected_index, "PRENOMR"] = prenom_var.get()
                    iade_data.loc[selected_index, "EMAIL"] = email_var.get()
                    iade_data.loc[selected_index, "DDNR"] = ddn_var.get()
                    iade_data.loc[selected_index, "LIEUNR"] = lieu_naissance_var.get()
                    iade_data.loc[selected_index, "DPTN"] = dept_naissance_var.get()
                    iade_data.loc[selected_index, "ADRESSER"] = adresse_var.get()
                    iade_data.loc[selected_index, "NOSSR"] = secu_var.get()
                    iade_data.loc[selected_index, "NATR"] = nationalite_var.get()
                    iade_data.loc[selected_index, "SEXE"] = sexe_var.get()
                    iade_data.loc[selected_index, "ER"] = er_var.get()
                    iade_data.loc[selected_index, "ILR"] = ilr_var.get()
                    iade_data.loc[selected_index, "SALARIER"] = salarier_var.get()
                    
                    # Ajouter IBAN s'il n'existe pas dans le DataFrame
                    if "IBAN" not in iade_data.columns:
                        iade_data["IBAN"] = ""
                    iade_data.loc[selected_index, "IBAN"] = iban_var.get()
                    
                    try:
                        # Sauvegarde dans Excel
                        from modules.contracts.contrat53_new import save_excel_with_updated_sheet
                        save_excel_with_updated_sheet(excel_path, "Coordonnées IADEs", iade_data)
                        show_success("Succès", "Modifications enregistrées avec succès.")
                        
                        # Mise à jour de l'affichage
                        refresh_listbox()
                        modify_window.destroy()
                    except Exception as e:
                        show_error("Erreur", f"Erreur lors de la sauvegarde : {e}")
                
                # Boutons de validation
                buttons_frame = tk.Frame(modify_window)
                buttons_frame.grid(row=row, column=0, columnspan=2, pady=15)
                
                ttk.Button(buttons_frame, text="Enregistrer", command=save_changes).pack(side="left", padx=10)
                ttk.Button(buttons_frame, text="Annuler", command=modify_window.destroy).pack(side="left", padx=10)
            
            def on_add():
                """Ajouter un nouveau IADE remplaçant."""
                # Créer une fenêtre d'ajout
                add_window = tk.Toplevel(management_frame.winfo_toplevel())
                add_window.title("Ajouter un IADE remplaçant")
                add_window.geometry("500x600")
                add_window.grab_set()  # Rendre la fenêtre modale
                
                # Variables pour les champs
                nom_var = StringVar()
                prenom_var = StringVar()
                email_var = StringVar()
                ddn_var = StringVar()
                lieu_naissance_var = StringVar()
                dept_naissance_var = StringVar()
                adresse_var = StringVar()
                secu_var = StringVar()
                nationalite_var = StringVar()
                iban_var = StringVar()
                sexe_var = StringVar(value="Monsieur")
                
                # Variables dérivées du sexe
                er_var = StringVar()
                ilr_var = StringVar()
                salarier_var = StringVar()
                
                # Fonction pour mettre à jour les champs dérivés du sexe
                def update_gender_fields(*args):
                    if sexe_var.get() == "Madame":
                        er_var.set("e")
                        ilr_var.set("elle")
                        salarier_var.set("à la salariée")
                    else:
                        er_var.set("")
                        ilr_var.set("il")
                        salarier_var.set("au salarié")
                
                # Lier la mise à jour automatique au changement de sexe
                sexe_var.trace_add("write", update_gender_fields)
                
                # Mise à jour initiale des champs dérivés
                update_gender_fields()
                
                # Création des champs
                padx, pady = 10, 5
                row = 0
                
                ttk.Label(add_window, text="Nom:").grid(row=row, column=0, sticky="w", padx=padx, pady=pady)
                ttk.Entry(add_window, textvariable=nom_var, width=30).grid(row=row, column=1, padx=padx, pady=pady)
                row += 1
                
                ttk.Label(add_window, text="Prénom:").grid(row=row, column=0, sticky="w", padx=padx, pady=pady)
                ttk.Entry(add_window, textvariable=prenom_var, width=30).grid(row=row, column=1, padx=padx, pady=pady)
                row += 1
                
                ttk.Label(add_window, text="Date de naissance:").grid(row=row, column=0, sticky="w", padx=padx, pady=pady)
                ttk.Entry(add_window, textvariable=ddn_var, width=30).grid(row=row, column=1, padx=padx, pady=pady)
                row += 1
                
                ttk.Label(add_window, text="Lieu de naissance:").grid(row=row, column=0, sticky="w", padx=padx, pady=pady)
                ttk.Entry(add_window, textvariable=lieu_naissance_var, width=30).grid(row=row, column=1, padx=padx, pady=pady)
                row += 1
                
                ttk.Label(add_window, text="Département de naissance:").grid(row=row, column=0, sticky="w", padx=padx, pady=pady)
                ttk.Entry(add_window, textvariable=dept_naissance_var, width=30).grid(row=row, column=1, padx=padx, pady=pady)
                row += 1
                
                ttk.Label(add_window, text="Nationalité:").grid(row=row, column=0, sticky="w", padx=padx, pady=pady)
                ttk.Entry(add_window, textvariable=nationalite_var, width=30).grid(row=row, column=1, padx=padx, pady=pady)
                row += 1
                
                ttk.Label(add_window, text="Adresse:").grid(row=row, column=0, sticky="w", padx=padx, pady=pady)
                ttk.Entry(add_window, textvariable=adresse_var, width=30).grid(row=row, column=1, padx=padx, pady=pady)
                row += 1
                
                ttk.Label(add_window, text="Email:").grid(row=row, column=0, sticky="w", padx=padx, pady=pady)
                ttk.Entry(add_window, textvariable=email_var, width=30).grid(row=row, column=1, padx=padx, pady=pady)
                row += 1
                
                ttk.Label(add_window, text="N° Sécurité Sociale:").grid(row=row, column=0, sticky="w", padx=padx, pady=pady)
                ttk.Entry(add_window, textvariable=secu_var, width=30).grid(row=row, column=1, padx=padx, pady=pady)
                row += 1
                
                ttk.Label(add_window, text="IBAN:").grid(row=row, column=0, sticky="w", padx=padx, pady=pady)
                ttk.Entry(add_window, textvariable=iban_var, width=30).grid(row=row, column=1, padx=padx, pady=pady)
                row += 1
                
                ttk.Label(add_window, text="Sexe:").grid(row=row, column=0, sticky="w", padx=padx, pady=pady)
                ttk.OptionMenu(add_window, sexe_var, "Monsieur", "Madame").grid(row=row, column=1, sticky="w", padx=padx, pady=pady)
                row += 1
                
                # Champs dérivés du sexe (non modifiables)
                ttk.Label(add_window, text="ER:").grid(row=row, column=0, sticky="w", padx=padx, pady=pady)
                ttk.Entry(add_window, textvariable=er_var, state="readonly", width=30).grid(row=row, column=1, padx=padx, pady=pady)
                row += 1
                
                ttk.Label(add_window, text="ILR:").grid(row=row, column=0, sticky="w", padx=padx, pady=pady)
                ttk.Entry(add_window, textvariable=ilr_var, state="readonly", width=30).grid(row=row, column=1, padx=padx, pady=pady)
                row += 1
                
                ttk.Label(add_window, text="SALARIER:").grid(row=row, column=0, sticky="w", padx=padx, pady=pady)
                ttk.Entry(add_window, textvariable=salarier_var, state="readonly", width=30).grid(row=row, column=1, padx=padx, pady=pady)
                row += 1
                
                def save_new():
                    """Enregistre le nouveau IADE et met à jour Excel."""
                    # Vérification des champs obligatoires
                    if not nom_var.get().strip() or not prenom_var.get().strip():
                        show_warning("Attention", "Veuillez renseigner au moins le nom et le prénom.")
                        return
                    
                    # Préparation des nouvelles données
                    new_row = {
                        "NOMR": nom_var.get(),
                        "PRENOMR": prenom_var.get(),
                        "EMAIL": email_var.get(),
                        "DDNR": ddn_var.get(),
                        "LIEUNR": lieu_naissance_var.get(),
                        "DPTN": dept_naissance_var.get(),
                        "ADRESSER": adresse_var.get(),
                        "NOSSR": secu_var.get(),
                        "NATR": nationalite_var.get(),
                        "SEXE": sexe_var.get(),
                        "ER": er_var.get(),
                        "ILR": ilr_var.get(),
                        "SALARIER": salarier_var.get()
                    }
                    
                    # Ajouter IBAN s'il n'existe pas dans le DataFrame
                    if "IBAN" not in iade_data.columns:
                        iade_data["IBAN"] = ""
                    new_row["IBAN"] = iban_var.get()
                    
                    try:
                        # Ajout au DataFrame
                        iade_data.loc[len(iade_data)] = new_row
                        
                        # Sauvegarde dans Excel
                        from modules.contracts.contrat53_new import save_excel_with_updated_sheet
                        save_excel_with_updated_sheet(excel_path, "Coordonnées IADEs", iade_data)
                        show_success("Succès", "IADE remplaçant ajouté avec succès.")
                        
                        # Mise à jour de l'affichage
                        refresh_listbox()
                        add_window.destroy()
                    except Exception as e:
                        show_error("Erreur", f"Erreur lors de l'ajout : {e}")
                
                # Boutons de validation
                buttons_frame = tk.Frame(add_window)
                buttons_frame.grid(row=row, column=0, columnspan=2, pady=15)
                
                ttk.Button(buttons_frame, text="Ajouter", command=save_new).pack(side="left", padx=10)
                ttk.Button(buttons_frame, text="Annuler", command=add_window.destroy).pack(side="left", padx=10)
            
            def on_delete():
                """Supprimer un IADE remplaçant."""
                selected_index = listbox.curselection()
                if not selected_index:
                    show_warning("Avertissement", "Veuillez sélectionner un IADE.")
                    return
                
                selected_index = selected_index[0]
                selected_row = iade_data_list[selected_index]
                
                # Confirmation de suppression
                nom = selected_row["NOMR"] if not pd.isna(selected_row["NOMR"]) else ""
                prenom = selected_row["PRENOMR"] if not pd.isna(selected_row["PRENOMR"]) else ""
                full_name = f"{nom} {prenom}".strip()
                
                if ask_question("Confirmation", f"Voulez-vous vraiment supprimer {full_name} ?"):
                    try:
                        # Suppression de la ligne
                        iade_data.drop(iade_data.index[selected_index], inplace=True)
                        iade_data.reset_index(drop=True, inplace=True)
                        
                        # Sauvegarde dans Excel
                        from modules.contracts.contrat53_new import save_excel_with_updated_sheet
                        save_excel_with_updated_sheet(excel_path, "Coordonnées IADEs", iade_data)
                        show_success("Succès", f"{full_name} supprimé avec succès.")
                        
                        # Mise à jour de l'affichage
                        refresh_listbox()
                    except Exception as e:
                        show_error("Erreur", f"Erreur lors de la suppression : {e}")
            
            # Boutons d'action
            buttons_frame = create_frame(management_frame)
            buttons_frame.pack(pady=10)
            
            create_button(
                buttons_frame, 
                text="➕ Ajouter", 
                command=on_add,
                style="primary",
                width=15
            ).pack(side="left", padx=10)
            
            create_button(
                buttons_frame, 
                text="✏️ Modifier", 
                command=on_modify,
                style="info",
                width=15
            ).pack(side="left", padx=10)
            
            create_button(
                buttons_frame, 
                text="🗑️ Supprimer", 
                command=on_delete,
                style="danger",
                width=15
            ).pack(side="left", padx=10)
            
            self.status_bar.set_message("Interface de gestion des IADE remplaçants ouverte.")
        except Exception as e:
            import traceback
            traceback.print_exc()
            show_error("Erreur", f"Impossible d'ouvrir l'interface de gestion des IADE remplaçants: {str(e)}")
            self.status_bar.set_message("Erreur lors de l'ouverture de l'interface de gestion des IADE remplaçants.")


        
    def manage_salaries(self):
        """
        Ouvre l'interface de gestion des salariés.
        """
        self.status_bar.set_message("Ouverture de l'interface de gestion des salariés...")
        
        # Nettoyage de l'interface existante si nécessaire
        for widget in self.settings_tab.winfo_children():
            if isinstance(widget, TabView):
                continue
            widget.destroy()
        
        # Créer un cadre pour l'interface de gestion des salariés
        management_frame = create_frame(self.settings_tab)
        management_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Bouton de retour
        return_frame = create_frame(management_frame, bg=COLORS["secondary"])
        return_frame.pack(fill="x", pady=(0, 10))
        
        create_button(
            return_frame,
            text="🔙 Retour aux paramètres",
            command=lambda: self.create_settings_tab(),
            style="neutral",
            width=20
        ).pack(side="left", padx=5, pady=5)
        
        # Titre
        create_label(
            management_frame,
            "Gestion des salariés",
            style="title"
        ).pack(pady=10)
        
        try:
            # Chargement des données depuis la bonne feuille
            excel_path = get_file_path("excel_salaries", verify_exists=True)
            if not excel_path:
                raise ValueError("Le chemin du fichier Excel des salariés n'est pas défini.")
            
            salaries_data = pd.read_excel(excel_path, sheet_name="Salariés")
            
            # Cadre principal
            main_frame = create_frame(management_frame, bg=COLORS["secondary"])
            main_frame.pack(fill="both", expand=True, padx=10, pady=10)
            
            # Créer une listbox pour afficher les salariés
            listbox = tk.Listbox(main_frame, width=50, height=15, font=("Arial", 12))
            listbox.pack(side="left", fill="both", expand=True, padx=5, pady=10)
            
            # Ajouter une scrollbar
            scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=listbox.yview)
            scrollbar.pack(side="right", fill="y")
            listbox.config(yscrollcommand=scrollbar.set)
            
            # Variable pour stocker les données
            salaries_list = []
            
            def refresh_listbox():
                """Met à jour la liste des salariés affichée dans la Listbox."""
                listbox.delete(0, tk.END)
                salaries_list.clear()
                
                for _, row in salaries_data.iterrows():
                    nom = row["NOM"] if not pd.isna(row["NOM"]) else ""
                    prenom = row["PRENOM"] if not pd.isna(row["PRENOM"]) else ""
                    poste = row.get("POSTE", "") if "POSTE" in row and not pd.isna(row["POSTE"]) else ""
                    full_name = f"{nom} {prenom}" + (f" - {poste}" if poste else "")
                    salaries_list.append(row)
                    listbox.insert(tk.END, full_name.strip())
            
            # Remplir la listbox initialement
            refresh_listbox()
            
            # Fonctions pour les boutons
            def on_modify():
                """Modifier un salarié."""
                selected_index = listbox.curselection()
                if not selected_index:
                    show_warning("Avertissement", "Veuillez sélectionner un salarié.")
                    return
                
                selected_index = selected_index[0]
                selected_row = salaries_list[selected_index]
                
                # Créer une fenêtre de modification
                modify_window = tk.Toplevel(management_frame.winfo_toplevel())
                modify_window.title("Modifier un salarié")
                modify_window.geometry("500x500")
                modify_window.grab_set()  # Rendre la fenêtre modale
                
                # Variables pour les champs
                nom_var = StringVar(value=selected_row["NOM"] if not pd.isna(selected_row["NOM"]) else "")
                prenom_var = StringVar(value=selected_row["PRENOM"] if not pd.isna(selected_row["PRENOM"]) else "")
                email_var = StringVar(value=selected_row["EMAIL"] if not pd.isna(selected_row["EMAIL"]) else "")
                poste_var = StringVar(value=selected_row.get("POSTE", "") if "POSTE" in selected_row and not pd.isna(selected_row["POSTE"]) else "")
                adresse_var = StringVar(value=selected_row.get("ADRESSE", "") if "ADRESSE" in selected_row and not pd.isna(selected_row["ADRESSE"]) else "")
                iban_var = StringVar(value=selected_row.get("IBAN", "") if "IBAN" in selected_row and not pd.isna(selected_row["IBAN"]) else "")
                
                # Création des champs
                padx, pady = 10, 5
                row = 0
                
                ttk.Label(modify_window, text="Nom:").grid(row=row, column=0, sticky="w", padx=padx, pady=pady)
                ttk.Entry(modify_window, textvariable=nom_var, width=30).grid(row=row, column=1, padx=padx, pady=pady)
                row += 1
                
                ttk.Label(modify_window, text="Prénom:").grid(row=row, column=0, sticky="w", padx=padx, pady=pady)
                ttk.Entry(modify_window, textvariable=prenom_var, width=30).grid(row=row, column=1, padx=padx, pady=pady)
                row += 1
                
                ttk.Label(modify_window, text="Email:").grid(row=row, column=0, sticky="w", padx=padx, pady=pady)
                ttk.Entry(modify_window, textvariable=email_var, width=30).grid(row=row, column=1, padx=padx, pady=pady)
                row += 1
                
                ttk.Label(modify_window, text="Poste:").grid(row=row, column=0, sticky="w", padx=padx, pady=pady)
                ttk.Entry(modify_window, textvariable=poste_var, width=30).grid(row=row, column=1, padx=padx, pady=pady)
                row += 1
                
                ttk.Label(modify_window, text="Adresse:").grid(row=row, column=0, sticky="w", padx=padx, pady=pady)
                ttk.Entry(modify_window, textvariable=adresse_var, width=30).grid(row=row, column=1, padx=padx, pady=pady)
                row += 1
                
                ttk.Label(modify_window, text="IBAN:").grid(row=row, column=0, sticky="w", padx=padx, pady=pady)
                ttk.Entry(modify_window, textvariable=iban_var, width=30).grid(row=row, column=1, padx=padx, pady=pady)
                row += 1
                
                def save_changes():
                    """Enregistre les modifications et met à jour Excel."""
                    # Mise à jour des données
                    salaries_data.loc[selected_index, "NOM"] = nom_var.get()
                    salaries_data.loc[selected_index, "PRENOM"] = prenom_var.get()
                    salaries_data.loc[selected_index, "EMAIL"] = email_var.get()
                    
                    # Mise à jour des champs optionnels
                    # Vérifier que les colonnes existent, sinon les créer
                    for col_name, var in [
                        ("POSTE", poste_var),
                        ("ADRESSE", adresse_var),
                        ("IBAN", iban_var)
                    ]:
                        if col_name not in salaries_data.columns:
                            salaries_data[col_name] = ""
                        salaries_data.loc[selected_index, col_name] = var.get()
                    
                    try:
                        # Sauvegarde dans Excel
                        from modules.contracts.contrat53_new import save_excel_with_updated_sheet
                        save_excel_with_updated_sheet(excel_path, "Salariés", salaries_data)
                        show_success("Succès", "Modifications enregistrées avec succès.")
                        
                        # Mise à jour de l'affichage
                        refresh_listbox()
                        modify_window.destroy()
                    except Exception as e:
                        show_error("Erreur", f"Erreur lors de la sauvegarde : {e}")
                
                # Boutons de validation
                buttons_frame = tk.Frame(modify_window)
                buttons_frame.grid(row=row, column=0, columnspan=2, pady=15)
                
                ttk.Button(buttons_frame, text="Enregistrer", command=save_changes).pack(side="left", padx=10)
                ttk.Button(buttons_frame, text="Annuler", command=modify_window.destroy).pack(side="left", padx=10)
            
            def on_add():
                """Ajouter un nouveau salarié."""
                # Créer une fenêtre d'ajout
                add_window = tk.Toplevel(management_frame.winfo_toplevel())
                add_window.title("Ajouter un salarié")
                add_window.geometry("500x500")
                add_window.grab_set()  # Rendre la fenêtre modale
                
                # Variables pour les champs
                nom_var = StringVar()
                prenom_var = StringVar()
                email_var = StringVar()
                poste_var = StringVar()
                adresse_var = StringVar()
                iban_var = StringVar()
                
                # Création des champs
                padx, pady = 10, 5
                row = 0
                
                ttk.Label(add_window, text="Nom:").grid(row=row, column=0, sticky="w", padx=padx, pady=pady)
                ttk.Entry(add_window, textvariable=nom_var, width=30).grid(row=row, column=1, padx=padx, pady=pady)
                row += 1
                
                ttk.Label(add_window, text="Prénom:").grid(row=row, column=0, sticky="w", padx=padx, pady=pady)
                ttk.Entry(add_window, textvariable=prenom_var, width=30).grid(row=row, column=1, padx=padx, pady=pady)
                row += 1
                
                ttk.Label(add_window, text="Email:").grid(row=row, column=0, sticky="w", padx=padx, pady=pady)
                ttk.Entry(add_window, textvariable=email_var, width=30).grid(row=row, column=1, padx=padx, pady=pady)
                row += 1
                
                ttk.Label(add_window, text="Poste:").grid(row=row, column=0, sticky="w", padx=padx, pady=pady)
                ttk.Entry(add_window, textvariable=poste_var, width=30).grid(row=row, column=1, padx=padx, pady=pady)
                row += 1
                
                ttk.Label(add_window, text="Adresse:").grid(row=row, column=0, sticky="w", padx=padx, pady=pady)
                ttk.Entry(add_window, textvariable=adresse_var, width=30).grid(row=row, column=1, padx=padx, pady=pady)
                row += 1
                
                ttk.Label(add_window, text="IBAN:").grid(row=row, column=0, sticky="w", padx=padx, pady=pady)
                ttk.Entry(add_window, textvariable=iban_var, width=30).grid(row=row, column=1, padx=padx, pady=pady)
                row += 1
                
                def save_new():
                    """Enregistre le nouveau salarié et met à jour Excel."""
                    # Vérification des champs obligatoires
                    if not nom_var.get().strip() or not prenom_var.get().strip():
                        show_warning("Attention", "Veuillez renseigner au moins le nom et le prénom.")
                        return
                    
                    # Préparation des nouvelles données
                    new_row = {
                        "NOM": nom_var.get(),
                        "PRENOM": prenom_var.get(),
                        "EMAIL": email_var.get()
                    }
                    
                    # Ajout des champs optionnels
                    for col_name, var in [
                        ("POSTE", poste_var),
                        ("ADRESSE", adresse_var),
                        ("IBAN", iban_var)
                    ]:
                        if col_name not in salaries_data.columns:
                            salaries_data[col_name] = ""
                        new_row[col_name] = var.get()
                    
                    try:
                        # Ajout au DataFrame
                        salaries_data.loc[len(salaries_data)] = new_row
                        
                        # Sauvegarde dans Excel
                        from modules.contracts.contrat53_new import save_excel_with_updated_sheet
                        save_excel_with_updated_sheet(excel_path, "Salariés", salaries_data)
                        show_success("Succès", "Salarié ajouté avec succès.")
                        
                        # Mise à jour de l'affichage
                        refresh_listbox()
                        add_window.destroy()
                    except Exception as e:
                        show_error("Erreur", f"Erreur lors de l'ajout : {e}")
                
                # Boutons de validation
                buttons_frame = tk.Frame(add_window)
                buttons_frame.grid(row=row, column=0, columnspan=2, pady=15)
                
                ttk.Button(buttons_frame, text="Ajouter", command=save_new).pack(side="left", padx=10)
                ttk.Button(buttons_frame, text="Annuler", command=add_window.destroy).pack(side="left", padx=10)
            
            def on_delete():
                """Supprimer un salarié."""
                selected_index = listbox.curselection()
                if not selected_index:
                    show_warning("Avertissement", "Veuillez sélectionner un salarié.")
                    return
                
                selected_index = selected_index[0]
                selected_row = salaries_list[selected_index]
                
                # Confirmation de suppression
                nom = selected_row["NOM"] if not pd.isna(selected_row["NOM"]) else ""
                prenom = selected_row["PRENOM"] if not pd.isna(selected_row["PRENOM"]) else ""
                full_name = f"{nom} {prenom}".strip()
                
                if ask_question("Confirmation", f"Voulez-vous vraiment supprimer {full_name} ?"):
                    try:
                        # Suppression de la ligne
                        salaries_data.drop(salaries_data.index[selected_index], inplace=True)
                        salaries_data.reset_index(drop=True, inplace=True)
                        
                        # Sauvegarde dans Excel
                        from modules.contracts.contrat53_new import save_excel_with_updated_sheet
                        save_excel_with_updated_sheet(excel_path, "Salariés", salaries_data)
                        show_success("Succès", f"{full_name} supprimé avec succès.")
                        
                        # Mise à jour de l'affichage
                        refresh_listbox()
                    except Exception as e:
                        show_error("Erreur", f"Erreur lors de la suppression : {e}")
            
            # Boutons d'action
            buttons_frame = create_frame(management_frame)
            buttons_frame.pack(pady=10)
            
            create_button(
                buttons_frame, 
                text="➕ Ajouter", 
                command=on_add,
                style="primary",
                width=15
            ).pack(side="left", padx=10)
            
            create_button(
                buttons_frame, 
                text="✏️ Modifier", 
                command=on_modify,
                style="info",
                width=15
            ).pack(side="left", padx=10)
            
            create_button(
                buttons_frame, 
                text="🗑️ Supprimer", 
                command=on_delete,
                style="danger",
                width=15
            ).pack(side="left", padx=10)
            
            self.status_bar.set_message("Interface de gestion des salariés ouverte.")
        except Exception as e:
            import traceback
            traceback.print_exc()
            show_error("Erreur", f"Impossible d'ouvrir l'interface de gestion des salariés: {str(e)}")
            
        
def save_excel_with_updated_sheet(file_path, sheet_name, updated_data):
    """
    Sauvegarde une feuille spécifique dans un fichier Excel sans supprimer les autres feuilles.
    
    Args:
        file_path (str): Chemin du fichier Excel à mettre à jour
        sheet_name (str): Nom de la feuille à mettre à jour
        updated_data (pd.DataFrame): DataFrame contenant les données à sauvegarder
        
    Returns:
        bool: True si la sauvegarde a réussi, False sinon
    """
    print(f"📝 Sauvegarde de la feuille '{sheet_name}' dans le fichier '{file_path}'...")
    
    # Vérification des paramètres
    if not file_path or not os.path.exists(file_path):
        print(f"❌ Erreur : Le fichier '{file_path}' n'existe pas.")
        show_error("Erreur", f"Le fichier '{file_path}' n'existe pas.") if 'show_error' in globals() else None
        return False
    
    if not sheet_name:
        print("❌ Erreur : Nom de feuille non spécifié.")
        show_error("Erreur", "Nom de feuille non spécifié.") if 'show_error' in globals() else None
        return False
    
    if not isinstance(updated_data, pd.DataFrame):
        print("❌ Erreur : Les données à sauvegarder ne sont pas un DataFrame pandas.")
        show_error("Erreur", "Les données à sauvegarder ne sont pas au bon format.") if 'show_error' in globals() else None
        return False
    
    try:
        # Vérifier si le fichier est ouvert dans Excel
        try:
            # Tentative d'ouverture du fichier en mode 'r+b' pour vérifier s'il est verrouillé
            with open(file_path, 'r+b') as f:
                pass
        except IOError:
            print(f"⚠️ Le fichier {file_path} est actuellement ouvert dans une autre application.")
            show_warning("Fichier verrouillé", 
                        f"Le fichier {os.path.basename(file_path)} est actuellement ouvert.\n"
                        "Veuillez le fermer avant de sauvegarder.", 
                        parent=None) if 'show_warning' in globals() else None
            return False
        
        # Sauvegarde du fichier original en cas de problème
        backup_path = f"{file_path}.bak"
        try:
            import shutil
            shutil.copy2(file_path, backup_path)
            print(f"✅ Backup créé : {backup_path}")
        except Exception as e:
            print(f"⚠️ Impossible de créer un backup : {str(e)}")
        
        # Charger toutes les feuilles existantes
        sheets = {}
        try:
            with pd.ExcelFile(file_path, engine="openpyxl") as xls:
                # Afficher les feuilles trouvées
                print(f"📊 Feuilles trouvées : {xls.sheet_names}")
                
                # Charger chaque feuille
                for sheet in xls.sheet_names:
                    print(f"📄 Chargement de la feuille '{sheet}'...")
                    sheets[sheet] = xls.parse(sheet)
        except Exception as e:
            print(f"❌ Erreur lors du chargement des feuilles : {str(e)}")
            show_error("Erreur", f"Impossible de lire le fichier Excel : {str(e)}") if 'show_error' in globals() else None
            return False
        
        # Vérifier si la feuille à mettre à jour existe
        if sheet_name not in sheets:
            print(f"⚠️ La feuille '{sheet_name}' n'existe pas dans le fichier. Elle sera créée.")
        else:
            print(f"✅ La feuille '{sheet_name}' existe et va être mise à jour.")
            # Afficher un résumé des modifications
            old_rows = len(sheets[sheet_name])
            new_rows = len(updated_data)
            print(f"📊 Modification de {old_rows} → {new_rows} lignes")
        
        # Mettre à jour la feuille spécifiée
        sheets[sheet_name] = updated_data
        
        # Réécrire le fichier avec toutes les feuilles
        try:
            print(f"💾 Écriture du fichier avec {len(sheets)} feuilles...")
            with pd.ExcelWriter(file_path, engine="openpyxl", mode="w") as writer:
                for sheet, df in sheets.items():
                    print(f"📄 Écriture de la feuille '{sheet}' ({len(df)} lignes)...")
                    df.to_excel(writer, sheet_name=sheet, index=False)
            print(f"✅ Fichier sauvegardé avec succès : {file_path}")
            
            # Supprimer le backup si tout s'est bien passé
            if os.path.exists(backup_path):
                try:
                    os.remove(backup_path)
                    print("✅ Backup supprimé car sauvegarde réussie.")
                except:
                    print("⚠️ Impossible de supprimer le backup.")
            
            # Afficher un message de succès
            show_success("Sauvegarde réussie", 
                        f"La feuille '{sheet_name}' a été mise à jour avec succès dans\n{os.path.basename(file_path)}") if 'show_success' in globals() else None
            
            return True
            
        except Exception as e:
            print(f"❌ Erreur lors de l'écriture du fichier : {str(e)}")
            
            # Restaurer le backup si la sauvegarde a échoué
            if os.path.exists(backup_path):
                try:
                    import shutil
                    shutil.copy2(backup_path, file_path)
                    print("✅ Fichier restauré à partir du backup.")
                except Exception as restore_error:
                    print(f"⚠️ Impossible de restaurer le backup : {str(restore_error)}")
            
            show_error("Erreur", f"Impossible de sauvegarder le fichier Excel : {str(e)}") if 'show_error' in globals() else None
            return False
            
    except Exception as e:
        print(f"❌ Erreur inattendue lors de la sauvegarde : {str(e)}")
        import traceback
        traceback.print_exc()
        show_error("Erreur", f"Une erreur inattendue s'est produite : {str(e)}") if 'show_error' in globals() else None
        return False

def main():
    """
    Point d'entrée principal de l'application.
    """
    app = MainApplication()
    app.mainloop()

if __name__ == "__main__":
    main()