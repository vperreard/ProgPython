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
            self.open_invoice_analysis()
        except Exception as e:
            logger.error(f"Erreur lors de l'initialisation des interfaces: {str(e)}")
            
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
        
        FormField(
            parent, 
            "Mot de passe (laisser vide si non stocké) :", 
            variable=self.docusign_vars["password"]
        ).pack(fill="x", padx=50, pady=5)
        
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
        """
        Ouvre l'interface d'analyse des factures.
        """
        self.status_bar.set_message("Ouverture de l'interface d'analyse des factures...")
        
        # Nettoyer l'interface existante
        for widget in self.invoices_frame.winfo_children():
            widget.destroy()
        
        # Créer un cadre pour l'interface d'analyse des factures
        analysis_frame = create_frame(self.invoices_frame)
        analysis_frame.pack(fill="both", expand=True)
        
        # Créer l'instance de l'analyseur
        factures_path = get_file_path("dossier_factures", verify_exists=True)
        analyseur = analyse_facture.AnalyseFactures(factures_path)
        
        # Créer l'interface de l'analyseur dans le cadre
        analyseur.creer_interface(analysis_frame)
        
        self.status_bar.set_message("Interface d'analyse des factures ouverte.")
    
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
        
        # Appeler la fonction de gestion des MAR titulaires
        contrat53.manage_mar_titulaires()
        
        self.status_bar.set_message("Interface de gestion des MAR titulaires ouverte.")
    
    def manage_mar_remplacants(self):
        """
        Ouvre l'interface de gestion des MAR remplaçants.
        """
        self.status_bar.set_message("Ouverture de l'interface de gestion des MAR remplaçants...")
        
        # Appeler la fonction de gestion des MAR remplaçants
        contrat53.manage_mar_remplacants()
        
        self.status_bar.set_message("Interface de gestion des MAR remplaçants ouverte.")
    
    def manage_iade_remplacants(self):
        """
        Ouvre l'interface de gestion des IADE remplaçants.
        """
        self.status_bar.set_message("Ouverture de l'interface de gestion des IADE remplaçants...")
        
        # Appeler la fonction de gestion des IADE remplaçants
        contrat53.manage_iade_remplacants()
        
        self.status_bar.set_message("Interface de gestion des IADE remplaçants ouverte.")
    
    def manage_salaries(self):
        """
        Ouvre l'interface de gestion des salariés.
        """
        self.status_bar.set_message("Ouverture de l'interface de gestion des salariés...")
        
        # Appeler la fonction de gestion des salariés
        contrat53.manage_salaries()
        
        self.status_bar.set_message("Interface de gestion des salariés ouverte.")

def main():
    """
    Point d'entrée principal de l'application.
    """
    app = MainApplication()
    app.mainloop()

if __name__ == "__main__":
    main()