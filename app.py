#!/usr/bin/env python3
"""
Matomo ARK Extractor v2.0
Extraction des statistiques ARK depuis les exports Matomo XML
avec récupération des métadonnées via l'API OAI-PMH du catalogue Portfolio

Bibliothèques spécialisées de la Ville de Paris
"""

import os
import sys
import re
import threading
import xml.etree.ElementTree as ET
from datetime import datetime
from collections import defaultdict
from pathlib import Path
import webbrowser
import urllib.parse

# Interface moderne
import customtkinter as ctk
from tkinter import filedialog, messagebox

# Traitement données
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Requêtes HTTP
import httpx

# Configuration CustomTkinter
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

# Couleurs personnalisées
COLORS = {
    'primary': '#1f538d',
    'secondary': '#14375e', 
    'accent': '#3a7ebf',
    'success': '#2d8659',
    'warning': '#c9a227',
    'error': '#bf3636',
    'bg_dark': '#1a1a2e',
    'bg_card': '#16213e',
    'text': '#e8e8e8',
    'text_muted': '#a0a0a0'
}

# Configuration OAI-PMH
OAI_BASE_URL = "https://bibliotheques-specialisees.paris.fr/in/rest/oai"


class MatomoARKExtractor(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        # Configuration fenêtre
        self.title("📚 Matomo ARK Extractor v2.0 - Bibliothèques spécialisées Paris")
        self.geometry("1100x800")
        self.minsize(900, 650)
        
        # Variables
        self.xml_path = ctk.StringVar()
        self.status_text = ctk.StringVar(value="Sélectionnez un fichier XML Matomo")
        self.progress_value = ctk.DoubleVar(value=0)
        self.scrape_metadata = ctk.BooleanVar(value=True)
        self.include_components = ctk.BooleanVar(value=False)
        self.ark_data = []
        self.is_processing = False
        
        # Interface
        self.create_ui()
        
    def create_ui(self):
        # Frame principal avec padding
        self.main_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.main_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # Header
        self.create_header()
        
        # Zone de sélection fichier
        self.create_file_selector()
        
        # Options
        self.create_options()
        
        # Boutons d'action
        self.create_action_buttons()
        
        # Barre de progression
        self.create_progress_section()
        
        # Zone de résultats/aperçu
        self.create_results_section()
        
        # Footer
        self.create_footer()
    
    def create_header(self):
        header_frame = ctk.CTkFrame(self.main_frame, fg_color=COLORS['bg_card'], corner_radius=15)
        header_frame.pack(fill="x", pady=(0, 15))
        
        # Titre avec icône
        title_frame = ctk.CTkFrame(header_frame, fg_color="transparent")
        title_frame.pack(pady=20, padx=20)
        
        ctk.CTkLabel(
            title_frame, 
            text="📚 Matomo ARK Extractor",
            font=ctk.CTkFont(size=28, weight="bold"),
            text_color=COLORS['text']
        ).pack()
        
        ctk.CTkLabel(
            title_frame,
            text="Extraction des statistiques + métadonnées via API OAI-PMH",
            font=ctk.CTkFont(size=14),
            text_color=COLORS['text_muted']
        ).pack(pady=(5, 0))
        
        # Badges info
        badges_frame = ctk.CTkFrame(header_frame, fg_color="transparent")
        badges_frame.pack(pady=(0, 15))
        
        for text, color in [("Bibliothèques spécialisées", COLORS['primary']), 
                            ("Ville de Paris", COLORS['accent']),
                            ("v2.0 - OAI-PMH", COLORS['success'])]:
            badge = ctk.CTkLabel(
                badges_frame,
                text=text,
                font=ctk.CTkFont(size=11),
                fg_color=color,
                corner_radius=12,
                padx=12,
                pady=4
            )
            badge.pack(side="left", padx=5)
    
    def create_file_selector(self):
        file_frame = ctk.CTkFrame(self.main_frame, fg_color=COLORS['bg_card'], corner_radius=15)
        file_frame.pack(fill="x", pady=(0, 15))
        
        inner_frame = ctk.CTkFrame(file_frame, fg_color="transparent")
        inner_frame.pack(fill="x", padx=20, pady=20)
        
        ctk.CTkLabel(
            inner_frame,
            text="📁 Fichier XML Matomo",
            font=ctk.CTkFont(size=16, weight="bold"),
            text_color=COLORS['text']
        ).pack(anchor="w")
        
        # Ligne de sélection
        select_frame = ctk.CTkFrame(inner_frame, fg_color="transparent")
        select_frame.pack(fill="x", pady=(10, 0))
        
        self.file_entry = ctk.CTkEntry(
            select_frame,
            textvariable=self.xml_path,
            placeholder_text="Cliquez sur 'Parcourir' pour sélectionner le fichier XML...",
            font=ctk.CTkFont(size=13),
            height=45,
            corner_radius=10
        )
        self.file_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
        
        self.browse_btn = ctk.CTkButton(
            select_frame,
            text="📂 Parcourir",
            font=ctk.CTkFont(size=14, weight="bold"),
            height=45,
            width=140,
            corner_radius=10,
            fg_color=COLORS['primary'],
            hover_color=COLORS['accent'],
            command=self.browse_file
        )
        self.browse_btn.pack(side="right")
    
    def create_options(self):
        options_frame = ctk.CTkFrame(self.main_frame, fg_color=COLORS['bg_card'], corner_radius=15)
        options_frame.pack(fill="x", pady=(0, 15))
        
        inner_frame = ctk.CTkFrame(options_frame, fg_color="transparent")
        inner_frame.pack(fill="x", padx=20, pady=15)
        
        ctk.CTkLabel(
            inner_frame,
            text="⚙️ Options",
            font=ctk.CTkFont(size=16, weight="bold"),
            text_color=COLORS['text']
        ).pack(anchor="w")
        
        # Checkbox pour métadonnées OAI-PMH
        self.metadata_check = ctk.CTkCheckBox(
            inner_frame,
            text="Récupérer les métadonnées via API OAI-PMH (titre, auteur, date, type...)",
            variable=self.scrape_metadata,
            font=ctk.CTkFont(size=13),
            checkbox_height=22,
            checkbox_width=22,
            corner_radius=5
        )
        self.metadata_check.pack(anchor="w", pady=(10, 0))
        
        ctk.CTkLabel(
            inner_frame,
            text="✅ Utilise l'API officielle du catalogue - fiable et rapide",
            font=ctk.CTkFont(size=11),
            text_color=COLORS['success']
        ).pack(anchor="w", padx=(28, 0), pady=(2, 0))
        
        # Checkbox pour composantes
        self.components_check = ctk.CTkCheckBox(
            inner_frame,
            text="Inclure les composantes/vues (BAP..., pages numérisées)",
            variable=self.include_components,
            font=ctk.CTkFont(size=13),
            checkbox_height=22,
            checkbox_width=22,
            corner_radius=5
        )
        self.components_check.pack(anchor="w", pady=(10, 0))
        
        ctk.CTkLabel(
            inner_frame,
            text="📄 Crée une feuille supplémentaire avec le détail par composante",
            font=ctk.CTkFont(size=11),
            text_color=COLORS['text_muted']
        ).pack(anchor="w", padx=(28, 0), pady=(2, 0))
    
    def create_action_buttons(self):
        btn_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        btn_frame.pack(fill="x", pady=(0, 15))
        
        # Bouton principal
        self.run_btn = ctk.CTkButton(
            btn_frame,
            text="▶️  Extraire et générer l'Excel",
            font=ctk.CTkFont(size=16, weight="bold"),
            height=55,
            corner_radius=12,
            fg_color=COLORS['success'],
            hover_color="#238c4d",
            command=self.start_extraction
        )
        self.run_btn.pack(side="left", fill="x", expand=True, padx=(0, 10))
        
        # Bouton aperçu
        self.preview_btn = ctk.CTkButton(
            btn_frame,
            text="👁️ Aperçu",
            font=ctk.CTkFont(size=14),
            height=55,
            width=120,
            corner_radius=12,
            fg_color=COLORS['secondary'],
            hover_color=COLORS['primary'],
            command=self.show_preview
        )
        self.preview_btn.pack(side="right")
    
    def create_progress_section(self):
        self.progress_frame = ctk.CTkFrame(self.main_frame, fg_color=COLORS['bg_card'], corner_radius=15)
        self.progress_frame.pack(fill="x", pady=(0, 15))
        
        inner_frame = ctk.CTkFrame(self.progress_frame, fg_color="transparent")
        inner_frame.pack(fill="x", padx=20, pady=15)
        
        # Status
        self.status_label = ctk.CTkLabel(
            inner_frame,
            textvariable=self.status_text,
            font=ctk.CTkFont(size=13),
            text_color=COLORS['text']
        )
        self.status_label.pack(anchor="w")
        
        # Barre de progression
        self.progress_bar = ctk.CTkProgressBar(
            inner_frame,
            variable=self.progress_value,
            height=12,
            corner_radius=6,
            progress_color=COLORS['accent']
        )
        self.progress_bar.pack(fill="x", pady=(10, 0))
        self.progress_bar.set(0)
    
    def create_results_section(self):
        results_frame = ctk.CTkFrame(self.main_frame, fg_color=COLORS['bg_card'], corner_radius=15)
        results_frame.pack(fill="both", expand=True, pady=(0, 15))
        
        inner_frame = ctk.CTkFrame(results_frame, fg_color="transparent")
        inner_frame.pack(fill="both", expand=True, padx=20, pady=15)
        
        # Titre section
        header_row = ctk.CTkFrame(inner_frame, fg_color="transparent")
        header_row.pack(fill="x")
        
        ctk.CTkLabel(
            header_row,
            text="📊 Journal",
            font=ctk.CTkFont(size=16, weight="bold"),
            text_color=COLORS['text']
        ).pack(side="left")
        
        self.count_label = ctk.CTkLabel(
            header_row,
            text="",
            font=ctk.CTkFont(size=13),
            text_color=COLORS['text_muted']
        )
        self.count_label.pack(side="right")
        
        # Zone de texte pour les logs
        self.log_textbox = ctk.CTkTextbox(
            inner_frame,
            font=ctk.CTkFont(family="Consolas", size=12),
            corner_radius=10,
            fg_color="#0d1117",
            text_color="#c9d1d9"
        )
        self.log_textbox.pack(fill="both", expand=True, pady=(10, 0))
        self.log("🚀 Prêt ! Sélectionnez un fichier XML Matomo pour commencer.")
        self.log("")
        self.log("ℹ️  Cette version utilise l'API OAI-PMH pour récupérer les métadonnées.")
        self.log("   Endpoint: " + OAI_BASE_URL)
    
    def create_footer(self):
        footer_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        footer_frame.pack(fill="x")
        
        ctk.CTkLabel(
            footer_frame,
            text="CCPID - Bibliothèques de la Ville de Paris",
            font=ctk.CTkFont(size=11),
            text_color=COLORS['text_muted']
        ).pack(side="left")
        
        ctk.CTkLabel(
            footer_frame,
            text="v2.0 - OAI-PMH",
            font=ctk.CTkFont(size=11),
            text_color=COLORS['accent']
        ).pack(side="right")
    
    def log(self, message, level="INFO"):
        timestamp = datetime.now().strftime("%H:%M:%S")
        icons = {"INFO": "ℹ️", "SUCCESS": "✅", "ERROR": "❌", "WARNING": "⚠️", "PROGRESS": "🔄", "DATA": "📄"}
        icon = icons.get(level, "")
        
        if level == "INFO" and message.startswith("ℹ️"):
            # Déjà formaté
            self.log_textbox.insert("end", f"{message}\n")
        elif level == "INFO" and message == "":
            self.log_textbox.insert("end", "\n")
        else:
            self.log_textbox.insert("end", f"[{timestamp}] {icon} {message}\n")
        
        self.log_textbox.see("end")
        self.update_idletasks()
    
    def browse_file(self):
        filename = filedialog.askopenfilename(
            title="Sélectionner le fichier XML Matomo",
            filetypes=[("XML files", "*.xml"), ("All files", "*.*")]
        )
        if filename:
            self.xml_path.set(filename)
            self.log(f"Fichier sélectionné: {Path(filename).name}", "SUCCESS")
            self.status_text.set(f"Fichier: {Path(filename).name}")
    
    def start_extraction(self):
        if not self.xml_path.get():
            messagebox.showwarning("Attention", "Veuillez sélectionner un fichier XML")
            return
        
        if not os.path.exists(self.xml_path.get()):
            messagebox.showerror("Erreur", "Le fichier sélectionné n'existe pas")
            return
        
        if self.is_processing:
            return
        
        self.is_processing = True
        self.run_btn.configure(state="disabled", text="⏳ Traitement en cours...")
        self.browse_btn.configure(state="disabled")
        self.progress_bar.set(0)
        self.log_textbox.delete("1.0", "end")
        
        # Lancer dans un thread
        thread = threading.Thread(target=self.extraction_thread, daemon=True)
        thread.start()
    
    def extraction_thread(self):
        try:
            self.log("Démarrage de l'extraction...", "PROGRESS")
            
            # 1. Parser le XML
            self.status_text.set("Analyse du fichier XML...")
            self.progress_value.set(0.1)
            self.ark_data, self.components_data = self.parse_xml(self.xml_path.get())
            
            if not self.ark_data:
                self.log("Aucune donnée ARK trouvée dans le fichier", "ERROR")
                return
            
            self.log(f"Trouvé {len(self.ark_data)} notices ARK uniques", "SUCCESS")
            if self.components_data:
                self.log(f"Trouvé {len(self.components_data)} composantes/vues", "SUCCESS")
            self.count_label.configure(text=f"{len(self.ark_data)} notices")
            
            # 2. Récupérer les métadonnées si demandé
            if self.scrape_metadata.get():
                self.status_text.set("Récupération des métadonnées via OAI-PMH...")
                self.progress_value.set(0.2)
                self.fetch_oai_metadata()
            
            # 3. Générer l'Excel
            self.status_text.set("Génération du fichier Excel...")
            self.progress_value.set(0.9)
            output_path = self.generate_excel()
            
            self.progress_value.set(1.0)
            self.status_text.set("Terminé !")
            self.log(f"Fichier Excel généré: {Path(output_path).name}", "SUCCESS")
            
            # Message de succès
            messagebox.showinfo(
                "Extraction terminée",
                f"Le fichier Excel a été créé:\n\n{output_path}"
            )
            
            # Ouvrir le dossier
            os.startfile(os.path.dirname(output_path))
            
        except Exception as e:
            self.log(f"Erreur: {str(e)}", "ERROR")
            import traceback
            self.log(traceback.format_exc(), "ERROR")
            messagebox.showerror("Erreur", f"Une erreur s'est produite:\n{str(e)}")
        
        finally:
            self.is_processing = False
            self.run_btn.configure(state="normal", text="▶️  Extraire et générer l'Excel")
            self.browse_btn.configure(state="normal")
    
    def parse_xml(self, xml_path):
        """Parse le fichier XML Matomo et extrait les données ARK"""
        self.log("Parsing du fichier XML...")
        
        tree = ET.parse(xml_path)
        root = tree.getroot()
        
        notices = []  # Niveau notice (pf..., FRCGM...)
        components = []  # Niveau composante (BAP..., vues...)
        
        def get_type_from_ark(ark_id):
            """Détermine le type de ressource depuis l'identifiant ARK"""
            if ark_id.startswith('FRCGMNOV'):
                return 'Fonds iconographique - Nouvelles'
            elif ark_id.startswith('FRCGMSUP'):
                return 'Fonds iconographique - Suppléments'
            elif ark_id.startswith('FRCGM'):
                return 'Fonds iconographique'
            elif ark_id.startswith('pf'):
                return 'Notice bibliographique'
            else:
                return 'Autre'
        
        def extract_rows(element, parent_ark=None):
            for row in element.findall('.//row'):
                label = row.findtext('label', '')
                url_elem = row.find('url')
                url = url_elem.text if url_elem is not None else None
                
                # Données Matomo
                data = {
                    'nb_visits': int(row.findtext('nb_visits', '0') or 0),
                    'nb_uniq_visitors': row.findtext('nb_uniq_visitors', ''),
                    'nb_hits': int(row.findtext('nb_hits', '0') or 0),
                    'sum_time_spent': int(row.findtext('sum_time_spent', '0') or 0),
                    'avg_time_on_page': row.findtext('avg_time_on_page', ''),
                    'bounce_rate': row.findtext('bounce_rate', ''),
                    'exit_rate': row.findtext('exit_rate', ''),
                    'entry_nb_visits': row.findtext('entry_nb_visits', ''),
                    'entry_bounce_count': row.findtext('entry_bounce_count', ''),
                    'exit_nb_visits': row.findtext('exit_nb_visits', ''),
                }
                
                # Détection du type de ligne
                if url and '/ark:/' in url:
                    # Extraire l'ARK de l'URL
                    ark_match = re.search(r'ark:/(\d+)/([a-zA-Z0-9\-]+)(?:/([a-zA-Z0-9\-]+))?', url)
                    if ark_match:
                        naan = ark_match.group(1)
                        ark_id = ark_match.group(2)
                        component_id = ark_match.group(3)  # Peut être None
                        
                        ark_full = f"ark:/{naan}/{ark_id}"
                        
                        if component_id and (component_id.startswith('BAP') or 
                                            component_id.startswith('BHP') or
                                            component_id.isdigit() or
                                            re.match(r'^\d{4}$', component_id)):
                            # C'est une composante
                            components.append({
                                'ark_notice': ark_full,
                                'component_id': component_id,
                                'url': url,
                                **data
                            })
                        else:
                            # C'est une notice
                            # Nettoyer l'URL des paramètres de vue
                            clean_url = re.sub(r'/v\d+\..*$', '', url)
                            clean_url = re.sub(r'\?.*$', '', clean_url)
                            
                            notices.append({
                                'ark': ark_full,
                                'ark_id': ark_id,
                                'naan': naan,
                                'url': clean_url,
                                'type': get_type_from_ark(ark_id),
                                # Métadonnées (à remplir via OAI)
                                'titre': '',
                                'auteur': '',
                                'date': '',
                                'editeur': '',
                                'description': '',
                                'bibliotheque': '',
                                'cote': '',
                                **data
                            })
                
                elif label and not label.startswith('/') and not label.isdigit() and label != 'ark:':
                    # Peut être un identifiant de notice au niveau 3
                    if re.match(r'^(pf|FRCGM)', label):
                        ark_full = f"ark:/73873/{label}"
                        notices.append({
                            'ark': ark_full,
                            'ark_id': label,
                            'naan': '73873',
                            'url': f"https://bibliotheques-specialisees.paris.fr/ark:/73873/{label}",
                            'type': get_type_from_ark(label),
                            'titre': '',
                            'auteur': '',
                            'date': '',
                            'editeur': '',
                            'description': '',
                            'bibliotheque': '',
                            'cote': '',
                            **data
                        })
        
        extract_rows(root)
        
        # Agréger par ARK unique (notices)
        aggregated = defaultdict(lambda: {
            'visits': 0, 'hits': 0, 'sum_time': 0,
            'entry_visits': 0, 'entry_bounces': 0, 'exit_visits': 0,
            'data': None
        })
        
        for item in notices:
            ark = item['ark']
            aggregated[ark]['visits'] += item['nb_visits']
            aggregated[ark]['hits'] += item['nb_hits']
            aggregated[ark]['sum_time'] += item['sum_time_spent']
            try:
                aggregated[ark]['entry_visits'] += int(item.get('entry_nb_visits') or 0)
                aggregated[ark]['entry_bounces'] += int(item.get('entry_bounce_count') or 0)
                aggregated[ark]['exit_visits'] += int(item.get('exit_nb_visits') or 0)
            except:
                pass
            if aggregated[ark]['data'] is None:
                aggregated[ark]['data'] = item
        
        # Construire la liste finale des notices
        result_notices = []
        for ark, agg in aggregated.items():
            data = agg['data'].copy()
            data['nb_visits'] = agg['visits']
            data['nb_hits'] = agg['hits']
            data['sum_time_spent'] = agg['sum_time']
            data['entry_nb_visits'] = agg['entry_visits']
            data['entry_bounce_count'] = agg['entry_bounces']
            data['exit_nb_visits'] = agg['exit_visits']
            result_notices.append(data)
        
        # Trier par visites
        result_notices.sort(key=lambda x: x['nb_visits'], reverse=True)
        
        # Logger les top 5
        self.log("Top 5 des notices les plus consultées:")
        for i, item in enumerate(result_notices[:5], 1):
            self.log(f"  #{i}: {item['ark_id']} - {item['nb_visits']} visites", "DATA")
        
        return result_notices, components
    
    def fetch_oai_metadata(self):
        """Récupère les métadonnées via l'API OAI-PMH"""
        total = len(self.ark_data)
        self.log(f"Récupération des métadonnées pour {total} notices via OAI-PMH...")
        self.log(f"Endpoint: {OAI_BASE_URL}")
        
        success_count = 0
        error_count = 0
        no_record_count = 0
        
        with httpx.Client(timeout=30.0, follow_redirects=True) as client:
            for i, item in enumerate(self.ark_data):
                # Mise à jour progression
                progress = 0.2 + (i / total) * 0.6
                self.progress_value.set(progress)
                self.status_text.set(f"Métadonnées: {i+1}/{total} - {item['ark_id'][:20]}...")
                
                try:
                    # Construire l'URL OAI-PMH - format inmedia d'abord (plus riche)
                    ark_identifier = item['ark']
                    oai_url = f"{OAI_BASE_URL}?verb=GetRecord&identifier={ark_identifier}&metadataPrefix=inmedia"
                    
                    response = client.get(oai_url)
                    
                    if response.status_code == 200:
                        # Vérifier si c'est une erreur "idDoesNotExist"
                        if 'idDoesNotExist' in response.text or 'noRecordsMatch' in response.text:
                            no_record_count += 1
                            # Essayer avec oai_dc
                            oai_url2 = f"{OAI_BASE_URL}?verb=GetRecord&identifier={ark_identifier}&metadataPrefix=oai_dc"
                            response2 = client.get(oai_url2)
                            if response2.status_code == 200 and 'idDoesNotExist' not in response2.text:
                                metadata = self.parse_oai_response(response2.text, format='oai_dc')
                                if metadata and metadata.get('title'):
                                    item['titre'] = metadata.get('title', '')
                                    item['auteur'] = metadata.get('creator', '')
                                    item['date'] = metadata.get('date', '')
                                    item['editeur'] = metadata.get('publisher', '')
                                    item['description'] = metadata.get('description', '')[:200] if metadata.get('description') else ''
                                    item['type_oai'] = metadata.get('type', '')
                                    success_count += 1
                        else:
                            metadata = self.parse_oai_response(response.text, format='inmedia')
                            
                            if metadata:
                                item['titre'] = metadata.get('title', '')
                                item['auteur'] = metadata.get('creator', '')
                                item['date'] = metadata.get('date', '')
                                item['editeur'] = metadata.get('publisher', '')
                                item['description'] = metadata.get('description', '')[:200] if metadata.get('description') else ''
                                item['type_oai'] = metadata.get('type', '')
                                item['cote'] = metadata.get('identifier', '')
                                item['bibliotheque'] = metadata.get('source', '')
                                
                                if metadata.get('title'):
                                    success_count += 1
                                    # Log les premiers succès pour feedback
                                    if success_count <= 3:
                                        self.log(f"  ✓ {item['ark_id']}: {item['titre'][:50]}...", "DATA")
                    else:
                        error_count += 1
                    
                except Exception as e:
                    error_count += 1
                    if error_count <= 5:
                        self.log(f"Erreur pour {item['ark_id']}: {str(e)[:50]}", "WARNING")
        
        self.log(f"Métadonnées récupérées: {success_count} titres trouvés sur {total}", "SUCCESS")
        if no_record_count > 0:
            self.log(f"Notices non trouvées dans OAI: {no_record_count} (peut-être non numérisées)", "WARNING")
        if error_count > 0:
            self.log(f"Erreurs réseau: {error_count}", "WARNING")
    
    def parse_oai_response(self, xml_text, format='oai_dc'):
        """Parse la réponse XML OAI-PMH pour extraire les métadonnées Dublin Core ou InMedia"""
        try:
            # Namespaces OAI-PMH et Dublin Core
            namespaces = {
                'oai': 'http://www.openarchives.org/OAI/2.0/',
                'dc': 'http://purl.org/dc/elements/1.1/',
                'oai_dc': 'http://www.openarchives.org/OAI/2.0/oai_dc/',
                'dcterms': 'http://purl.org/dc/terms/'
            }
            
            root = ET.fromstring(xml_text)
            metadata = {}
            
            # Liste des champs Dublin Core à chercher
            dc_fields = ['title', 'creator', 'date', 'publisher', 'description', 'type', 'subject', 'identifier', 'source', 'format', 'rights']
            
            for dc_elem in dc_fields:
                found_value = None
                
                # Méthode 1: Avec namespace dc:
                elem = root.find(f'.//dc:{dc_elem}', namespaces)
                if elem is not None and elem.text:
                    found_value = elem.text.strip()
                
                # Méthode 2: Avec namespace complet
                if not found_value:
                    elem = root.find(f'.//{{{namespaces["dc"]}}}{dc_elem}')
                    if elem is not None and elem.text:
                        found_value = elem.text.strip()
                
                # Méthode 3: Chercher dans tout l'arbre (sans namespace)
                if not found_value:
                    for e in root.iter():
                        tag_local = e.tag.split('}')[-1] if '}' in e.tag else e.tag
                        if tag_local.lower() == dc_elem.lower() and e.text:
                            found_value = e.text.strip()
                            break
                
                # Méthode 4: Format InMedia spécifique (balises en majuscules parfois)
                if not found_value and format == 'inmedia':
                    for e in root.iter():
                        tag_local = e.tag.split('}')[-1] if '}' in e.tag else e.tag
                        # InMedia utilise parfois Title, Creator, etc.
                        if tag_local.lower() == dc_elem.lower() and e.text:
                            found_value = e.text.strip()
                            break
                
                if found_value:
                    metadata[dc_elem] = found_value
            
            # Chercher aussi des champs spécifiques InMedia
            if format == 'inmedia':
                # Chercher auteur dans différents champs
                if not metadata.get('creator'):
                    for tag in ['author', 'Author', 'contributor', 'Contributor']:
                        for e in root.iter():
                            tag_local = e.tag.split('}')[-1] if '}' in e.tag else e.tag
                            if tag_local == tag and e.text:
                                metadata['creator'] = e.text.strip()
                                break
                        if metadata.get('creator'):
                            break
                
                # Chercher la cote
                for e in root.iter():
                    tag_local = e.tag.split('}')[-1] if '}' in e.tag else e.tag
                    if tag_local.lower() in ['shelfmark', 'callnumber', 'cote'] and e.text:
                        metadata['identifier'] = e.text.strip()
                        break
            
            return metadata if metadata else None
            
        except Exception as e:
            return None
    
    def generate_excel(self):
        """Génère le fichier Excel avec toutes les données"""
        self.log("Génération du fichier Excel...")
        
        # Chemin de sortie horodaté
        xml_dir = os.path.dirname(self.xml_path.get())
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_filename = f"stats_matomo_ark_{timestamp}.xlsx"
        output_path = os.path.join(xml_dir, output_filename)
        
        wb = Workbook()
        
        # === Feuille 1: Données principales ===
        ws = wb.active
        ws.title = "Statistiques ARK"
        
        # Styles
        header_font = Font(bold=True, color='FFFFFF', size=11)
        header_fill = PatternFill('solid', fgColor='1f538d')
        alt_fill = PatternFill('solid', fgColor='e8f0fe')
        success_fill = PatternFill('solid', fgColor='d4edda')
        border = Border(
            left=Side(style='thin', color='cccccc'),
            right=Side(style='thin', color='cccccc'),
            top=Side(style='thin', color='cccccc'),
            bottom=Side(style='thin', color='cccccc')
        )
        link_font = Font(color='0563C1', underline='single')
        
        # En-têtes
        headers = [
            'Rang', 'ARK complet', 'ID', 'Type', 'Titre', 'Auteur', 'Date',
            'Visites', 'Visiteurs uniques', 'Pages vues', 
            'Temps total (s)', 'Temps moyen', 'Taux rebond', 'Taux sortie',
            'Entrées', 'Sorties', 'URL'
        ]
        
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=header)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = border
        
        ws.row_dimensions[1].height = 25
        
        # Données
        for idx, item in enumerate(self.ark_data, 1):
            row = idx + 1
            values = [
                idx,
                item['ark'],
                item['ark_id'],
                item.get('type', ''),
                item.get('titre', ''),
                item.get('auteur', ''),
                item.get('date', ''),
                item['nb_visits'],
                item.get('nb_uniq_visitors', ''),
                item['nb_hits'],
                item['sum_time_spent'],
                item.get('avg_time_on_page', ''),
                item.get('bounce_rate', ''),
                item.get('exit_rate', ''),
                item.get('entry_nb_visits', ''),
                item.get('exit_nb_visits', ''),
                item['url']
            ]
            
            for col, value in enumerate(values, 1):
                cell = ws.cell(row=row, column=col, value=value)
                cell.border = border
                
                if idx % 2 == 0:
                    cell.fill = alt_fill
                
                # Surligner si titre trouvé
                if col == 5 and value:  # Titre
                    cell.fill = success_fill
                
                if col == 17:  # URL
                    cell.font = link_font
                    cell.hyperlink = value
                elif col in [8, 9, 10, 11, 15, 16]:  # Numériques
                    cell.alignment = Alignment(horizontal='right')
        
        # Largeurs colonnes
        widths = [6, 32, 30, 28, 50, 30, 12, 10, 16, 12, 14, 12, 12, 12, 10, 10, 70]
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w
        
        # Filtre et gel
        ws.auto_filter.ref = f"A1:Q{len(self.ark_data)+1}"
        ws.freeze_panes = 'A2'
        
        # === Feuille 2: Résumé ===
        ws2 = wb.create_sheet("Résumé")
        
        ws2['A1'] = "📊 Résumé des statistiques"
        ws2['A1'].font = Font(bold=True, size=16)
        
        ws2['A3'] = "Date d'extraction:"
        ws2['B3'] = datetime.now().strftime("%d/%m/%Y %H:%M")
        
        ws2['A4'] = "Fichier source:"
        ws2['B4'] = os.path.basename(self.xml_path.get())
        
        ws2['A5'] = "Méthode métadonnées:"
        ws2['B5'] = "API OAI-PMH" if self.scrape_metadata.get() else "Non activée"
        
        ws2['A7'] = "Statistiques globales"
        ws2['A7'].font = Font(bold=True, size=12)
        
        ws2['A8'] = "Nombre de notices ARK:"
        ws2['B8'] = len(self.ark_data)
        
        ws2['A9'] = "Notices avec titre:"
        ws2['B9'] = sum(1 for d in self.ark_data if d.get('titre'))
        
        ws2['A10'] = "Total des visites:"
        ws2['B10'] = sum(d['nb_visits'] for d in self.ark_data)
        
        ws2['A11'] = "Total des pages vues:"
        ws2['B11'] = sum(d['nb_hits'] for d in self.ark_data)
        
        # Par type
        ws2['A13'] = "Par type de ressource"
        ws2['A13'].font = Font(bold=True, size=12)
        
        type_counts = defaultdict(lambda: {'count': 0, 'visits': 0, 'with_title': 0})
        for item in self.ark_data:
            t = item.get('type', 'Autre') or 'Autre'
            type_counts[t]['count'] += 1
            type_counts[t]['visits'] += item['nb_visits']
            if item.get('titre'):
                type_counts[t]['with_title'] += 1
        
        row = 14
        ws2.cell(row=row, column=1, value="Type").font = Font(bold=True)
        ws2.cell(row=row, column=2, value="Notices").font = Font(bold=True)
        ws2.cell(row=row, column=3, value="Avec titre").font = Font(bold=True)
        ws2.cell(row=row, column=4, value="Visites").font = Font(bold=True)
        row += 1
        
        for t, data in sorted(type_counts.items(), key=lambda x: x[1]['visits'], reverse=True):
            ws2.cell(row=row, column=1, value=t)
            ws2.cell(row=row, column=2, value=data['count'])
            ws2.cell(row=row, column=3, value=data['with_title'])
            ws2.cell(row=row, column=4, value=data['visits'])
            row += 1
        
        ws2.column_dimensions['A'].width = 35
        ws2.column_dimensions['B'].width = 15
        ws2.column_dimensions['C'].width = 15
        ws2.column_dimensions['D'].width = 15
        
        # === Feuille 3: Top 20 ===
        ws3 = wb.create_sheet("Top 20")
        
        ws3['A1'] = "🏆 Top 20 des ressources les plus consultées"
        ws3['A1'].font = Font(bold=True, size=14)
        
        headers3 = ['Rang', 'Titre / ARK', 'Type', 'Auteur', 'Visites', 'Pages vues']
        for col, h in enumerate(headers3, 1):
            cell = ws3.cell(row=3, column=col, value=h)
            cell.font = Font(bold=True)
            cell.fill = PatternFill('solid', fgColor='d9e2f3')
        
        for idx, item in enumerate(self.ark_data[:20], 1):
            title = item.get('titre') or item['ark_id']
            ws3.cell(row=idx+3, column=1, value=idx)
            ws3.cell(row=idx+3, column=2, value=title[:60])
            ws3.cell(row=idx+3, column=3, value=item.get('type', '')[:25])
            ws3.cell(row=idx+3, column=4, value=item.get('auteur', '')[:30])
            ws3.cell(row=idx+3, column=5, value=item['nb_visits'])
            ws3.cell(row=idx+3, column=6, value=item['nb_hits'])
        
        ws3.column_dimensions['A'].width = 8
        ws3.column_dimensions['B'].width = 60
        ws3.column_dimensions['C'].width = 28
        ws3.column_dimensions['D'].width = 30
        ws3.column_dimensions['E'].width = 12
        ws3.column_dimensions['F'].width = 12
        
        # === Feuille 4: Composantes (si activé) ===
        if self.include_components.get() and self.components_data:
            ws4 = wb.create_sheet("Composantes")
            
            ws4['A1'] = "📄 Détail par composante / vue"
            ws4['A1'].font = Font(bold=True, size=14)
            
            headers4 = ['ARK Notice', 'Composante', 'Visites', 'Pages vues', 'Temps (s)', 'URL']
            for col, h in enumerate(headers4, 1):
                cell = ws4.cell(row=3, column=col, value=h)
                cell.font = Font(bold=True)
                cell.fill = PatternFill('solid', fgColor='d9e2f3')
            
            for idx, comp in enumerate(self.components_data[:500], 1):  # Limiter à 500
                ws4.cell(row=idx+3, column=1, value=comp.get('ark_notice', ''))
                ws4.cell(row=idx+3, column=2, value=comp.get('component_id', ''))
                ws4.cell(row=idx+3, column=3, value=comp.get('nb_visits', 0))
                ws4.cell(row=idx+3, column=4, value=comp.get('nb_hits', 0))
                ws4.cell(row=idx+3, column=5, value=comp.get('sum_time_spent', 0))
                ws4.cell(row=idx+3, column=6, value=comp.get('url', ''))
            
            ws4.column_dimensions['A'].width = 35
            ws4.column_dimensions['B'].width = 20
            ws4.column_dimensions['C'].width = 10
            ws4.column_dimensions['D'].width = 12
            ws4.column_dimensions['E'].width = 12
            ws4.column_dimensions['F'].width = 70
        
        # Sauvegarder
        wb.save(output_path)
        
        return output_path
    
    def show_preview(self):
        """Affiche un aperçu des données"""
        if not self.xml_path.get():
            messagebox.showwarning("Attention", "Veuillez d'abord sélectionner un fichier XML")
            return
        
        if not self.ark_data:
            # Parser rapidement pour l'aperçu
            try:
                self.ark_data, self.components_data = self.parse_xml(self.xml_path.get())
            except Exception as e:
                messagebox.showerror("Erreur", f"Impossible de lire le fichier:\n{str(e)}")
                return
        
        # Créer fenêtre d'aperçu
        preview_window = ctk.CTkToplevel(self)
        preview_window.title("Aperçu des données")
        preview_window.geometry("950x550")
        
        # Frame scrollable
        table_frame = ctk.CTkScrollableFrame(preview_window)
        table_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # Header
        columns = ['#', 'ARK ID', 'Type', 'Visites', 'Hits', 'Titre']
        for col, header in enumerate(columns):
            ctk.CTkLabel(
                table_frame, 
                text=header,
                font=ctk.CTkFont(weight="bold"),
                width=150 if col > 1 else 50
            ).grid(row=0, column=col, padx=5, pady=8)
        
        # Data rows (max 50)
        for row_idx, item in enumerate(self.ark_data[:50], 1):
            data_row = [
                row_idx,
                item['ark_id'][:25],
                item.get('type', '')[:20],
                item['nb_visits'],
                item['nb_hits'],
                item.get('titre', '-')[:40] or '-'
            ]
            for col_idx, value in enumerate(data_row):
                ctk.CTkLabel(
                    table_frame,
                    text=str(value),
                    width=150 if col_idx > 1 else 50
                ).grid(row=row_idx, column=col_idx, padx=5, pady=2)


def main():
    app = MatomoARKExtractor()
    app.mainloop()


if __name__ == '__main__':
    main()
