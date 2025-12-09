#!/usr/bin/env python3
"""
Matomo ARK Extractor
Extraction des statistiques ARK depuis les exports Matomo XML
avec récupération des métadonnées du catalogue Portfolio

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

# Interface moderne
import customtkinter as ctk
from tkinter import filedialog, messagebox
from CTkTable import CTkTable

# Traitement données
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows

# Scraping
import httpx
from bs4 import BeautifulSoup
import asyncio

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


class MatomoARKExtractor(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        # Configuration fenêtre
        self.title("📚 Matomo ARK Extractor - Bibliothèques spécialisées Paris")
        self.geometry("1100x750")
        self.minsize(900, 600)
        
        # Variables
        self.xml_path = ctk.StringVar()
        self.status_text = ctk.StringVar(value="Sélectionnez un fichier XML Matomo")
        self.progress_value = ctk.DoubleVar(value=0)
        self.scrape_metadata = ctk.BooleanVar(value=True)
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
            text="Extraction des statistiques de consultation des ressources ARK",
            font=ctk.CTkFont(size=14),
            text_color=COLORS['text_muted']
        ).pack(pady=(5, 0))
        
        # Badges info
        badges_frame = ctk.CTkFrame(header_frame, fg_color="transparent")
        badges_frame.pack(pady=(0, 15))
        
        for text, color in [("Bibliothèques spécialisées", COLORS['primary']), 
                            ("Ville de Paris", COLORS['accent'])]:
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
        
        # Checkbox pour métadonnées
        self.metadata_check = ctk.CTkCheckBox(
            inner_frame,
            text="Tenter de récupérer les métadonnées (titre, auteur...) depuis le catalogue",
            variable=self.scrape_metadata,
            font=ctk.CTkFont(size=13),
            checkbox_height=22,
            checkbox_width=22,
            corner_radius=5
        )
        self.metadata_check.pack(anchor="w", pady=(10, 0))
        
        ctk.CTkLabel(
            inner_frame,
            text="⚠️ Cette option peut prendre plusieurs minutes selon le nombre de notices",
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
            text="📊 Résultats",
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
        self.log("En attente d'un fichier XML...")
    
    def create_footer(self):
        footer_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        footer_frame.pack(fill="x")
        
        ctk.CTkLabel(
            footer_frame,
            text="CCPID - Bibliothèques de la Ville de Paris",
            font=ctk.CTkFont(size=11),
            text_color=COLORS['text_muted']
        ).pack(side="left")
        
        # Lien GitHub
        github_btn = ctk.CTkButton(
            footer_frame,
            text="GitHub",
            font=ctk.CTkFont(size=11),
            fg_color="transparent",
            hover_color=COLORS['secondary'],
            text_color=COLORS['accent'],
            width=60,
            height=25,
            command=lambda: webbrowser.open("https://github.com")
        )
        github_btn.pack(side="right")
    
    def log(self, message, level="INFO"):
        timestamp = datetime.now().strftime("%H:%M:%S")
        icons = {"INFO": "ℹ️", "SUCCESS": "✅", "ERROR": "❌", "WARNING": "⚠️", "PROGRESS": "🔄"}
        icon = icons.get(level, "")
        
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
            self.ark_data = self.parse_xml(self.xml_path.get())
            
            if not self.ark_data:
                self.log("Aucune donnée ARK trouvée dans le fichier", "ERROR")
                return
            
            self.log(f"Trouvé {len(self.ark_data)} ressources ARK uniques", "SUCCESS")
            self.count_label.configure(text=f"{len(self.ark_data)} ressources")
            
            # 2. Récupérer les métadonnées si demandé
            if self.scrape_metadata.get():
                self.status_text.set("Récupération des métadonnées...")
                self.progress_value.set(0.3)
                asyncio.run(self.fetch_all_metadata())
            
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
        
        ark_items = []
        
        def extract_rows(element):
            for row in element.findall('.//row'):
                url_elem = row.find('url')
                if url_elem is not None and url_elem.text:
                    url = url_elem.text
                    if '/ark:/' in url or '/ark%3A/' in url:
                        # Extraire l'ARK
                        ark_match = re.search(r'ark[:%3A/]+(\d+)/([a-zA-Z0-9\-]+)', url, re.IGNORECASE)
                        if ark_match:
                            naan = ark_match.group(1)
                            ark_id = ark_match.group(2)
                            
                            # Nettoyer l'URL
                            clean_url = re.sub(r'\?.*$', '', url)
                            clean_url = re.sub(r'/v\d+\.\w+.*$', '', clean_url)
                            
                            ark_items.append({
                                'ark': f'ark:/{naan}/{ark_id}',
                                'ark_id': ark_id,
                                'naan': naan,
                                'url': clean_url,
                                'nb_visits': int(row.findtext('nb_visits', '0') or 0),
                                'nb_uniq_visitors': row.findtext('nb_uniq_visitors', ''),
                                'nb_hits': int(row.findtext('nb_hits', '0') or 0),
                                'sum_time_spent': int(row.findtext('sum_time_spent', '0') or 0),
                                'avg_time_on_page': row.findtext('avg_time_on_page', ''),
                                'bounce_rate': row.findtext('bounce_rate', ''),
                                'exit_rate': row.findtext('exit_rate', ''),
                                'entry_nb_visits': row.findtext('entry_nb_visits', ''),
                                # Métadonnées (à remplir)
                                'titre': '',
                                'auteur': '',
                                'editeur': '',
                                'date': '',
                                'type': '',
                                'bibliotheque': ''
                            })
        
        extract_rows(root)
        
        # Agréger par ARK unique
        aggregated = defaultdict(lambda: {
            'visits': 0, 'hits': 0, 'sum_time': 0, 
            'urls': set(), 'data': None
        })
        
        for item in ark_items:
            ark = item['ark']
            aggregated[ark]['visits'] += item['nb_visits']
            aggregated[ark]['hits'] += item['nb_hits']
            aggregated[ark]['sum_time'] += item['sum_time_spent']
            aggregated[ark]['urls'].add(item['url'])
            if aggregated[ark]['data'] is None:
                aggregated[ark]['data'] = item
        
        # Construire la liste finale
        result = []
        for ark, agg in aggregated.items():
            data = agg['data'].copy()
            data['nb_visits'] = agg['visits']
            data['nb_hits'] = agg['hits']
            data['sum_time_spent'] = agg['sum_time']
            data['url'] = min(agg['urls'], key=len)  # URL la plus courte
            result.append(data)
        
        # Trier par visites
        result.sort(key=lambda x: x['nb_visits'], reverse=True)
        
        # Logger les top 5
        for i, item in enumerate(result[:5], 1):
            self.log(f"  #{i}: {item['ark_id']} - {item['nb_visits']} visites")
        
        return result
    
    async def fetch_metadata(self, session, item, semaphore):
        """Récupère les métadonnées d'une notice"""
        async with semaphore:
            url = item['url']
            try:
                headers = {
                    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
                    'Accept': 'text/html,application/xhtml+xml',
                    'Accept-Language': 'fr-FR,fr;q=0.9'
                }
                
                response = await session.get(url, headers=headers, timeout=15, follow_redirects=True)
                
                if response.status_code == 200:
                    html = response.text
                    soup = BeautifulSoup(html, 'html.parser')
                    
                    # Titre depuis <title> ou meta
                    title_tag = soup.find('title')
                    if title_tag:
                        title = title_tag.get_text().strip()
                        title = re.sub(r'\s*[-|–].*$', '', title)
                        item['titre'] = title[:200]
                    
                    # Meta description
                    meta_desc = soup.find('meta', {'name': 'description'})
                    if meta_desc and meta_desc.get('content'):
                        if not item['titre']:
                            item['titre'] = meta_desc['content'][:200]
                    
                    # Chercher dans le contenu structuré
                    for script in soup.find_all('script', type='application/ld+json'):
                        try:
                            import json
                            data = json.loads(script.string)
                            if isinstance(data, dict):
                                if 'name' in data and not item['titre']:
                                    item['titre'] = data['name'][:200]
                                if 'author' in data:
                                    author = data['author']
                                    if isinstance(author, dict):
                                        item['auteur'] = author.get('name', '')
                                    else:
                                        item['auteur'] = str(author)[:100]
                        except:
                            pass
                    
                    # Déterminer le type depuis l'ARK
                    ark_id = item['ark_id']
                    if ark_id.startswith('FRCGMNOV'):
                        item['type'] = 'Fonds iconographique - Nouvelles'
                    elif ark_id.startswith('FRCGMSUP'):
                        item['type'] = 'Fonds iconographique - Suppléments'
                    elif ark_id.startswith('pf'):
                        item['type'] = 'Notice bibliographique'
                    
            except Exception as e:
                pass  # Silencieux en cas d'erreur
            
            return item
    
    async def fetch_all_metadata(self):
        """Récupère les métadonnées pour toutes les notices"""
        self.log(f"Récupération des métadonnées pour {len(self.ark_data)} notices...")
        
        semaphore = asyncio.Semaphore(5)  # Max 5 requêtes simultanées
        
        async with httpx.AsyncClient() as session:
            tasks = []
            for i, item in enumerate(self.ark_data):
                tasks.append(self.fetch_metadata(session, item, semaphore))
                
                # Mise à jour progression
                if i % 10 == 0:
                    progress = 0.3 + (i / len(self.ark_data)) * 0.5
                    self.progress_value.set(progress)
                    self.status_text.set(f"Métadonnées: {i+1}/{len(self.ark_data)}")
            
            results = await asyncio.gather(*tasks)
            
            # Compter les titres trouvés
            titles_found = sum(1 for r in results if r.get('titre'))
            self.log(f"Métadonnées récupérées: {titles_found} titres trouvés", "SUCCESS")
    
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
        border = Border(
            left=Side(style='thin', color='cccccc'),
            right=Side(style='thin', color='cccccc'),
            top=Side(style='thin', color='cccccc'),
            bottom=Side(style='thin', color='cccccc')
        )
        link_font = Font(color='0563C1', underline='single')
        
        # En-têtes
        headers = [
            'Rang', 'ARK', 'ID', 'Type', 'Titre', 'Auteur', 
            'Visites', 'Visiteurs uniques', 'Pages vues', 
            'Temps total (s)', 'Temps moyen', 'Taux rebond', 'URL'
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
                item['nb_visits'],
                item.get('nb_uniq_visitors', ''),
                item['nb_hits'],
                item['sum_time_spent'],
                item.get('avg_time_on_page', ''),
                item.get('bounce_rate', ''),
                item['url']
            ]
            
            for col, value in enumerate(values, 1):
                cell = ws.cell(row=row, column=col, value=value)
                cell.border = border
                
                if idx % 2 == 0:
                    cell.fill = alt_fill
                
                if col == 13:  # URL
                    cell.font = link_font
                    cell.hyperlink = value
                elif col in [7, 8, 9, 10]:  # Numériques
                    cell.alignment = Alignment(horizontal='right')
        
        # Largeurs colonnes
        widths = [6, 30, 28, 28, 45, 25, 10, 16, 12, 14, 12, 12, 70]
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w
        
        # Filtre et gel
        ws.auto_filter.ref = f"A1:M{len(self.ark_data)+1}"
        ws.freeze_panes = 'A2'
        
        # === Feuille 2: Résumé ===
        ws2 = wb.create_sheet("Résumé")
        
        ws2['A1'] = "📊 Résumé des statistiques"
        ws2['A1'].font = Font(bold=True, size=16)
        
        ws2['A3'] = "Date d'extraction:"
        ws2['B3'] = datetime.now().strftime("%d/%m/%Y %H:%M")
        
        ws2['A4'] = "Fichier source:"
        ws2['B4'] = os.path.basename(self.xml_path.get())
        
        ws2['A6'] = "Statistiques globales"
        ws2['A6'].font = Font(bold=True, size=12)
        
        ws2['A7'] = "Nombre de ressources ARK:"
        ws2['B7'] = len(self.ark_data)
        
        ws2['A8'] = "Total des visites:"
        ws2['B8'] = sum(d['nb_visits'] for d in self.ark_data)
        
        ws2['A9'] = "Total des pages vues:"
        ws2['B9'] = sum(d['nb_hits'] for d in self.ark_data)
        
        # Par type
        ws2['A11'] = "Par type de ressource"
        ws2['A11'].font = Font(bold=True, size=12)
        
        type_counts = defaultdict(lambda: {'count': 0, 'visits': 0})
        for item in self.ark_data:
            t = item.get('type', 'Autre') or 'Autre'
            type_counts[t]['count'] += 1
            type_counts[t]['visits'] += item['nb_visits']
        
        row = 12
        for t, data in sorted(type_counts.items(), key=lambda x: x[1]['visits'], reverse=True):
            ws2.cell(row=row, column=1, value=t)
            ws2.cell(row=row, column=2, value=f"{data['count']} ressources")
            ws2.cell(row=row, column=3, value=f"{data['visits']} visites")
            row += 1
        
        ws2.column_dimensions['A'].width = 35
        ws2.column_dimensions['B'].width = 20
        ws2.column_dimensions['C'].width = 15
        
        # === Feuille 3: Top 20 ===
        ws3 = wb.create_sheet("Top 20")
        
        ws3['A1'] = "🏆 Top 20 des ressources les plus consultées"
        ws3['A1'].font = Font(bold=True, size=14)
        
        headers3 = ['Rang', 'Titre / ARK', 'Visites', 'Pages vues']
        for col, h in enumerate(headers3, 1):
            cell = ws3.cell(row=3, column=col, value=h)
            cell.font = Font(bold=True)
            cell.fill = PatternFill('solid', fgColor='d9e2f3')
        
        for idx, item in enumerate(self.ark_data[:20], 1):
            title = item.get('titre') or item['ark_id']
            ws3.cell(row=idx+3, column=1, value=idx)
            ws3.cell(row=idx+3, column=2, value=title[:60])
            ws3.cell(row=idx+3, column=3, value=item['nb_visits'])
            ws3.cell(row=idx+3, column=4, value=item['nb_hits'])
        
        ws3.column_dimensions['A'].width = 8
        ws3.column_dimensions['B'].width = 60
        ws3.column_dimensions['C'].width = 12
        ws3.column_dimensions['D'].width = 12
        
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
                self.ark_data = self.parse_xml(self.xml_path.get())
            except Exception as e:
                messagebox.showerror("Erreur", f"Impossible de lire le fichier:\n{str(e)}")
                return
        
        # Créer fenêtre d'aperçu
        preview_window = ctk.CTkToplevel(self)
        preview_window.title("Aperçu des données")
        preview_window.geometry("900x500")
        
        # Tableau
        columns = ['#', 'ARK ID', 'Visites', 'Pages vues', 'Type']
        data = [[i+1, d['ark_id'][:25], d['nb_visits'], d['nb_hits'], 
                 d.get('type', '')[:20]] for i, d in enumerate(self.ark_data[:50])]
        
        table_frame = ctk.CTkScrollableFrame(preview_window)
        table_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # Header
        for col, header in enumerate(columns):
            ctk.CTkLabel(
                table_frame, 
                text=header,
                font=ctk.CTkFont(weight="bold"),
                width=150 if col > 0 else 50
            ).grid(row=0, column=col, padx=5, pady=5)
        
        # Data rows
        for row_idx, row_data in enumerate(data, 1):
            for col_idx, value in enumerate(row_data):
                ctk.CTkLabel(
                    table_frame,
                    text=str(value),
                    width=150 if col_idx > 0 else 50
                ).grid(row=row_idx, column=col_idx, padx=5, pady=2)


def main():
    app = MatomoARKExtractor()
    app.mainloop()


if __name__ == '__main__':
    main()
