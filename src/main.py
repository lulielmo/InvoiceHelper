import os
import json
import logging
import re
from datetime import datetime
from pathlib import Path
import pandas as pd
import pytesseract
from pdf2image import convert_from_path
from PIL import Image
from pytesseract import Output

# Konfigurera logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(os.path.join('logs', f'invoice_helper_{datetime.now().strftime("%Y%m%d")}.log')),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

class InvoiceHelper:
    def __init__(self):
        self.users_file = 'data/users.xlsx'
        self.output_dir = 'output'
        self.users_data = None
        self.project_settings = None
        
        # Skapa output-mapp om den inte finns
        os.makedirs(self.output_dir, exist_ok=True)
        
        # Läs in användardata och projektinställningar
        self._load_excel_data()

    def _load_excel_data(self):
        """Läser in data från Excel-filen."""
        try:
            self.users_data = pd.read_excel(self.users_file, sheet_name='Power BI Users')
            self.project_settings = pd.read_excel(self.users_file, sheet_name='Project Settings')
            
            # Validera Project Settings-data
            required_columns = ['ProjektID', 'Kon/Proj', 'Aktivitet', 'ProjKat', 'Mottagare']
            missing_columns = [col for col in required_columns if col not in self.project_settings.columns]
            if missing_columns:
                logger.warning(f"Varning: Följande kolumner saknas i Project Settings: {', '.join(missing_columns)}")
            
            # Normalisera ProjektID (ta bort 'P.' om det finns)
            self.project_settings['ProjektID'] = self.project_settings['ProjektID'].astype(str).apply(
                lambda x: x.replace('P.', '') if x.startswith('P.') else x
            )
            
            logger.info("Excel-data inläst framgångsrikt")
            
            # Logga tillgängliga projekt för felsökning
            logger.info("Tillgängliga projekt i Project Settings:")
            for _, row in self.project_settings.iterrows():
                logger.info(f"- ProjektID: {row['ProjektID']}, Kon/Proj: {row['Kon/Proj']}, Aktivitet: {row['Aktivitet']}")
            
        except Exception as e:
            logger.error(f"Fel vid inläsning av Excel-data: {str(e)}")
            raise

    def get_project_settings(self, project_id):
        """Hämtar projektinställningar för ett specifikt projekt."""
        # Ta bort 'P.' från project_id om det finns
        search_id = project_id.replace('P.', '') if project_id.startswith('P.') else project_id
        
        project = self.project_settings[self.project_settings['ProjektID'] == search_id]
        
        if project.empty:
            logger.warning(f"Kunde inte hitta inställningar för projekt {project_id}")
            # Returnera standardvärden baserat på projekttyp
            default_settings = {
                '20257601': {  # Automation
                    'Kon/Proj': 'P.20257601',
                    'Aktivitet': '050',
                    'ProjKat': '5420',
                    'Mottagare': 'Digital Utveckling och integration'
                },
                '20257407': {  # Microsoft 365
                    'Kon/Proj': 'P.20257407',
                    'Aktivitet': '738',
                    'ProjKat': '5420',
                    'Mottagare': 'Digital Arbetsplats'
                },
                '20257403': {  # Teams Room
                    'Kon/Proj': 'P.20257403',
                    'Aktivitet': '738',
                    'ProjKat': '5420',
                    'Mottagare': 'Digital Arbetsplats'
                }
            }
            return default_settings.get(search_id, {
                'Kon/Proj': f'P.{search_id}',
                'Aktivitet': '738',
                'ProjKat': '5420',
                'Mottagare': 'Okänd mottagare'
            })
        
        return {
            'Kon/Proj': f"P.{project.iloc[0]['ProjektID']}",
            'Aktivitet': project.iloc[0]['Aktivitet'],
            'ProjKat': project.iloc[0]['ProjKat'],
            'Mottagare': project.iloc[0]['Mottagare']
        }

    def extract_text_from_pdf(self, pdf_path):
        """Extraherar text från PDF med OCR."""
        try:
            # Konvertera PDF till bilder
            images = convert_from_path(pdf_path)
            
            # Extrahera text från varje sida
            text_content = []
            for i, image in enumerate(images):
                logger.info(f"Processar sida {i+1}")
                text = pytesseract.image_to_string(image, lang='swe')
                text_content.append(text)
            
            return '\n'.join(text_content)
        except Exception as e:
            logger.error(f"Fel vid PDF-läsning: {str(e)}")
            raise

    def parse_license_info(self, text):
        """Extraherar licensinformation från OCR-texten."""
        try:
            logger.info("Börjar parsning av licensinformation")
            
            # Definiera licenstyper att leta efter
            license_patterns = {
                'power_bi': r'CSP -Power BI Pro \(cycle\)\s+\d{6}\s+-\s+\d{6}\s+(\d+,\d+)\s+ST\s+(\d+[\s,]*\d*,\d+)\s+(\d+[\s,]*\d*,\d+)',
                'power_automate_rpa': r'CSP -Power Automate unattended RPA add-on \(Cycle\)\s+\d{6}\s+-\s+\d{6}\s+(\d+,\d+)\s+ST\s+(\d+[\s,]*\d*,\d+)\s+(\d+[\s,]*\d*,\d+)',
                'teams_rooms': r'CSP -MS Teams Rooms Pro \(Cycle\)\s+\d{6}\s+-\s+\d{6}\s+(\d+,\d+)\s+ST\s+(\d+[\s,]*\d*,\d+)\s+(\d+[\s,]*\d*,\d+)',
                'power_automate_plan': r'CSP -Power Automate with att RPA plan \(cycle\)\s+\d{6}\s+-\s+\d{6}\s+(\d+,\d+)\s+ST\s+(\d+[\s,]*\d*,\d+)\s+(\d+[\s,]*\d*,\d+)',
                'teams_eea': r'CSP -MS Teams EEA \(Cycle\)\s+\d{6}\s+-\s+\d{6}\s+(\d+,\d+)\s+ST\s+(\d+[\s,]*\d*,\d+)\s+(\d+[\s,]*\d*,\d+)',
                'copilot': r'CSP -MS Copilot for MS 365 \(Corr\)\s+\d{6}\s+-\s+\d{6}\s+(\d+,\d+)\s+ST\s+(\d+[\s,]*\d*,\d+)\s+(\d+[\s,]*\d*,\d+)',
                'ms365_eea': r'CSP -MS 365 E3 EEA \(no Teams\) \(Cycle\)\s+\d{6}\s+-\s+\d{6}\s+(\d+,\d+)\s+ST\s+(\d+[\s,]*\d*,\d+)\s+(\d+[\s,]*\d*,\d+)',
                'power_automate_prem': r'CSP -Power Automate prem\. \(Corr\)\s+\d{6}\s+-\s+\d{6}\s+(\d+,\d+)\s+ST\s+(\d+[\s,]*\d*,\d+)\s+(\d+[\s,]*\d*,\d+)'
            }
            
            # Extrahera information för varje licenstyp
            license_info = {}
            for license_type, pattern in license_patterns.items():
                match = re.search(pattern, text, re.IGNORECASE)
                if match:
                    quantity, unit_price, total = match.groups()
                    # Rensa och konvertera värden
                    quantity = float(quantity.replace(' ', '').replace(',', '.'))
                    unit_price = float(unit_price.replace(' ', '').replace(',', '.'))
                    total = float(total.replace(' ', '').replace(',', '.'))
                    
                    license_info[license_type] = {
                        'quantity': quantity,
                        'unit_price': unit_price,
                        'total': total
                    }
                    logger.info(f"Hittade licensinformation för {license_type}: {quantity} st à {unit_price} kr")
                else:
                    logger.warning(f"Kunde inte hitta information för {license_type}")
            
            # Validera att vi hittat all nödvändig information
            if not license_info:
                raise ValueError("Ingen licensinformation hittades i texten")
            
            # Spara parsed data som backup
            self.save_backup(
                {'license_info': license_info},
                f'parsed_licenses_{datetime.now().strftime("%Y%m%d_%H%M%S")}.json'
            )
            
            return license_info
            
        except Exception as e:
            logger.error(f"Fel vid parsning av licensinformation: {str(e)}")
            raise

    def generate_accounting_rows(self, license_info):
        """Genererar konteringsrader baserat på licensinformation."""
        try:
            logger.info("Börjar generera konteringsrader")
            accounting_rows = []
            
            # 1. Hantera Power BI Pro-licenser (förutom Mattias)
            if 'power_bi' in license_info:
                power_bi_info = license_info['power_bi']
                
                # Gruppera användare per RG
                non_automation_users = self.users_data[self.users_data['Specialhantering'] != 'Automation']
                rg_groups = non_automation_users.groupby('RG')
                
                for rg, group in rg_groups:
                    num_users = len(group)
                    if num_users > 0:
                        row = {
                            'Kon/Proj': '5420',
                            '': '',
                            'RG': rg,
                            'Aktivitet': '738',
                            'ProjKat': '',
                            ' ': '',
                            'Netto': round(num_users * power_bi_info['unit_price'], 2),
                            'Godkänt av': 'John Munthe'
                        }
                        accounting_rows.append(row)
                        logger.info(f"Lade till Power BI Pro-kontering för RG {rg}: {num_users} användare")
            
            # 2. Hantera Automation-projekt (20257601)
            automation_total = 0
            
            # Lägg till Mattias Power BI Pro-licens
            if 'power_bi' in license_info:
                automation_users = self.users_data[self.users_data['Specialhantering'] == 'Automation']
                num_automation_users = len(automation_users)
                if num_automation_users > 0:
                    automation_total += num_automation_users * license_info['power_bi']['unit_price']
                    logger.info(f"Lade till Power BI Pro-licens för {num_automation_users} automation-användare")
            
            # Lägg till övriga automation-licenser
            if 'power_automate_rpa' in license_info:
                automation_total += license_info['power_automate_rpa']['total']
            if 'power_automate_plan' in license_info:
                automation_total += license_info['power_automate_plan']['total']
            if 'power_automate_prem' in license_info:
                automation_total += license_info['power_automate_prem']['total']
            
            # Hämta automation-projektinställningar
            automation_settings = self.get_project_settings('20257601')
            automation_row = {
                'Kon/Proj': automation_settings['Kon/Proj'],
                '': '',
                'RG': '',
                'Aktivitet': automation_settings['Aktivitet'],
                'ProjKat': automation_settings['ProjKat'],
                ' ': '',
                'Netto': round(automation_total, 2),
                'Godkänt av': 'John Munthe'
            }
            accounting_rows.append(automation_row)
            logger.info(f"Lade till Automation-projektkontering: {automation_total} kr")
            
            # 3. Hantera Microsoft 365-projekt (20257407)
            ms365_total = 0
            
            if 'teams_eea' in license_info:
                ms365_total += license_info['teams_eea']['total']
            if 'copilot' in license_info:
                ms365_total += license_info['copilot']['total']
            if 'ms365_eea' in license_info:
                ms365_total += license_info['ms365_eea']['total']
            
            # Hämta MS365-projektinställningar
            ms365_settings = self.get_project_settings('20257407')
            ms365_row = {
                'Kon/Proj': ms365_settings['Kon/Proj'],
                '': '',
                'RG': '',
                'Aktivitet': ms365_settings['Aktivitet'],
                'ProjKat': ms365_settings['ProjKat'],
                ' ': '',
                'Netto': round(ms365_total, 2),
                'Godkänt av': 'John Munthe'
            }
            accounting_rows.append(ms365_row)
            logger.info(f"Lade till Microsoft 365-projektkontering: {ms365_total} kr")
            
            # 4. Hantera Teams Room-projekt (20257403)
            if 'teams_rooms' in license_info:
                # Hämta Teams-projektinställningar
                teams_settings = self.get_project_settings('20257403')
                teams_row = {
                    'Kon/Proj': teams_settings['Kon/Proj'],
                    '': '',
                    'RG': '',
                    'Aktivitet': teams_settings['Aktivitet'],
                    'ProjKat': teams_settings['ProjKat'],
                    ' ': '',
                    'Netto': round(license_info['teams_rooms']['total'], 2),
                    'Godkänt av': 'John Munthe'
                }
                accounting_rows.append(teams_row)
                logger.info(f"Lade till Teams Room-projektkontering: {license_info['teams_rooms']['total']} kr")
            
            return accounting_rows
            
        except Exception as e:
            logger.error(f"Fel vid generering av konteringsrader: {str(e)}")
            raise

    def save_to_excel(self, accounting_rows, output_file):
        """Sparar konteringsrader till Excel i Medius-format."""
        try:
            df = pd.DataFrame(accounting_rows, columns=[
                'Kon/Proj', '', 'RG', 'Aktivitet', 'ProjKat', '', 'Netto', 'Godkänt av'
            ])
            df.to_excel(output_file, index=False)
            logger.info(f"Konteringsrader sparade till {output_file}")
        except Exception as e:
            logger.error(f"Fel vid sparande av konteringsrader: {str(e)}")
            raise

    def _convert_to_serializable(self, obj):
        """Konverterar NumPy-typer till JSON-serialiserbara typer."""
        if isinstance(obj, dict):
            return {key: self._convert_to_serializable(value) for key, value in obj.items()}
        elif isinstance(obj, list):
            return [self._convert_to_serializable(item) for item in obj]
        elif hasattr(obj, 'item'):  # NumPy scalar
            return obj.item()
        return obj

    def save_backup(self, data, filename):
        """Sparar backup av data som JSON."""
        try:
            # Konvertera data till JSON-serialiserbart format
            serializable_data = self._convert_to_serializable(data)
            
            backup_path = os.path.join(self.output_dir, filename)
            with open(backup_path, 'w', encoding='utf-8') as f:
                json.dump(serializable_data, f, ensure_ascii=False, indent=2)
            logger.info(f"Backup sparad till {backup_path}")
        except Exception as e:
            logger.error(f"Fel vid sparande av backup: {str(e)}")
            raise

    def validate_accounting_rows(self, accounting_rows, license_info):
        """Validerar konteringsrader mot licensinformation med detaljerade kontroller."""
        try:
            logger.info("\n=== VALIDERING AV KONTERINGSRADER ===")
            
            # 1. Summera Power BI Pro-konteringar (RG-baserade)
            power_bi_total = sum(row['Netto'] for row in accounting_rows if row['Kon/Proj'] == '5420')
            logger.info(f"\nPower BI Pro-konteringar per RG:")
            for row in accounting_rows:
                if row['Kon/Proj'] == '5420':
                    logger.info(f"RG {row['RG']}: {row['Netto']} kr")
            logger.info(f"Totalt Power BI Pro (RG): {power_bi_total} kr")
            
            # 2. Validera automationslicenser
            automation_row = next(row for row in accounting_rows if row['Kon/Proj'] == 'P.20257601')
            automation_total = automation_row['Netto']
            
            logger.info(f"\nAutomationslicenser (P.20257601):")
            if 'power_automate_rpa' in license_info:
                logger.info(f"Power Automate RPA: {license_info['power_automate_rpa']['total']} kr")
            if 'power_automate_plan' in license_info:
                logger.info(f"Power Automate Plan: {license_info['power_automate_plan']['total']} kr")
            if 'power_automate_prem' in license_info:
                logger.info(f"Power Automate Prem: {license_info['power_automate_prem']['total']} kr")
            
            automation_users = self.users_data[self.users_data['Specialhantering'] == 'Automation']
            if len(automation_users) > 0 and 'power_bi' in license_info:
                logger.info(f"Power BI Pro (Automation): {len(automation_users) * license_info['power_bi']['unit_price']} kr")
            
            logger.info(f"Totalt Automation: {automation_total} kr")
            
            # 3. Validera Microsoft 365-licenser
            ms365_row = next(row for row in accounting_rows if row['Kon/Proj'] == 'P.20257407')
            ms365_total = ms365_row['Netto']
            
            logger.info(f"\nMicrosoft 365-licenser (P.20257407):")
            if 'teams_eea' in license_info:
                logger.info(f"Teams EEA: {license_info['teams_eea']['total']} kr")
            if 'copilot' in license_info:
                logger.info(f"Copilot: {license_info['copilot']['total']} kr")
            if 'ms365_eea' in license_info:
                logger.info(f"MS365 EEA: {license_info['ms365_eea']['total']} kr")
            logger.info(f"Totalt Microsoft 365: {ms365_total} kr")
            
            # 4. Validera Teams Room-licenser
            teams_row = next(row for row in accounting_rows if row['Kon/Proj'] == 'P.20257403')
            teams_total = teams_row['Netto']
            logger.info(f"\nTeams Room-licenser (P.20257403):")
            logger.info(f"Totalt Teams Room: {teams_total} kr")
            
            # 5. Validera totalsumma
            total_sum = sum(row['Netto'] for row in accounting_rows)
            expected_sum = sum(info['total'] for info in license_info.values())
            
            logger.info(f"\n=== SUMMERING ===")
            logger.info(f"Totalsumma från konteringsrader: {total_sum} kr")
            logger.info(f"Förväntad summa från PDF: {expected_sum} kr")
            
            if abs(total_sum - expected_sum) <= 0.02:
                logger.info("[OK] Totalsumman stämmer med fakturan")
            else:
                logger.warning(f"[!] Differens i totalsumma: {total_sum - expected_sum} kr")
            
            # 6. Validera delsummor
            logger.info("\n=== VALIDERING AV DELSUMMOR ===")
            
            # Beräkna förväntad automationssumma
            expected_automation = (
                (license_info.get('power_automate_rpa', {}).get('total', 0)) +
                (license_info.get('power_automate_plan', {}).get('total', 0)) +
                (license_info.get('power_automate_prem', {}).get('total', 0))
            )
            if len(automation_users) > 0 and 'power_bi' in license_info:
                expected_automation += len(automation_users) * license_info['power_bi']['unit_price']
            
            if abs(automation_total - expected_automation) <= 0.02:
                logger.info("[OK] Automationssumman stämmer")
            else:
                logger.warning(f"[!] Differens i automationssumma: {automation_total - expected_automation} kr")
            
            # Beräkna förväntad MS365-summa
            expected_ms365 = (
                (license_info.get('teams_eea', {}).get('total', 0)) +
                (license_info.get('copilot', {}).get('total', 0)) +
                (license_info.get('ms365_eea', {}).get('total', 0))
            )
            
            if abs(ms365_total - expected_ms365) <= 0.02:
                logger.info("[OK] Microsoft 365-summan stämmer")
            else:
                logger.warning(f"[!] Differens i Microsoft 365-summa: {ms365_total - expected_ms365} kr")
            
            # Validera Teams Room
            expected_teams = license_info.get('teams_rooms', {}).get('total', 0)
            if abs(teams_total - expected_teams) <= 0.02:
                logger.info("[OK] Teams Room-summan stämmer")
            else:
                logger.warning(f"[!] Differens i Teams Room-summa: {teams_total - expected_teams} kr")
            
            logger.info("\n=== VALIDERING SLUTFÖRD ===")
            
        except Exception as e:
            logger.error(f"Fel vid validering av konteringsrader: {str(e)}")
            raise

    def generate_invoice_comment(self, license_info, accounting_rows):
        """Genererar kommentar för fakturan baserat på licensinformation."""
        comment_parts = []
        
        # 1. Power BI Pro-användare
        logger.info("Genererar Power BI Pro-användarlista")
        try:
            # Filtrera användare som inte är markerade för specialhantering
            pbi_users = self.users_data[
                (self.users_data['Specialhantering'].isna()) | 
                (self.users_data['Specialhantering'] != 'Automation')
            ].sort_values(['RG', 'Namn'])
            
            if not pbi_users.empty:
                comment_parts.append("\nPower BI Pro-licenser")
                for _, user in pbi_users.iterrows():
                    comment_parts.append(f"{user['Namn']}\t{user['RG']}:{user['Kostnadsställe']}")
        except Exception as e:
            logger.error(f"Fel vid generering av Power BI Pro-användarlista: {str(e)}")
            raise
        
        # 2. Gruppera övriga licenser per mottagare
        license_by_receiver = {}
        
        # Automation/Digital Utveckling
        automation_licenses = []
        if 'power_automate_rpa' in license_info:
            automation_licenses.append("Power Automate unattended RPA add-on")
        if 'power_automate_plan' in license_info:
            automation_licenses.append("Power Automate with att RPA plan")
        if 'power_automate_prem' in license_info:
            automation_licenses.append("Power Automate prem")
        
        automation_users = self.users_data[self.users_data['Specialhantering'] == 'Automation']
        if not automation_users.empty:
            automation_licenses.append("Power BI Pro (Mattias)")
        
        if automation_licenses:
            # Hämta mottagare från Project Settings
            automation_project = self.project_settings[self.project_settings['ProjektID'] == '20257601']
            receiver = 'Digital Utveckling och integration'  # Standard
            if not automation_project.empty and 'Mottagare' in automation_project.columns:
                receiver = automation_project.iloc[0]['Mottagare']
            license_by_receiver[receiver] = automation_licenses
        
        # Microsoft 365/Digital Arbetsplats
        ms365_licenses = []
        if 'teams_eea' in license_info:
            ms365_licenses.append("MS Teams EEA")
        if 'copilot' in license_info:
            ms365_licenses.append("MS Copilot for MS 365")
        if 'ms365_eea' in license_info:
            ms365_licenses.append("MS 365 E3 EEA (no Teams)")
        
        if ms365_licenses:
            # Hämta mottagare från Project Settings
            ms365_project = self.project_settings[self.project_settings['ProjektID'] == '20257407']
            receiver = 'Digital Arbetsplats'  # Standard
            if not ms365_project.empty and 'Mottagare' in ms365_project.columns:
                receiver = ms365_project.iloc[0]['Mottagare']
            license_by_receiver[receiver] = ms365_licenses
        
        # Teams Room/Digital Arbetsplats
        if 'teams_rooms' in license_info:
            # Hämta mottagare från Project Settings
            teams_project = self.project_settings[self.project_settings['ProjektID'] == '20257403']
            receiver = 'Digital Arbetsplats'  # Standard
            if not teams_project.empty and 'Mottagare' in teams_project.columns:
                receiver = teams_project.iloc[0]['Mottagare']
            
            if receiver in license_by_receiver:
                license_by_receiver[receiver].append("MS Teams Rooms Pro")
            else:
                license_by_receiver[receiver] = ["MS Teams Rooms Pro"]
        
        # Lägg till grupperade licenser i kommentaren
        for receiver, licenses in license_by_receiver.items():
            comment_parts.append(f"\nTill {receiver} licenser för {', '.join(licenses)}")
        
        return "\n".join(comment_parts)

    def process_invoice(self, pdf_path):
        """Huvudfunktion för att processa en faktura."""
        try:
            logger.info(f"Börjar processa faktura: {pdf_path}")
            
            # Extrahera text från PDF
            ocr_text = self.extract_text_from_pdf(pdf_path)
            
            # Spara backup av OCR-text
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            self.save_backup(ocr_text, f"ocr_backup_{timestamp}.json")
            
            # Parsa licensinformation
            logger.info("Börjar parsning av licensinformation")
            license_info = self.parse_license_info(ocr_text)
            
            # Spara backup av parsad licensinformation
            self.save_backup(license_info, f"parsed_licenses_{timestamp}.json")
            
            # Generera konteringsrader
            logger.info("Börjar generera konteringsrader")
            accounting_rows = self.generate_accounting_rows(license_info)
            
            # Spara backup av konteringsrader
            self.save_backup(accounting_rows, f"accounting_rows_{timestamp}.json")
            
            # Validera konteringsrader
            self.validate_accounting_rows(accounting_rows, license_info)
            
            # Generera och skriv ut kommentar
            comment = self.generate_invoice_comment(license_info, accounting_rows)
            logger.info("\n=== KOMMENTAR FÖR MEDIUS ===\n")
            logger.info(comment)
            logger.info("\n===========================")
            
            # Spara konteringsrader till Excel
            output_file = f"output/kontering_{timestamp}.xlsx"
            self.save_to_excel(accounting_rows, output_file)
            logger.info(f"Konteringsrader sparade till {output_file}")
            
            logger.info("Faktura processad framgångsrikt")
            return output_file
            
        except Exception as e:
            logger.error(f"Fel vid processning av faktura: {str(e)}")
            raise

def main():
    try:
        # Skapa en instans av InvoiceHelper
        helper = InvoiceHelper()
        
        # Använd tkinter för att välja PDF-fil
        import tkinter as tk
        from tkinter import filedialog
        
        root = tk.Tk()
        root.withdraw()  # Dölj huvudfönstret
        
        print("\nVälj PDF-fil med Atea-faktura...")
        pdf_path = filedialog.askopenfilename(
            title="Välj Atea-faktura (PDF)",
            filetypes=[("PDF-filer", "*.pdf")],
            initialdir="C:/Users/Public/downloads"  # Ändrat till Public downloads-mappen
        )
        
        if not pdf_path:
            print("Ingen fil valdes. Avslutar.")
            return
        
        print(f"\nVald fil: {pdf_path}")
        
        # Processa fakturan
        output_file = helper.process_invoice(pdf_path)
        print(f"\nKontering genererad och sparad till: {output_file}")
        
        # Öppna utdatamappen
        os.startfile(os.path.dirname(output_file))
        
    except Exception as e:
        logger.error(f"Programmet avbröts på grund av ett fel: {str(e)}")
        raise

if __name__ == '__main__':
    main() 