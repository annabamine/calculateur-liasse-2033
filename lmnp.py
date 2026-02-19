import streamlit as st
from streamlit_gsheets import GSheetsConnection
import gspread 
import pandas as pd
from fpdf import FPDF
import datetime

# 1. Configuration de la page
st.set_page_config(page_title="Calculateur Liasses fiscales 2033", layout="wide")

# 2. Création de la connexion
conn = st.connection("gsheets", type=GSheetsConnection)

# --- FONCTION DE RÉINITIALISATION ---
def reset_form():
    for key in st.session_state.keys():
        del st.session_state[key]
    st.rerun()

st.title("🏠 Calculateur des liasses fiscales 2033")
st.write("Remplissez les champs ci-dessous. Les champs seront vidés après le téléchargement.")

# Organisation par onglets
tab1, tab2, tab3 = st.tabs(["🏗️ Acquisition & Travaux", "💰 Exploitation (Annuelle)", "📉 Antériorité & Report"])

with tab1:
    col1, col2 = st.columns(2)
    with col1:
        prix_achat = st.number_input("Prix d'achat (€)", min_value=0, value=200000, step=1000, key="prix_achat")
        frais_notaire = st.number_input("Frais de notaire (€)", min_value=0, value=15000, key="frais_notaire")
        frais_agence = st.number_input("Frais d'agence (€)", min_value=0, value=0, key="frais_agence")
    with col2:
        mobilier = st.number_input("Mobilier (€)", min_value=0, value=5000, key="mobilier")
        travaux = st.number_input("Travaux à amortir (€)", min_value=0, value=0, key="travaux")
        debut_activite = st.date_input("Date de début d'activité", value=datetime.date.today(), key="debut_activite")

with tab2:
    col_a, col_b = st.columns(2)
    with col_a:
        loyer = st.number_input("Loyer annuel CC (€)", min_value=0, value=12000, key="loyer")
        charges = st.number_input("Charges de copropriété (€/an)", min_value=0, value=1200, key="charges")
        taxe_fonciere = st.number_input("Taxe foncière (€/an)", min_value=0, value=800, key="taxe_fonciere")
    with col_b:
        interets = st.number_input("Intérêts d'emprunt (€/an)", min_value=0, value=3000, key="interets")

with tab3:
    st.info("Ces champs concernent vos reports des années précédentes.")
    amort_excedentaires = st.number_input("Amortissements excédentaires ANTÉRIEURS (€)", min_value=0, value=0, key="amort_ant")
    deficits_anterieurs = st.number_input("Déficits ANTÉRIEURS (€)", min_value=0, value=0, key="deficits_ant")
    dispo_anterieurs = st.number_input("Disponibilité Antérieurs (€)", min_value=0, value=0, key="dispo_ant")

# --- FONCTION PDF (fpdf2 compatible) ---
def create_pdf(p1, p2, p3, p4, p5):
    pdf = FPDF(orientation='L', unit='mm', format='A4')
    pdf.set_auto_page_break(auto=True, margin=15)
    
    def ajouter_page_tableau(titre, data, largeurs_manuelles=None):
        pdf.add_page()
        pdf.set_font("helvetica", 'B', 14)
        pdf.cell(0, 10, txt=titre, ln=True, align='C')
        pdf.ln(5)
        if not data: return
        nb_cols = len(data[0])
        col_widths = largeurs_manuelles if largeurs_manuelles else [270 / nb_cols] * nb_cols
        pdf.set_font("helvetica", size=7)
        for row in data:
            start_y = pdf.get_y()
            clean_row = [str(item).replace('"', '') if item is not None else "" for item in row]
            h_list = [len(pdf.multi_cell(col_widths[i], 5, txt=clean_row[i], split_only=True)) * 5 for i in range(min(len(clean_row), len(col_widths)))]
            line_h = max(h_list) if h_list else 5
            curr_x = pdf.get_x()
            for i in range(min(len(clean_row), len(col_widths))):
                pdf.rect(curr_x, start_y, col_widths[i], line_h)
                pdf.multi_cell(col_widths[i], 5, txt=clean_row[i], align='L')
                curr_x += col_widths[i]
                pdf.set_xy(curr_x, start_y)
            pdf.ln(line_h)

    ajouter_page_tableau("PAGE 1 : BILAN", p1)
    ajouter_page_tableau("PAGE 2 : COMPTE DE RÉSULTAT", p2, [90, 20, 40, 40, 40, 40])
    ajouter_page_tableau("PAGE 3 : IMMOBILISATIONS", p3)
    ajouter_page_tableau("PAGE 3 : AMORTISSEMENTS", p4)
    ajouter_page_tableau("PAGE 4 : SUIVI DES DÉFICITS", p5)
    return pdf.output()

# --- LOGIQUE PRINCIPALE ---
if st.button("Enregistrer et Calculer"):
    try:
        creds = st.secrets["connections"]["gsheets"]
        gc = gspread.service_account_from_dict(creds)
        sh = gc.open_by_url(creds["spreadsheet"])
        worksheet = sh.worksheet("test python")
        
        date_str = debut_activite.strftime("%d/%m/%Y") if debut_activite else ""

        # Mise à jour du Master
        updates = [
            {'range': 'B4', 'values': [[prix_achat]]}, {'range': 'B5', 'values': [[frais_notaire]]},
            {'range': 'B6', 'values': [[frais_agence]]}, {'range': 'B8', 'values': [[travaux]]},
            {'range': 'B7', 'values': [[mobilier]]}, {'range': 'B12', 'values': [[date_str]]},
            {'range': 'B84', 'values': [[loyer]]}, {'range': 'B85', 'values': [[charges]]},
            {'range': 'B86', 'values': [[taxe_fonciere]]}, {'range': 'B87', 'values': [[interets]]},
            {'range': 'B93', 'values': [[amort_excedentaires]]}, {'range': 'B98', 'values': [[deficits_anterieurs]]},
            {'range': 'B103', 'values': [[dispo_anterieurs]]}
        ]
        worksheet.batch_update(updates)
        
        st.success("✅ Données envoyées au Master !")

        with st.spinner("Calcul en cours..."):
            ws_res = sh.worksheet("test python")
            p1 = ws_res.get('A109:G124')
            p2 = ws_res.get('A129:E144')
            p3 = ws_res.get('A151:I158')
            p4 = ws_res.get('A162:I169')
            p5 = ws_res.get('A178:C182')

            st.dataframe(pd.DataFrame(p2), use_container_width=True)

            pdf_raw = create_pdf(p1, p2, p3, p4, p5)
            
            # Bouton de téléchargement
            st.download_button(
                label="📥 Télécharger la Liasse (Ceci réinitialisera le formulaire)",
                data=bytes(pdf_raw),
                file_name=f"liasse_{date_str.replace('/', '-')}.pdf",
                mime="application/pdf",
                on_click=reset_form, # Déclenche la remise à zéro
                use_container_width=True
            )

    except Exception as e:
        st.error(f"Erreur : {e}")