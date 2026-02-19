import streamlit as st
import gspread
import pandas as pd
from fpdf import FPDF
import datetime
import time

# 1. Configuration de la page
st.set_page_config(page_title="Calculateur Liasses fiscales 2033", layout="wide")

# 2. Extraction de l'ID Master depuis tes Secrets
url_master = st.secrets["connections"]["gsheets"]["spreadsheet"]
MASTER_ID = url_master.split("/d/")[1].split("/")[0]

st.title("🏠 Calculateur des liasses fiscales 2033")
st.write("Mode sécurisé : Une copie temporaire est créée pour chaque calcul.")

# 3. Formulaire de saisie
tab1, tab2, tab3 = st.tabs(["🏗️ Acquisition & Travaux", "💰 Exploitation (Annuelle)", "📉 Antériorité & Report"])

with tab1:
    col1, col2 = st.columns(2)
    with col1:
        prix_achat = st.number_input("Prix d'achat (€)", min_value=0, value=200000)
        frais_notaire = st.number_input("Frais de notaire (€)", min_value=0, value=15000)
        frais_agence = st.number_input("Frais d'agence (€)", min_value=0, value=0)
    with col2:
        mobilier = st.number_input("Mobilier (€)", min_value=0, value=5000)
        travaux = st.number_input("Travaux à amortir (€)", min_value=0, value=0)
        debut_activite = st.date_input("Date de début d'activité", value=datetime.date.today())

with tab2:
    col_a, col_b = st.columns(2)
    with col_a:
        loyer = st.number_input("Loyer annuel CC (€)", min_value=0, value=12000)
        charges = st.number_input("Charges de copropriété (€/an)", min_value=0, value=1200)
        taxe_fonciere = st.number_input("Taxe foncière (€/an)", min_value=0, value=800)
    with col_b:
        interets = st.number_input("Intérêts d'emprunt (€/an)", min_value=0, value=3000)

with tab3:
    amort_excedentaires = st.number_input("Amortissements excédentaires ANTÉRIEURS (€)", min_value=0, value=0)
    deficits_anterieurs = st.number_input("Déficits ANTÉRIEURS (€)", min_value=0, value=0)
    dispo_anterieurs = st.number_input("Disponibilité Antérieurs (€)", min_value=0, value=0)

# 4. Fonction PDF robuste
def create_pdf(p1, p2, p3, p4, p5):
    pdf = FPDF(orientation='L', unit='mm', format='A4')
    pdf.set_auto_page_break(auto=True, margin=15)
    def add_page_data(titre, data):
        pdf.add_page()
        pdf.set_font("helvetica", 'B', 14); pdf.cell(0, 10, titre, ln=True, align='C'); pdf.ln(5)
        if not data: return
        pdf.set_font("helvetica", size=8)
        cw = 270 / len(data[0])
        for row in data:
            for item in row: pdf.cell(cw, 7, str(item)[:20], border=1)
            pdf.ln(7)
    add_page_data("BILAN", p1); add_page_data("RESULTAT", p2); add_page_data("IMMO", p3); add_page_data("AMORT", p4); add_tab("DEFICITS", p5)
    return pdf.output()

# 5. Logique principale
if st.button("Enregistrer et Calculer"):
    temp_sheet_id = None
    try:
        creds = st.secrets["connections"]["gsheets"]
        gc = gspread.service_account_from_dict(creds)
        
        with st.spinner("Création de l'espace isolé..."):
            timestamp = int(time.time())
            new_copy = gc.copy(MASTER_ID, title=f"TEMP_{timestamp}")
            temp_sheet_id = new_copy.id
            sh = gc.open_by_key(temp_sheet_id)
            worksheet = sh.worksheet("test python")

            date_str = debut_activite.strftime("%d/%m/%Y")
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

        with st.spinner("Extraction..."):
            ws_res = sh.get_worksheet(1)
            p1, p2, p3, p4, p5 = ws_res.get('A109:G124'), ws_res.get('A129:E144'), ws_res.get('A151:I158'), ws_res.get('A162:I169'), ws_res.get('A178:C182')
            
            st.success("✅ Calcul terminé !")
            st.dataframe(pd.DataFrame(p2), use_container_width=True)

            pdf_out = create_pdf(p1, p2, p3, p4, p5)
            st.download_button("📥 Télécharger PDF", data=pdf_out, file_name=f"liasse_{timestamp}.pdf", mime="application/pdf")

    except Exception as e:
        st.error(f"Erreur : {e}")
    finally:
        if temp_sheet_id:
            try: gc.del_spreadsheet(temp_sheet_id)
            except: pass

