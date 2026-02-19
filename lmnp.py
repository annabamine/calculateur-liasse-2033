import streamlit as st
from streamlit_gsheets import GSheetsConnection
import gspread 
import pandas as pd
from fpdf import FPDF

# 1. Configuration de la page
st.set_page_config(page_title="Calculateur Liasses fiscales 2033", layout="wide")

# 2. Création de la connexion
conn = st.connection("gsheets", type=GSheetsConnection)

st.title("🏠 Calculateur des liasses fiscales 2033")
st.write("Remplissez les champs ci-dessous pour calculer vos liasses fiscales 2033.")

# Organisation par onglets
tab1, tab2, tab3 = st.tabs(["🏗️ Acquisition & Travaux", "💰 Exploitation (Annuelle)", "📉 Antériorité & Report"])

with tab1:
    col1, col2 = st.columns(2)
    with col1:
        prix_achat = st.number_input("Prix d'achat (€)", min_value=0, value=200000, step=1000)
        frais_notaire = st.number_input("Frais de notaire (€)", min_value=0, value=15000)
        frais_agence = st.number_input("Frais d'agence (€)", min_value=0, value=0)
    with col2:
        mobilier = st.number_input("Mobilier (€)", min_value=0, value=5000)
        travaux = st.number_input("Travaux à amortir (€)", min_value=0, value=0)
        debut_activite = st.date_input("Date de début d'activité", value=None)

with tab2:
    col_a, col_b = st.columns(2)
    with col_a:
        loyer = st.number_input("Loyer annuel CC (€)", min_value=0, value=12000)
        charges = st.number_input("Charges de copropriété (€/an)", min_value=0, value=1200)
        taxe_fonciere = st.number_input("Taxe foncière (€/an)", min_value=0, value=800)
    with col_b:
        interets = st.number_input("Intérêts d'emprunt (€/an)", min_value=0, value=3000)

with tab3:
    st.info("Ces champs concernent vos reports des années précédentes.")
    amort_excedentaires = st.number_input("Amortissements excédentaires ANTÉRIEURS (€)", min_value=0, value=0)
    deficits_anterieurs = st.number_input("Déficits ANTÉRIEURS (€)", min_value=0, value=0)
    dispo_anterieurs = st.number_input("Disponibilité Antérieurs (€)", min_value=0, value=0)

def create_pdf(p1, p2, p3, p4, p5):
    pdf = FPDF(orientation='L', unit='mm', format='A4') # Mode Paysage pour plus de place
    pdf.set_auto_page_break(auto=True, margin=15)
    
    def ajouter_page_tableau(titre, data):
        pdf.add_page()
        pdf.set_font("helvetica", 'B', 14)
        pdf.cell(0, 10, txt=titre, ln=True, align='C')
        pdf.ln(5)
        
        if not data:
            return

        # Calcul automatique de la largeur des colonnes
        # On divise les 270mm utilisables par le nombre de colonnes trouvées
        nb_cols = len(data[0])
        col_width = 270 / nb_cols
        
        pdf.set_font("helvetica", size=7) # Police petite pour que tout rentre
        
        for row in data:
            # On récupère la position Y de départ pour la ligne
            start_y = pdf.get_y()
            max_row_height = 0
            
            # 1. On calcule d'abord la hauteur nécessaire pour cette ligne
            # (pour que les bordures soient jolies)
            row_heights = []
            for item in row:
                # On simule le texte pour voir combien de lignes il prend
                text = str(item).replace('"', '') if item is not None else ""
                lines = pdf.multi_cell(col_width, 5, txt=text, split_only=True)
                row_heights.append(len(lines) * 5)
            
            line_h = max(row_heights) if row_heights else 5
            
            # 2. On dessine les cellules une par une
            for item in row:
                curr_x = pdf.get_x()
                text = str(item).replace('"', '') if item is not None else ""
                
                # On dessine le rectangle de fond pour la structure
                pdf.rect(curr_x, start_y, col_width, line_h)
                
                # On écrit le texte
                pdf.multi_cell(col_width, 5, txt=text, align='L')
                
                # On se replace pour la colonne suivante
                pdf.set_xy(curr_x + col_width, start_y)
            
            # On passe à la ligne suivante
            pdf.ln(line_h)

    # Appels simplifiés (le code calcule les largeurs tout seul)
    ajouter_page_tableau("PAGE 1 : BILAN", p1)
    ajouter_page_tableau("PAGE 2 : COMPTE DE RÉSULTAT", p2, largeurs_manuelles=[90, 20, 40, 40, 40, 40])
    ajouter_page_tableau("PAGE 3 : IMMOBILISATIONS", p3)
    ajouter_page_tableau("PAGE 3 : AMORTISSEMENTS", p4)
    ajouter_page_tableau("PAGE 4 : SUIVI DES DÉFICITS", p5)
    
    return pdf.output()

    # Pages du PDF
    ajouter_page_tableau("PAGE 1 : BILAN", p1, [90, 20, 40, 40])
    ajouter_page_tableau("PAGE 2 : COMPTE DE RÉSULTAT", p2, [150, 20, 33, 33, 34])
    ajouter_page_tableau("PAGE 3 : IMMOBILISATIONS", p3, [70, 30, 30, 30, 30])
    ajouter_page_tableau("PAGE 3 : AMORTISSEMENTS", p4, [70, 30, 30, 30, 30])
    ajouter_page_tableau("PAGE 4 : DÉFICITS", p5, [130, 60])
    
    return pdf.output()

if st.button("Enregistrer les données"):
    try:
        # 1. Récupération des crédentials
        creds = st.secrets["connections"]["gsheets"]
        
        # 2. Connexion Service Account
        gc = gspread.service_account_from_dict(creds)
        
        # 3. Ouverture du fichier Master
        sh = gc.open_by_url(creds["spreadsheet"])
        
        # Formatage de la date
        date_str = debut_activite.strftime("%d/%m/%Y") if debut_activite else ""

        # 4. Écriture directe dans l'onglet Master
        worksheet = sh.worksheet("test python")
        worksheet.update_acell('B4', prix_achat)
        worksheet.update_acell('B5', frais_notaire)
        worksheet.update_acell('B6', frais_agence)
        worksheet.update_acell('B8', travaux)
        worksheet.update_acell('B7', mobilier)
        worksheet.update_acell('B12', date_str)

        worksheet.update_acell('B84', loyer)
        worksheet.update_acell('B85', charges)
        worksheet.update_acell('B86', taxe_fonciere)
        worksheet.update_acell('B87', interets)

        worksheet.update_acell('B93', amort_excedentaires)
        worksheet.update_acell('B98', deficits_anterieurs)
        worksheet.update_acell('B103', dispo_anterieurs)
        
        st.success(f"✅ Liasse mise à jour !")
        st.divider() 

        # 5. Récupération des résultats
        with st.spinner("Extraction des résultats..."):
            ws_res = sh.worksheet("test python")
            
            def afficher_tableau_pro(data, titre):
                if not data:
                    return st.warning(f"Aucune donnée trouvée pour {titre}")
                df = pd.DataFrame(data)
                df = df.fillna("")
                df = df.mask(df == "").dropna(how='all').dropna(axis=1, how='all').fillna("")
                st.subheader(titre)
                return st.dataframe(df, use_container_width=True, hide_index=True)

            tab_p1, tab_p2, tab_p3, tab_p4 = st.tabs([
                "📄 Page 1 : Bilan", 
                "📄 Page 2 : Résultat", 
                "📄 Page 3 : Immo & Amort.", 
                "📄 Page 4 : Déficits"
            ])

            with tab_p1:
                p1_data = ws_res.get('A109:G124')
                afficher_tableau_pro(p1_data, "État de l'Actif / Passif")

            with tab_p2:
                p2_data = ws_res.get('A129:E144')
                if p2_data:
                    res_fiscal = p2_data[-1][-3]
                    st.metric("RÉSULTAT FISCAL FINAL", f"{res_fiscal} €")
                afficher_tableau_pro(p2_data, "Détail du Résultat Fiscal")

            with tab_p3:
                p3_data = ws_res.get('A151:I158')
                afficher_tableau_pro(p3_data, "Tableau des Immobilisations")
                p4_data = ws_res.get('A162:I169')
                afficher_tableau_pro(p4_data, "Tableau des Amortissements")

            with tab_p4:
                p5_data = ws_res.get('A178:C182')
                afficher_tableau_pro(p5_data, "Suivi des Déficits")

            # 6. Génération et téléchargement du PDF
            pdf_raw = create_pdf(p1_data, p2_data, p3_data, p4_data, p5_data)
            pdf_final = bytes(pdf_raw) 
            
            st.download_button(
                label="📥 Télécharger la Liasse en PDF",
                data=pdf_final,
                file_name=f"liasse_fiscale_{date_str.replace('/', '-')}.pdf",
                mime="application/pdf",
                use_container_width=True
            )

    except Exception as e:
        st.error(f"Erreur de liaison : {e}")