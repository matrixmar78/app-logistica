import streamlit as st
import pandas as pd
from io import BytesIO
from fpdf import FPDF
import streamlit.components.v1 as components

# --- REGOLE CSS PER LA STAMPA PULITA ---
st.markdown("""
    <style>
    @media print {
        [data-testid="stSidebar"] { display: none !important; }
        header { display: none !important; }
        [data-testid="stForm"] { display: none !important; }
        .block-container { padding-top: 1rem !important; max-width: 100% !important; }
    }
    </style>
""", unsafe_allow_html=True)

# --- 1. INIZIALIZZAZIONE DATI ---
if 'materiali' not in st.session_state:
    st.session_state.materiali = [
        {"nome": "Cavi di rete 15m", "quantita": 2, "procurato": False, "posizione": ""},
        {"nome": "Switch 8 porte", "quantita": 1, "procurato": True, "posizione": "Ufficio"}
    ]
if 'riga_in_modifica' not in st.session_state:
    st.session_state.riga_in_modifica = None

st.title("📦 Logistica Materiali")
st.write("Gestisci l'attrezzatura per il tuo lavoro.")

# --- 2. BARRA LATERALE: IMPORTA, ESPORTA E STAMPA ---
st.sidebar.title("⚙️ Azioni")

# 2A. IMPORTA FILE
st.sidebar.write("**📤 Importa Dati**")
file_caricato = st.sidebar.file_uploader("Carica Excel o CSV", type=["xlsx", "csv"])

if file_caricato is not None:
    if st.sidebar.button("Aggiungi alla lista"):
        try:
            if file_caricato.name.endswith('.csv'):
                df_in = pd.read_csv(file_caricato)
            else:
                df_in = pd.read_excel(file_caricato)
            
            conteggio = 0
            for index, row in df_in.iterrows():
                nome_mat = ""
                if 'Nome Materiale' in df_in.columns:
                    nome_mat = str(row['Nome Materiale'])
                elif len(df_in.columns) > 0:
                    nome_mat = str(row.iloc[0])
                
                if nome_mat and nome_mat.lower() != 'nan':
                    # Recupero Quantità
                    qta_val = 1
                    if 'Quantità' in df_in.columns:
                        try:
                            qta_val = int(row['Quantità'])
                        except:
                            qta_val = 1

                    # Recupero Stato
                    procurato_val = False
                    if 'Procurato?' in df_in.columns:
                        val_proc = str(row['Procurato?']).lower().strip()
                        if val_proc in ['sì', 'si', 'true', '1', 'yes']:
                            procurato_val = True
                    
                    # Recupero Posizione
                    posizione_val = ""
                    if 'Posizione' in df_in.columns:
                        val_pos = str(row['Posizione']).strip()
                        if val_pos in ["Ufficio", "Pedana", "Furgone"]:
                            posizione_val = val_pos
                    
                    st.session_state.materiali.append({
                        "nome": nome_mat,
                        "quantita": qta_val,
                        "procurato": procurato_val,
                        "posizione": posizione_val
                    })
                    conteggio += 1
            
            st.sidebar.success(f"✅ {conteggio} materiali importati!")
        except Exception as e:
            st.sidebar.error(f"Errore durante la lettura: {e}")

st.sidebar.divider()

# 2B. ESPORTAZIONE (Excel e PDF)
st.sidebar.write("**📥 Esporta Dati**")
df = pd.DataFrame(st.session_state.materiali)

if not df.empty:
    df_export = df.copy()
    df_export['procurato'] = df_export['procurato'].map({True: 'Sì', False: 'No'})
    # Riordiniamo le colonne per inserire la quantità
    df_export = df_export[['nome', 'quantita', 'procurato', 'posizione']]
    df_export.columns = ['Nome Materiale', 'Quantità', 'Procurato?', 'Posizione']
    
    # EXCEL
    output_excel = BytesIO()
    with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
        df_export.to_excel(writer, index=False, sheet_name='Materiali')
    excel_data = output_excel.getvalue()
    
    st.sidebar.download_button(
        label="📊 Esporta in Excel",
        data=excel_data,
        file_name="Lista_Materiali.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    
    # PDF
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("helvetica", "B", 16)
    pdf.cell(0, 10, "Distinta Materiali di Lavoro", align="C", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(5)
    
    pdf.set_font("helvetica", size=12)
    for item in st.session_state.materiali:
        stato = "Procurato" if item['procurato'] else "Da trovare"
        pos = item['posizione'] if item['posizione'] else "---"
        qta = item.get('quantita', 1)
        riga = f"- {qta}x {item['nome']}  |  Stato: {stato}  |  Posizione: {pos}"
        pdf.cell(0, 10, riga, new_x="LMARGIN", new_y="NEXT")
        
    pdf_bytes = bytes(pdf.output()) 
    
    st.sidebar.download_button(
        label="📄 Esporta in PDF",
        data=pdf_bytes,
        file_name="Lista_Materiali.pdf",
        mime="application/pdf"
    )

st.sidebar.divider()

# 2C. STAMPA RAPIDA
st.sidebar.write("**🖨️ Stampa**")
components.html(
    """
    <button onclick="window.parent.print()" 
    style="width:100%; padding: 8px; border-radius: 6px; border: 1px solid #dcdcdc; 
    background-color: white; color: #333; cursor: pointer; font-family: sans-serif; font-size: 14px;">
        Stampa questa pagina
    </button>
    """,
    height=50
)

# --- 3. INSERIMENTO MANUALE MATERIALI ---
st.subheader("Aggiungi nuovo materiale")
with st.form("form_inserimento", clear_on_submit=True):
    # Aggiunta la colonna per la quantità nel form
    col_input, col_qta, col_btn = st.columns([3, 1, 1])
    with col_input:
        nuovo_nome = st.text_input("Cosa devi portare?", label_visibility="collapsed", placeholder="Es. Trapano a batteria...")
    with col_qta:
        nuova_qta = st.number_input("Q.tà", min_value=1, value=1, label_visibility="collapsed")
    with col_btn:
        inviato = st.form_submit_button("Aggiungi ➕")
    
    if inviato and nuovo_nome:
        st.session_state.materiali.append({
            "nome": nuovo_nome, 
            "quantita": nuova_qta, 
            "procurato": False, 
            "posizione": ""
        })
        st.success(f"Aggiunto: {nuova_qta}x {nuovo_nome}")
        st.rerun()

# --- 4. LISTA MATERIALI ---
st.divider()
st.subheader("Da preparare e caricare")

indice_da_eliminare = None
indice_da_modificare = None
indice_da_salvare = None
nuovo_nome_modificato = ""
nuova_quantita_modificata = 1

for i, item in enumerate(st.session_state.materiali):
    # Suddiviso l'elenco in 5 colonne per far spazio alla quantità
    col1, col_qta, col2, col3, col4 = st.columns([3, 1, 2, 1, 1])
    
    qta_attuale = item.get('quantita', 1) # Fallback a 1 se ci sono vecchi dati senza quantità
    
    if st.session_state.riga_in_modifica == i:
        with col1:
            nuovo_testo = st.text_input("Modifica", value=item['nome'], key=f"edit_in_{i}", label_visibility="collapsed")
        with col_qta:
            nuova_qta_input = st.number_input("Q.tà", min_value=1, value=qta_attuale, key=f"edit_qta_{i}", label_visibility="collapsed")
        with col2:
            st.write("*(In modifica...)*")
        with col3:
            if st.button("💾", key=f"save_{i}", help="Salva"):
                indice_da_salvare = i
                nuovo_nome_modificato = nuovo_testo
                nuova_quantita_modificata = nuova_qta_input
        with col4:
            if st.button("❌", key=f"cancel_{i}", help="Annulla"):
                st.session_state.riga_in_modifica = None
                st.rerun()
    else:
        with col1:
            procurato = st.checkbox(item['nome'], value=item['procurato'], key=f"check_{i}")
            st.session_state.materiali[i]['procurato'] = procurato
            
        with col_qta:
            st.write(f"**{qta_attuale} pz**")
            
        with col2:
            if procurato:
                posizione = st.selectbox("Posizione", ["Ufficio", "Pedana", "Furgone"], 
                                         index=["Ufficio", "Pedana", "Furgone"].index(item['posizione']) if item['posizione'] in ["Ufficio", "Pedana", "Furgone"] else 0,
                                         key=f"pos_{i}", label_visibility="collapsed")
                st.session_state.materiali[i]['posizione'] = posizione
            else:
                st.session_state.materiali[i]['posizione'] = ""
                st.write("🔴 *Da trovare*")
                
        with col3:
            if st.button("✏️", key=f"edit_btn_{i}", help="Modifica"):
                indice_da_modificare = i
                
        with col4:
            if st.button("🗑️", key=f"del_{i}", help="Elimina"):
                indice_da_eliminare = i
                
    st.divider()

if indice_da_modificare is not None:
    st.session_state.riga_in_modifica = indice_da_modificare
    st.rerun()

if indice_da_salvare is not None:
    st.session_state.materiali[indice_da_salvare]['nome'] = nuovo_nome_modificato
    st.session_state.materiali[indice_da_salvare]['quantita'] = nuova_quantita_modificata
    st.session_state.riga_in_modifica = None
    st.rerun()
    
if indice_da_eliminare is not None:
    st.session_state.materiali.pop(indice_da_eliminare)
    st.session_state.riga_in_modifica = None 
    st.rerun()
