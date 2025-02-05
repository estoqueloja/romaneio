import streamlit as st
import pandas as pd
from datetime import datetime
import os
from utils import validate_currency, validate_date, save_to_excel
import tempfile
import openpyxl

st.set_page_config(
    page_title="Gerenciador de Romaneio",
    page_icon="📋",
    layout="centered"
)

# Custom CSS
st.markdown("""
    <style>
    .stButton>button {
        width: 100%;
    }
    .stTextInput>div>div>input {
        color: #000000;
    }
    .main {
        padding: 2rem;
    }
    </style>
    """, unsafe_allow_html=True)

def initialize_excel_file(data):
    """Initialize a new Excel file with basic structure"""
    wb = openpyxl.Workbook()
    sheet = wb.active
    sheet.title = data.strftime('%d_%m_%Y')
    return wb

def read_sheet_data(file_path, sheet_name):
    """Read data from Excel sheet and return as DataFrame"""
    try:
        wb = openpyxl.load_workbook(file_path)
        
        # Verifica se a aba existe no arquivo. Se não, retorna None sem avisar o usuário.
        if sheet_name not in wb.sheetnames:
            return None
        # Se a aba existe, lê os dados
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=0)
        if not df.empty:
            return df
    except Exception as e:
        # Se ocorrer algum erro ao tentar abrir o arquivo ou ler a aba, o erro será ignorado sem mensagem.
        return None
    return None

def main():
    st.title("📋 Gerenciador de Romaneio")
    # Lista de cidades disponíveis
    CIDADES = ["Paulínia", "Monte Mor", "Santo Antônio de Posse"]
    # Initialize session state
    if 'step' not in st.session_state:
        st.session_state.step = 1
    if 'current_file' not in st.session_state:
        st.session_state.current_file = None
    if 'show_download' not in st.session_state:
        st.session_state.show_download = False
    if 'cidade' not in st.session_state:
        st.session_state.cidade = CIDADES[0]  # Default to first city
    if 'data' not in st.session_state:
        st.session_state.data = datetime.now().date()  # Corrigido para usar datetime corretamente
    if 'current_sheet' not in st.session_state:
        st.session_state.current_sheet = None

    # Step 1: Cidade e Data
    if st.session_state.step == 1:
        st.subheader("Informações Iniciais")
        # Opção para carregar arquivo existente
        uploaded_file = st.file_uploader("📂 Carregar Romaneio Existente", type=['xlsx'])
        if uploaded_file:
            # Create temporary file and save uploaded content
            temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx')
            temp_file.write(uploaded_file.getvalue())
            temp_file.close()
            st.session_state.current_file = temp_file.name
            # Carregar workbook para pegar a data da primeira aba
            wb = openpyxl.load_workbook(temp_file.name)
            sheet_name = wb.sheetnames[0]
            ws = wb[sheet_name]
            st.session_state.cidade = ws['A1'].value if ws['A1'].value in CIDADES else CIDADES[0]
            st.session_state.current_sheet = sheet_name
            # Parse date from sheet name (format: dd_mm_yyyy)
            try:
                st.session_state.data = datetime.strptime(sheet_name, '%d_%m_%Y').date()
            except:
                st.session_state.data = datetime.now().date()
            st.session_state.step = 2
            st.rerun()
        # Separador visual
        st.markdown("---")
        st.subheader("Criar Novo Romaneio")
        with st.form("initial_form"):
            cidade = st.selectbox("Cidade", CIDADES, index=CIDADES.index(st.session_state.cidade))
            data_romaneio = st.date_input("Data do Romaneio", value=st.session_state.data, format="DD/MM/YYYY")
            submitted = st.form_submit_button("Criar Romaneio")
            if submitted:
                st.session_state.cidade = cidade
                st.session_state.data = data_romaneio
                if not st.session_state.current_file:
                    # Create new Excel file only if not already loaded
                    temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx')
                    temp_file.close()
                    wb = initialize_excel_file(data_romaneio)
                    wb.save(temp_file.name)
                    st.session_state.current_file = temp_file.name
                st.session_state.current_sheet = data_romaneio.strftime('%d_%m_%Y')
                st.session_state.step = 2
                st.rerun()

    # Step 2: Detalhes do Romaneio
    elif st.session_state.step == 2:
        st.subheader(f"Romaneio - {st.session_state.cidade}")
        # Campo de data
        nova_data = st.date_input("Data do Romaneio", value=st.session_state.data, format="DD/MM/YYYY")
        # Update session state data when user changes the date
        if nova_data != st.session_state.data:
            st.session_state.data = nova_data
            st.session_state.current_sheet = nova_data.strftime('%d_%m_%Y')
            st.rerun()

        # Chave dinâmica para forçar a recriação dos widgets
        if 'widget_key' not in st.session_state:
            st.session_state.widget_key = 0

        with st.form(key=f"romaneio_form_{st.session_state.widget_key}"):
            # Campos do formulário
            numero_pedido = st.text_input(
                "Número do Pedido",
                placeholder="Digite o número do pedido",
                max_chars=9,
                key=f"numero_pedido_{st.session_state.widget_key}"
            )
            if numero_pedido and not numero_pedido.isdigit():
                st.error("O número do pedido deve conter apenas números.")
                return
            revendedor = st.text_input(
                "Nome do Revendedor",
                placeholder="Digite o nome do revendedor",
                key=f"revendedor_{st.session_state.widget_key}"
            )
            payment_options = ["Dinheiro", "Cartão", "Boleto"]
            pagamento = st.selectbox(
                "Forma de Pagamento",
                payment_options,
                key=f"pagamento_{st.session_state.widget_key}"
            )
            valor = st.text_input(
                "Valor a Pagar (R$)",
                placeholder="0,00",
                key=f"valor_{st.session_state.widget_key}"
            )
            col1, col2, col3 = st.columns(3)
            with col1:
                submitted_add = st.form_submit_button("➕ Adicionar")
            with col2:
                submitted_save = st.form_submit_button("💾 Salvar")
            with col3:
                if st.form_submit_button("🔄 Tela Inicial"):
                    st.session_state.step = 1
                    st.rerun()

            if submitted_add:
                if not numero_pedido:
                    st.error("Por favor, preencha o número do pedido.")
                    return
                if len(numero_pedido) < 9:
                    st.error("O número do pedido deve ter 9 dígitos.")
                    return
                if not revendedor:
                    st.error("Por favor, preencha o nome do revendedor.")
                    return
                revendedor = revendedor.upper()
                valor_float, error = validate_currency(valor)
                if error:
                    st.error(error)
                    return
                # Prepare data
                initial_data = [st.session_state.cidade, nova_data.strftime('%d/%m/%Y')]
                details_data = [numero_pedido, revendedor, pagamento, f"R$ {valor_float:.2f}"]
                # Save to Excel
                success, message = save_to_excel(
                    [initial_data, details_data], 
                    st.session_state.current_file,
                    st.session_state.current_sheet,
                    append_mode=True
                )
                if success:
                    st.success("Item adicionado com sucesso!")
                    
                    # Incrementa a chave dinâmica para forçar a recriação dos widgets
                    st.session_state.widget_key += 1

                    # Força a interface a recarregar
                    st.rerun()
                else:
                    st.error(f"Erro ao salvar: {message}")
            elif submitted_save:
                st.success("Romaneio salvo com sucesso!")
                st.session_state.show_download = True
                    
        # Display current sheet data
        st.markdown("---")
        st.subheader("📋 Itens Adicionados")
        df = read_sheet_data(st.session_state.current_file, st.session_state.current_sheet)
        if df is not None and not df.empty:
            st.dataframe(df, hide_index=True)
        else:
            st.info("Nenhum item adicionado ainda.")
        # Download button outside the form
        if st.session_state.show_download:
            with open(st.session_state.current_file, 'rb') as f:
                st.download_button(
                    label="📥 Baixar Planilha",
                    data=f,
                    file_name=f"Romaneio_{st.session_state.cidade}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

if __name__ == "__main__":
    main()
