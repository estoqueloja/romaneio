import streamlit as st
import pandas as pd
from datetime import datetime
import os
from utils import validate_currency, validate_date, save_to_excel, delete_row_from_excel
import tempfile
import openpyxl

st.set_page_config(
    page_title="Gerenciador de Romaneio",
    page_icon="📋",
    layout="centered"
)

# Custom CSS para melhorar o layout
st.markdown("""
    <style>
    /* Estilo geral */
    .stApp {
        font-family: Arial, sans-serif;
        background-color: #f4f6f9;
        padding: 2rem;
    }
    /* Título principal */
    h1 {
        text-align: center;
        color: #333;
        margin-bottom: 2rem;
    }
    /* Subtítulo */
    h3 {
        color: #555;
        margin-bottom: 1rem;
    }
    /* Formulário */
    .stForm {
        background-color: #ffffff;
        padding: 2rem;
        border-radius: 10px;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
    }
    /* Botões */
    .stButton>button {
        width: 100%;
        margin: 0.5rem 0;
        padding: 0.75rem;
        border-radius: 5px;
        font-size: 1rem;
    }
    /* Inputs */
    .stTextInput>div>div>input,
    .stDateInput>div>div>input {
        padding: 0.75rem;
        border-radius: 5px;
        border: 1px solid #ccc;
    }
    /* Tabela */
    .stTable {
        margin-top: 2rem;
        background-color: #ffffff;
        padding: 1rem;
        border-radius: 10px;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
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
        st.subheader("📝 Informações Iniciais")
        uploaded_file = st.file_uploader("📂 Carregar Romaneio Existente", type=['xlsx'])
        if uploaded_file:
            temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx')
            temp_file.write(uploaded_file.getvalue())
            temp_file.close()
            st.session_state.current_file = temp_file.name
            wb = openpyxl.load_workbook(temp_file.name)
            sheet_name = wb.sheetnames[0]
            ws = wb[sheet_name]
            st.session_state.cidade = ws['A1'].value if ws['A1'].value in CIDADES else CIDADES[0]
            st.session_state.current_sheet = sheet_name
            try:
                st.session_state.data = datetime.strptime(sheet_name, '%d_%m_%Y').date()
            except:
                st.session_state.data = datetime.now().date()
            st.session_state.step = 2
            st.rerun()

        st.markdown("---")
        st.subheader("➕ Criar Novo Romaneio")
        with st.form("initial_form"):
            cidade = st.selectbox("Cidade", CIDADES, index=CIDADES.index(st.session_state.cidade))
            data_romaneio = st.date_input("Data do Romaneio", value=st.session_state.data, format="DD/MM/YYYY")
            submitted = st.form_submit_button("Criar Romaneio")
            if submitted:
                st.session_state.cidade = cidade
                st.session_state.data = data_romaneio
                if not st.session_state.current_file:
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
        st.subheader(f"📋 Romaneio - {st.session_state.cidade}")
        nova_data = st.date_input("📅 Data do Romaneio", value=st.session_state.data, format="DD/MM/YYYY")
        if nova_data != st.session_state.data:
            st.session_state.data = nova_data
            st.session_state.current_sheet = nova_data.strftime('%d_%m_%Y')
            st.rerun()

        if 'widget_key' not in st.session_state:
            st.session_state.widget_key = 0

        with st.form(key=f"romaneio_form_{st.session_state.widget_key}"):
            st.markdown("#### 📝 Adicionar Item")
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
                "💳 Forma de Pagamento",
                payment_options,
                key=f"pagamento_{st.session_state.widget_key}"
            )
            valor = st.text_input(
                "💰 Valor a Pagar (R$)",
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
                initial_data = [st.session_state.cidade, nova_data.strftime('%d/%m/%Y')]
                details_data = [numero_pedido, revendedor, pagamento, f"R$ {valor_float:.2f}"]
                success, message = save_to_excel(
                    [initial_data, details_data], 
                    st.session_state.current_file,
                    st.session_state.current_sheet,
                    append_mode=True
                )
                if success:
                    st.success("✅ Item adicionado com sucesso!")
                    st.session_state.widget_key += 1
                    st.rerun()
                else:
                    st.error(f"❌ Erro ao salvar: {message}")
            elif submitted_save:
                st.success("✅ Romaneio salvo com sucesso!")
                st.session_state.show_download = True

        # Display current sheet data
        st.markdown("---")
        st.subheader("📋 Itens Adicionados")
        df = read_sheet_data(st.session_state.current_file, st.session_state.current_sheet)
        if df is not None and not df.empty:
            # Exibir todas as linhas em uma única tabela
            st.write("**Tabela de Itens Adicionados:**")
            for i, row in df.iterrows():
                cols = st.columns([4, 1])  # Colunas para os dados e o botão
                with cols[0]:
                    st.table(pd.DataFrame([row]))
                with cols[1]:
                    if st.button(f"❌ Excluir {i}", key=f"delete_{i}"):
                        success, message = delete_row_from_excel(
                            st.session_state.current_file,
                            st.session_state.current_sheet,
                            i
                        )
                        if success:
                            st.success(f"✅ Linha {i + 1} excluída com sucesso!")
                            st.rerun()
                        else:
                            st.error(f"❌ Erro ao excluir linha {i + 1}: {message}")
        else:
            st.info("ℹ️ Nenhum item adicionado ainda.")

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
