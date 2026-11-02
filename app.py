import docx
from docx.oxml import OxmlElement
from google import genai
import pypdf
import streamlit as st

# Configuração da página para celular
st.set_page_config(
    page_title="Conversor & Processador para Word",
    page_icon="📄",
    layout="centered",
)

st.title("📄 Processador Automático para Word (.docx)")
st.write(
    "Insira seu texto/PDF e a IA irá desmembrar o conteúdo em blocos otimizados com fórmulas formatadas para o Word e NotebookLM."
)

# Campo para inserir a Chave da API (Guarda no navegador durante o uso)
api_key = st.text_input(
    "Insira sua API Key do Google AI Studio:", type="password"
)


def add_math_equation(paragraph, math_text):
    """Adiciona a fórmula ao parágrafo com a fonte matemática do Word"""
    run = paragraph.add_run(math_text)
    run.font.name = "Cambria Math"
    run.font.italic = True


def format_with_gemini(raw_text, key):
    """Envia o texto bruto para a API do Gemini desmembrar automaticamente"""
    client = genai.Client(api_key=key)

    prompt = f"""
    Você é um especialista pedagógico em engenharia elétrica. 
    Reorganize o texto bruto a seguir em blocos de estudo autocontidos (BLOCO 1, BLOCO 2...), 
    isolando fórmulas matemáticas em formato linear nativo (sem tabelas complexas, sem sintaxe markdown como #, *, _ ou blocos de código).
    Mantenha os termos técnicos rigorosos e garanta que as fórmulas usem operadores como =, +, -, ×, /, ^, Δ, μ, π.
    
    TEXTO BRUTO:
    {raw_text}
    """

    response = client.models.generate_content(
        model="gemini-flash", contents=prompt
    )
    return response.text


def generate_word_doc(text_input):
    doc = docx.Document()
    lines = text_input.split("\n")

    for line in lines:
        line_str = line.strip()
        if not line_str:
            continue

        p = doc.add_paragraph()

        # Identifica títulos de blocos ou materiais
        if line_str.startswith("BLOCO") or line_str.startswith("MATERIAL"):
            run = p.add_run(line_str)
            run.bold = True
            run.font.size = docx.shared.Pt(14)
        # Identifica fórmulas por operadores matemáticos
        elif "=" in line_str and any(
            op in line_str for op in ["+", "-", "×", "/", "^", "Δ", "μ", "π"]
        ):
            p.paragraph_format.space_before = docx.shared.Pt(4)
            p.paragraph_format.space_after = docx.shared.Pt(4)
            add_math_equation(p, line_str)
        else:
            p.add_run(line_str)

    output_path = "temp_doc.docx"
    doc.save(output_path)
    return output_path


# 1. Entrada de Arquivo / Texto
uploaded_file = st.file_uploader(
    "Carregue um arquivo (.txt ou .pdf)", type=["txt", "pdf"]
)
extracted_text = ""

if uploaded_file is not None:
    if uploaded_file.name.endswith(".txt"):
        extracted_text = uploaded_file.read().decode("utf-8")
    elif uploaded_file.name.endswith(".pdf"):
        pdf_reader = pypdf.PdfReader(uploaded_file)
        for page in pdf_reader.pages:
            text = page.extract_text()
            if text:
                extracted_text += text + "\n"

user_text = st.text_area(
    "Ou cole o texto bruto aqui:", value=extracted_text, height=250
)

# 2. Configuração do Nome do Arquivo Final
file_name_input = st.text_input(
    "Nome do arquivo final (sem extensão):", value="estudo_desmembrado"
)

# 3. Processamento e Geração
if st.button("Processar e Gerar Word"):
    if not api_key:
        st.error(
            "Por favor, informe sua API Key do Google AI Studio no campo acima."
        )
    elif not user_text.strip():
        st.warning("Por favor, carregue um arquivo ou cole um texto.")
    else:
        with st.spinner("Desmembrando o conteúdo via IA..."):
            try:
                # Chama a IA para organizar e formatar o conteúdo
                formatted_text = format_with_gemini(user_text, api_key)

                # Cria o documento Word
                file_path = generate_word_doc(formatted_text)

                clean_name = (
                    file_name_input.strip()
                    if file_name_input.strip()
                    else "estudo"
                )
                final_filename = f"{clean_name}.docx"

                with open(file_path, "rb") as fp:
                    st.download_button(
                        label="📥 Baixar arquivo .docx",
                        data=fp,
                        file_name=final_filename,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    )
                st.success("Conteúdo processado com sucesso!")
            except Exception as e:
                st.error(f"Erro ao processar: {e}")
