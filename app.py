import docx
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import pypdf
import streamlit as st

# Configuração da página para telemóvel/PC
st.set_page_config(
    page_title="Conversor para Word", page_icon="📄", layout="centered"
)

st.title("📄 Conversor de Texto/Fórmulas para Word (.docx)")
st.write(
    "Envia um ficheiro (.txt ou .pdf) ou cola o texto formatado abaixo para gerar o documento Word."
)


def add_math_equation(paragraph, math_text):
    """Adiciona a fórmula ao parágrafo com a fonte matemática do Word"""
    run = paragraph.add_run(math_text)
    run.font.name = "Cambria Math"
    run.font.italic = True


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

    output_path = "documento_formatado.docx"
    doc.save(output_path)
    return output_path


# 1. Opção de Upload de Ficheiro
uploaded_file = st.file_uploader(
    "Carrega um ficheiro (.txt ou .pdf)", type=["txt", "pdf"]
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

# 2. Caixa de Texto
user_text = st.text_area(
    "Ou cola o teu conteúdo aqui:", value=extracted_text, height=300
)

# 3. Nome Personalizado do Ficheiro
file_name_input = st.text_input(
    "Nome do ficheiro para salvar (sem extensão):", value="estudo_formatado"
)

# 4. Botão de Geração
if st.button("Gerar Ficheiro Word"):
    if user_text.strip():
        file_path = generate_word_doc(user_text)

        clean_name = (
            file_name_input.strip() if file_name_input.strip() else "estudo"
        )
        final_filename = f"{clean_name}.docx"

        with open(file_path, "rb") as fp:
            st.download_button(
                label="📥 Descarregar ficheiro .docx",
                data=fp,
                file_name=final_filename,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )
    else:
        st.warning(
            "Por favor, carrega um ficheiro ou cola texto antes de gerar."
        )
