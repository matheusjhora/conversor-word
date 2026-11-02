import docx
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import re
import streamlit as st

# Configuração da página para telemóvel
st.set_page_config(
    page_title="Conversor para Word", page_icon="📄", layout="centered"
)

st.title("📄 Conversor de Texto/Fórmulas para Word (.docx)")
st.write(
    "Cola o teu texto com fórmulas abaixo para gerar o ficheiro formatado para o Word."
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

        # Verifica se a linha é um título principal
        if line_str.startswith("BLOCO") or line_str.startswith("MATERIAL"):
            run = p.add_run(line_str)
            run.bold = True
            run.font.size = docx.shared.Pt(14)
        # Verifica se é uma fórmula (contém '=' e operadores matemáticos)
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


# Interface no Telemóvel
user_text = st.text_area("Cola o teu conteúdo aqui:", height=300)

if st.button("Gerar Ficheiro Word"):
    if user_text.strip():
        file_path = generate_word_doc(user_text)
        with open(file_path, "rb") as fp:
            st.download_button(
                label="📥 Descarregar ficheiro .docx",
                data=fp,
                file_name="estudo_formatado.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )
    else:
        st.warning("Por favor, cola algum texto antes de gerar.")
