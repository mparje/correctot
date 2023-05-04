import io
import openai
import streamlit as st
import docx

openai.api_key = os.getenv("OPENAI_API_KEY")

def gpt_correct_prompt(prompt):
    completions = openai.Completion.create(engine="text-davinci-003", prompt=prompt, max_tokens=1024, n=1, stop=None,
                                           temperature=0.5)
    message = completions.choices[0].text.strip()
    return message


def process_document(doc_buffer):
    doc = docx.Document(doc_buffer)
    corrected_doc = docx.Document()

    for paragraph in doc.paragraphs:
        original_text = paragraph.text
        if original_text.strip() == "":
            corrected_doc.add_paragraph()
            continue

        corrected_text = gpt_correct_prompt(
            f"Reescribe el siguiente párrafo corrigiendo errores gramaticales y mejorando el estilo:\n'{original_text}'\n\nTexto corregido:")
        corrected_paragraph = corrected_doc.add_paragraph(corrected_text)
        for run in paragraph.runs:
            if run.bold:
                corrected_paragraph.runs[-1].bold = True
            if run.italic:
                corrected_paragraph.runs[-1].italic = True
            if run.underline:
                corrected_paragraph.runs[-1].underline = True

    return corrected_doc


st.title('Corrección Gramatical Utilizando GPT-3.5-turbo')
st.write('Sube un documento de Word (.docx) y dejaré que GPT-3.5-turbo corrija los errores gramaticales y de estilo.')

uploaded_file = st.file_uploader("Subir archivo", type=[".docx"])

if uploaded_file is not None:
    with io.BytesIO(uploaded_file.getvalue()) as doc_buffer:
        corrected_doc = process_document(doc_buffer)

        with io.BytesIO() as bytes_io:
            corrected_doc.save(bytes_io)
            bytes_io.seek(0)
            st.download_button(label="Descargar documento corregido", data=bytes_io,
                               file_name="documento_corregido.docx")
