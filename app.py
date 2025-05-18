import streamlit as st
import pandas as pd
from datetime import datetime
from logica import gerar_orcamento_xlsx
import os

import base64

def get_base64_logo(path):
    with open(path, "rb") as image_file:
        return base64.b64encode(image_file.read()).decode()
    
logo_base64 = get_base64_logo("logo.png")


st.set_page_config(page_title="Orçamento Banrisul", layout="wide")


# Cabeçalho com imagem e título
col_logo, col_titulo = st.columns([1, 5])
with col_logo:
    st.image("logo.png", width=80)
with col_titulo:
    st.title("Formulário de Orçamento - Banrisul Centro 0100853/2024")
    datahora_atual = datetime.now()
    st.markdown(f"""
    <div style="display: flex; justify-content: flex-end; align-items: center; font-weight: 600; font-size: 20px; margin-top: 0.5em;">
        <span style="font-size: 24px; margin-right: 10px;">🕒</span>
        <span>{datahora_atual.strftime('%d/%m/%Y %H:%M:%S')}</span>
    </div>
""", unsafe_allow_html=True)




# Dados de identificação
col1, col2, col3, col4 = st.columns(4)
col1.text_input("🔴 NÚMERO RAT:", key="numero_rat")
col2.text_input("🔴 ORÇAMENTISTA:", key="nome_orcamentista")
col3.empty()
col4.empty()

TIPOS = ["CIVIL", "INSTALAÇÕES ELÉTRICAS", "INSTALAÇÕES MECÂNICAS"]

# Estado dinâmico
if "qtd_itens" not in st.session_state:
    st.session_state.qtd_itens = {tipo: 1 for tipo in TIPOS}

# Carregar planilha
try:
    df_ref = pd.read_excel("referencia.xlsx")
except PermissionError:
    st.error("❌ Feche o arquivo 'referencia.xlsx' no Excel antes de iniciar o sistema.")
    st.stop()

desconto = 0.1013
bdi = 0.2928

for tipo in TIPOS:
    with st.expander(f"🧾 ITENS - {tipo.upper()}", expanded=False):
        subtotal_material = 0.0
        subtotal_mao_obra = 0.0
        subtotal_total = 0.0

        opcoes = df_ref[df_ref["TIPO"] == tipo][["ITENS", "DESCRIÇÃO"]]
        itens_formatados = [row["ITENS"] for _, row in opcoes.iterrows()]

        for i in range(st.session_state.qtd_itens[tipo]):
            cols = st.columns([1.2, 3, 1, 1, 1, 1, 1])
            selecionado = cols[0].selectbox(f"Item {i+1} - {tipo}", [""] + itens_formatados, key=f"{tipo}_item_{i}")

            desc = ""
            unid = ""
            mat = 0.0
            mao = 0.0

            if selecionado:
                linha = df_ref[(df_ref["ITENS"] == selecionado) & (df_ref["TIPO"] == tipo)]
                if not linha.empty:
                    linha = linha.iloc[0]
                    desc = linha.get("DESCRIÇÃO", "")
                    unid = linha.get("UNID.", "")
                    mat = float(linha["CUSTOS UNITÁRIOS R$MATERIAL"]) if pd.notna(linha["CUSTOS UNITÁRIOS R$MATERIAL"]) else 0.0
                    mao = float(linha["CUSTOS UNITÁRIOS R$MÃO DE OBRA"]) if pd.notna(linha["CUSTOS UNITÁRIOS R$MÃO DE OBRA"]) else 0.0

            cols[1].text_area("Descrição", value=desc, key=f"{tipo}_desc_{i}", disabled=True, height=80)
            quant = cols[2].number_input("Quant.", key=f"{tipo}_quant_{i}", min_value=0.0, step=1.0)
            cols[3].text_input("Unid.", value=unid, key=f"{tipo}_unid_{i}", disabled=True)
            cols[4].number_input("Material", value=mat, key=f"{tipo}_mat_{i}", disabled=True)
            cols[5].number_input("Mão de Obra", value=mao, key=f"{tipo}_mao_{i}", disabled=True)
            total_bruto = quant * (mat + mao)
            total_com_desconto = total_bruto * (1 - desconto)
            cols[6].text_input("Custo Total", value=f"R$ {total_com_desconto:,.2f}" if total_com_desconto else "", key=f"{tipo}_total_{i}", disabled=True)

            subtotal_material += quant * mat
            subtotal_mao_obra += quant * mao
            subtotal_total += total_com_desconto

        st.markdown(f"**Subtotal {tipo.upper()}**")
        sub1, sub2, sub3, _, _, _ = st.columns([1, 1, 1, 0.5, 0.5, 0.5])
        sub1.text_input("Material", value=f"R$ {subtotal_material:,.2f}", key=f"{tipo}_subtotal_material_display_{tipo}", disabled=True)
        sub2.text_input("Mão de Obra", value=f"R$ {subtotal_mao_obra:,.2f}", key=f"{tipo}_subtotal_mao_display_{tipo}", disabled=True)
        sub3.text_input("Custo Total", value=f"R$ {subtotal_total:,.2f}", key=f"{tipo}_subtotal_total_display_{tipo}", disabled=True)

        st.caption("O custo total já inclui desconto de 10,13% aplicado sobre o somatório de material e mão de obra.")

        col_add, col_rem = st.columns([1, 1])
        with col_add:
            if st.button(f"➕ Add {tipo}", key=f"add_{tipo}"):
                st.session_state.qtd_itens[tipo] += 1
                st.rerun()
        with col_rem:
            if st.session_state.qtd_itens[tipo] > 1 and st.button(f"➖ Remover {tipo}", key=f"rem_{tipo}"):
                st.session_state.qtd_itens[tipo] -= 1
                st.rerun()

        st.session_state[f"{tipo}_subtotal_material"] = subtotal_material
        st.session_state[f"{tipo}_subtotal_mao"] = subtotal_mao_obra
        st.session_state[f"{tipo}_subtotal_total"] = subtotal_total

# Variável para armazenar o arquivo gerado
arquivo_gerado = None

with st.form("formulario_orcamento"):
    total_geral_material = sum(st.session_state.get(f"{tipo}_subtotal_material", 0.0) for tipo in TIPOS)
    total_geral_mao_obra = sum(st.session_state.get(f"{tipo}_subtotal_mao", 0.0) for tipo in TIPOS)
    total_geral = sum(st.session_state.get(f"{tipo}_subtotal_total", 0.0) for tipo in TIPOS)

    st.markdown("**TOTAL GERAL**")
    tg1, tg2, tg3, _, _, _ = st.columns([1, 1, 1, 0.5, 0.5, 0.5])
    tg1.text_input("Material", value=f"R$ {total_geral_material:,.2f}", key="total_geral_material_display", disabled=True)
    tg2.text_input("Mão de Obra", value=f"R$ {total_geral_mao_obra:,.2f}", key="total_geral_mao_display", disabled=True)
    tg3.text_input("Custo Total R$", value=f"R$ {total_geral:,.2f}", key="total_geral_total_display", disabled=True)

    total_com_bdi = total_geral * (1 + bdi)

    st.markdown("**TOTAL COM BDI**")
    bdi1, bdi2, bdi3, _, _, _ = st.columns([1, 1, 1, 0.5, 0.5, 0.5])
    bdi1.text_input("Material", value=f"R$ {total_geral_material * (1 + bdi):,.2f}", key="bdi_material_display", disabled=True)
    bdi2.text_input("Mão de Obra", value=f"R$ {total_geral_mao_obra * (1 + bdi):,.2f}", key="bdi_mao_display", disabled=True)
    bdi3.text_input("Custo Total R$", value=f"R$ {total_com_bdi:,.2f}", key="bdi_total_display", disabled=True)

    gerar = st.form_submit_button("Gerar Arquivo XLSX")

    if gerar:
        referencia_path = "referencia.xlsx"
        modelo_path = "Planilha orçamento.xlsx"

        itens_selecionados = []
        for tipo in TIPOS:
            for i in range(st.session_state.qtd_itens[tipo]):
                cod_selecionado = st.session_state.get(f"{tipo}_item_{i}", "")
                if not cod_selecionado:
                    continue
                quant = st.session_state.get(f"{tipo}_quant_{i}", 0)
                if cod_selecionado and quant > 0:
                    itens_selecionados.append({"item": cod_selecionado, "quant": quant, "tipo": tipo})

        if itens_selecionados:
            numero_rat = st.session_state.get("numero_rat", "")
            nome = st.session_state.get("nome_orcamentista", "")
            datahora = datahora_atual.strftime("%d/%m/%Y %H:%M:%S")

            # Salva log
            log_df = pd.DataFrame([[numero_rat, nome, datahora]], columns=["RAT", "ORÇAMENTISTA", "DATAHORA"])
            if os.path.exists("log_orcamentos.csv"):
                log_df.to_csv("log_orcamentos.csv", mode="a", header=False, index=False)
            else:
                log_df.to_csv("log_orcamentos.csv", index=False)

            nome_arquivo = f"Planilha Orçamentária BANRISUL CENTRO - Ocorrência{numero_rat}.xlsx"
            arquivo_gerado = gerar_orcamento_xlsx(itens_selecionados, referencia_path, nome_arquivo)

        else:
            st.warning("Nenhum item válido preenchido.")

# Fora do formulário
if arquivo_gerado:
    st.success("✅ Arquivo gerado com sucesso!")
    st.download_button(
    "📥 Baixar",
    data=arquivo_gerado,
    file_name=nome_arquivo,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
    
# Ocultar colunas indesejadas
colunas_exibir = ["GRUPO", "ITENS", "DESCRIÇÃO", "QUANT.", "UNID.", "CUSTOS UNITÁRIOS R$MATERIAL", "CUSTOS UNITÁRIOS R$MÃO DE OBRA"]
df_filtrado = df_ref[colunas_exibir]

# Exibir tabela limpa
with st.expander("🔍 Visualizar Tabela de Referência", expanded=False):
    st.dataframe(df_filtrado, use_container_width=True)


    
st.markdown(f"""
<hr style='margin-top: 2em;'>
<div style='display: flex; justify-content: center; align-items: center; gap: 10px; padding: 10px 0; color: #666; font-size: 14px;'>
    <img src="data:image/png;base64,{logo_base64}" alt="Logo" style="height: 24px;" />
    <span>Sistema de Orçamentos - Gennesis Engenharia • <span>Todos os direitos reservados © 2025</span>    <strong>By Vilmar William Ferreira</strong></span>
</div>
""", unsafe_allow_html=True)



