import streamlit as st
import pandas as pd
import pypdf
import re
import io

st.set_page_config(page_title="Sistema de Abastecimento - CD Loja 100", layout="wide")

st.title("📦 Distribuição e Abastecimento de Estoque (CD Loja 100)")
st.subheader("Upload do Relatório de Vendas e Estoque")

def processar_pdf(file_bytes):
    reader = pypdf.PdfReader(file_bytes)
    linhas = []
    for page in reader.pages:
        linhas.extend(page.extract_text().split('\n'))
    
    dados, item_atual = [], {}
    lojas = ['0001', '0004', '0006', '0100', '0200']
    
    for l in linhas:
        l_str = l.strip()
        if '-' in l_str and re.match(r'^\d{6,7}-', l_str):
            parts = l_str.split('-', 1)
            item_atual = {'PLU': parts[0].strip(), 'Descrição': parts[1].strip()}
            for loja in lojas:
                item_atual[f'Est_{loja}'], item_atual[f'Giro_{loja}'] = 0, 0
                
        for loja in lojas:
            if l_str.startswith(loja):
                nums = [int(t) for t in l_str.split() if t.isdigit() or (t.startswith('-') and t[1:].isdigit())]
                if len(nums) >= 5:
                    item_atual[f'Giro_{loja}'] = sum(nums[:4])
                    item_atual[f'Est_{loja}'] = nums[4]
                if loja == '0200' and 'PLU' in item_atual:
                    dados.append(item_atual.copy())

    df = pd.DataFrame(dados).drop_duplicates(subset=['PLU'])
    resultado = []
    lojas_ponta = ['0001', '0200', '0004', '0006']
    
    for _, r in df.iterrows():
        est_cd = max(0, r['Est_0100'])
        giro_cd = max(0, r['Giro_0100'])
        
        if est_cd <= 0:
            continue
            
        reserva_cd = min(est_cd, giro_cd)
        disponivel = est_cd - reserva_cd
        if disponivel <= 0:
            continue
            
        nec = {l: max(0, r[f'Giro_{l}'] - max(0, r[f'Est_{l}'])) for l in lojas_ponta}
        tot_nec = sum(nec.values())
        dist = {l: 0 for l in lojas_ponta}
        
        if tot_nec > 0:
            shares = {l: (nec[l] / tot_nec) * disponivel for l in lojas_ponta}
            inteiros = {l: int(shares[l]) for l in lojas_ponta}
            resto = disponivel - sum(inteiros.values())
            fracs = sorted([(l, shares[l] - inteiros[l]) for l in lojas_ponta], key=lambda x: x[1], reverse=True)
            for i in range(resto):
                inteiros[fracs[i][0]] += 1
            dist = inteiros
        else:
            com_giro = [l for l in lojas_ponta if r[f'Giro_{l}'] > 0]
            if com_giro:
                dist[min(com_giro, key=lambda l: r[f'Est_{l}'])] = disponivel

        resultado.append({
            'PLU': r['PLU'], 'Descrição': r['Descrição'],
            'Estoque_CD': est_cd, 'Reserva_CD': reserva_cd,
            'Disponivel_Transf': disponivel,
            'Loja_01': dist['0001'], 'Loja_02': dist['0200'], 
            'Loja_04': dist['0004'], 'Loja_06': dist['0006']
        })

    return pd.DataFrame(resultado)

arquivo_uploaded = st.file_uploader("Arraste ou selecione o arquivo PDF do relatório", type=['pdf'])

if arquivo_uploaded is not None:
    with st.spinner('Processando relatório...'):
        df_resultado = processar_pdf(arquivo_uploaded)
    
    if not df_resultado.empty:
        st.success(f"Sucesso! {len(df_resultado)} produtos processados.")
        
        col1, col2, col3, col4, col5 = st.columns(5)
        col1.metric("Disponível CD", int(df_resultado['Disponivel_Transf'].sum()))
        col2.metric("Loja 01", int(df_resultado['Loja_01'].sum()))
        col3.metric("Loja 02", int(df_resultado['Loja_02'].sum()))
        col4.metric("Loja 04", int(df_resultado['Loja_04'].sum()))
        col5.metric("Loja 06", int(df_resultado['Loja_06'].sum()))
        
        st.dataframe(df_resultado, use_container_width=True)
        
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df_resultado.to_excel(writer, index=False, sheet_name='Abastecimento_CD')
            
        st.download_button("📥 Baixar Excel (.xlsx)", buffer.getvalue(), "Sugestao_Abastecimento_CD.xlsx")
