import pandas as pd
from datetime import datetime
import requests
import re

# === CONFIGURAÇÃO DO BOT ===
TELEGRAM_TOKEN = "8019760498:AAFc9Ro215yCe2c4hFcOJX3UXLJwPFGLNmU"
TELEGRAM_CHAT_ID = "-4813455628"
TELEGRAM_CHAT_ID_CONTABILIDADE = "-4847853376"  # Altere para o ID real do grupo da contabilidade

# =============================
# FUNÇÃO DE LIMPAR NF
# =============================
def limpar_nf(texto):
    if pd.isna(texto):
        return None
    texto = str(texto).upper().strip()
    texto = re.sub(r'^(NF|NOTA)\s*', '', texto)
    match = re.search(r'\d+', texto)
    return match.group(0) if match else None

# =============================
# FUNÇÃO DE SAUDAÇÃO AUTOMÁTICA
# =============================
def saudacao_automatica():
    hora = datetime.now().hour
    if hora < 12:
        return "Bom dia"
    elif 12 <= hora < 18:
        return "Boa tarde"
    else:
        return "Boa noite"

# =============================
# 1. Ler CONTAS A RECEBER
# =============================
try:
    receber_df = pd.read_csv(
        "Contas/Documentos a receber (3).csv",
        sep=";",
        encoding="latin1",
        skiprows=6,
        header=0
    )
    print("✔️ Contas a receber carregadas.")

    # Limpar NF
    receber_df['NF'] = receber_df['Núm.'].apply(limpar_nf)

    # Remover ESCRITORIO CAICARA
    receber_df = receber_df[receber_df['Obra'].str.upper() != 'ESCRITORIO CAICARA']

    # Converter valor
    receber_df['Valor_receber'] = pd.to_numeric(
        receber_df["Líquido (R$)"].astype(str).str.replace(".", "").str.replace(",", "."),
        errors='coerce'
    ).fillna(0)

    # Filtrar apenas NFs válidas
    receber_df_limpo = receber_df[['NF', 'Obra', 'Valor_receber']].dropna(subset=['NF'])

    # Extrair lista única de NFs de interesse
    lista_nf_interesse = receber_df_limpo['NF'].unique().tolist()

    print(f"✔️ NFs extraídas do Contas a Receber: {lista_nf_interesse}")
    print(f"✔️ Contas a receber (todas limpas): {len(receber_df_limpo)} linhas.")

except Exception as e:
    print(f"❌ Erro ao processar Contas a Receber: {e}")
    exit()

# =============================
# 2. Ler CONTAS A PAGAR
# =============================
try:
    pagar_df = pd.read_csv(
        "Contas/Documentos a pagar.csv",
        sep=";",
        encoding="latin1",
        skiprows=6,
        header=0
    )
    print("✔️ Contas a pagar carregadas.")

    pagar_df['NF'] = pagar_df['Núm.'].apply(limpar_nf)
    pagar_df['Valor_pagar'] = pd.to_numeric(
        pagar_df["Líquido (R$)"].astype(str).str.replace(".", "").str.replace(",", "."),
        errors='coerce'
    ).fillna(0)

    pagar_df_limpo = pagar_df[['NF', 'Valor_pagar']].dropna(subset=['NF'])
    print(f"✔️ Contas a pagar (todas limpas): {len(pagar_df_limpo)} linhas.")

except Exception as e:
    print(f"❌ Erro ao processar Contas a Pagar: {e}")
    exit()

# =============================
# 3. Filtrar PAGAR usando NFs do RECEBER
# =============================
pagar_filtrado = pagar_df_limpo[pagar_df_limpo['NF'].isin(lista_nf_interesse)]
print(f"✔️ NF encontradas no PAGAR: {len(pagar_filtrado)}")

# =============================
# 4. Juntar RECEBER + PAGAR
# =============================
try:
    resultado = pd.merge(
        receber_df_limpo,
        pagar_filtrado,
        on="NF",
        how="left"
    ).fillna(0)

    resultado["Diferença"] = resultado["Valor_receber"] - resultado["Valor_pagar"]

    # Renomear colunas para exportação
    resultado.rename(columns={
        'Valor_receber': 'VALOR TOTAL',
        'Valor_pagar': 'IMPOSTOS RETIDOS',
        'Diferença': 'LIQUIDO'
    }, inplace=True)

    print(f"✔️ Merge concluído: {len(resultado)} linhas.")

except Exception as e:
    print(f"❌ Erro ao consolidar dados: {e}")
    exit()

# =============================
# 5. Resumo Geral
# =============================
try:
    total_receber = resultado['VALOR TOTAL'].sum()
    total_pagar = resultado['IMPOSTOS RETIDOS'].sum()
    saldo_liquido = resultado['LIQUIDO'].sum()

    resumo_df = pd.DataFrame({
        "Descrição": ["Total a Receber", "Total a Pagar", "Saldo Líquido"],
        "Valor": [total_receber, total_pagar, saldo_liquido]
    })

    print("\n✅ ✅ ✅ Resumo Final ✅ ✅ ✅")
    print(f"🔹 Total Receber : R$ {total_receber:,.2f}")
    print(f"🔹 Total Pagar   : R$ {total_pagar:,.2f}")
    print(f"🔹 Saldo Líquido : R$ {saldo_liquido:,.2f}")

except Exception as e:
    print(f"❌ Erro ao calcular resumo: {e}")
    exit()

# =============================
# 6. Gerar Planilha Excel
# =============================
try:
    nome_arquivo = f"Relatorio_Detalhado_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"

    with pd.ExcelWriter(nome_arquivo, engine='xlsxwriter') as writer:
        resultado.to_excel(writer, sheet_name='Detalhado', index=False)
        resumo_df.to_excel(writer, sheet_name='Resumo Geral', index=False)

        workbook = writer.book
        money_fmt = workbook.add_format({'num_format': 'R$ #,##0.00'})

        ws1 = writer.sheets['Detalhado']
        ws1.set_column('A:A', 10)
        ws1.set_column('B:B', 40)
        ws1.set_column('C:E', 20, money_fmt)

        ws2 = writer.sheets['Resumo Geral']
        ws2.set_column('A:A', 25)
        ws2.set_column('B:B', 20, money_fmt)

    print(f"\n✅ Planilha '{nome_arquivo}' criada com sucesso!")

except Exception as e:
    print(f"❌ Erro ao gerar a planilha: {e}")
    exit()

# =============================
# 7. Enviar MENSAGEM no Telegram (Financeiro)
# =============================
try:
    mensagem = (
        f"*📊 Resumo Financeiro (Filtrado)*\n\n"
        f"- NFs encontradas: {', '.join(lista_nf_interesse)}\n"
        f"- Total a Receber: R$ {total_receber:,.2f}\n"
        f"- Total a Pagar: R$ {total_pagar:,.2f}\n"
        f"- Saldo Líquido: R$ {saldo_liquido:,.2f}"
    )

    send_message_url = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendMessage"
    response = requests.post(send_message_url, data={
        "chat_id": TELEGRAM_CHAT_ID,
        "text": mensagem,
        "parse_mode": "Markdown"
    })

    if response.ok:
        print("✅ Mensagem enviada para o grupo financeiro!")
    else:
        print(f"❌ Erro ao enviar mensagem: {response.text}")

    send_file_url = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendDocument"
    with open(nome_arquivo, "rb") as f:
        response = requests.post(send_file_url, data={"chat_id": TELEGRAM_CHAT_ID}, files={"document": f})

    if response.ok:
        print("✅ Arquivo Excel enviado para o grupo financeiro!")
    else:
        print(f"❌ Erro ao enviar arquivo: {response.text}")

except Exception as e:
    print(f"❌ Erro ao enviar para financeiro: {e}")

# =============================
# 8. Perguntar envio para contabilidade
# =============================
try:
    resposta = input("\nDeseja enviar também para o grupo de contabilidade? (s/n): ").strip().lower()

    if resposta == 's':
        saudacao = saudacao_automatica()

        mensagem_contabilidade = (
            f"{saudacao}!\n\n"
            f"Segue o relatório detalhado para a contabilidade."
        )

        response = requests.post(send_message_url, data={
            "chat_id": TELEGRAM_CHAT_ID_CONTABILIDADE,
            "text": mensagem_contabilidade
        })

        if response.ok:
            print("✅ Mensagem enviada ao grupo de contabilidade!")
        else:
            print(f"❌ Erro ao enviar mensagem: {response.text}")

        with open(nome_arquivo, "rb") as f:
            response = requests.post(send_file_url, data={"chat_id": TELEGRAM_CHAT_ID_CONTABILIDADE}, files={"document": f})

        if response.ok:
            print("✅ Arquivo Excel enviado para o grupo de contabilidade!")
        else:
            print(f"❌ Erro ao enviar arquivo: {response.text}")

    else:
        print("✅ Envio para contabilidade cancelado.")

except Exception as e:
    print(f"❌ Erro na etapa de envio para contabilidade: {e}")
