import pandas as pd
import os
import requests
from datetime import datetime

# =============================
# Telegram helpers
# =============================

def send_telegram_message(bot_token, chat_id, message, parse_mode='Markdown'):
    """Envia mensagem para o Telegram"""
    url = f"https://api.telegram.org/bot{bot_token}/sendMessage"
    params = {'chat_id': chat_id, 'text': message, 'parse_mode': parse_mode}
    try:
        print(f"[Telegram] Enviando mensagem para chat {chat_id}...")
        response = requests.post(url, params=params)
        response.raise_for_status()
        print("[Telegram] Mensagem enviada com sucesso.")
        return True
    except Exception as e:
        print(f"⚠ Erro ao enviar mensagem: {e}")
        return False


def send_telegram_file(bot_token, chat_id, file_path, caption=""):
    """Envia arquivo para o Telegram"""
    url = f"https://api.telegram.org/bot{bot_token}/sendDocument"
    try:
        print(f"[Telegram] Enviando arquivo {file_path} para chat {chat_id}...")
        with open(file_path, 'rb') as file:
            files = {'document': file}
            data = {'chat_id': chat_id, 'caption': caption}
            response = requests.post(url, files=files, data=data)
            response.raise_for_status()
            print("[Telegram] Arquivo enviado com sucesso.")
            return True
    except Exception as e:
        print(f"⚠ Erro ao enviar arquivo: {e}")
        return False

# =============================
# Relatório Principal (Telegram)
# =============================

def send_main_report(df_grouped, df_totals, total_geral, total_chesf,
                     df_financeiro, output_file_path,
                     bot_token, main_chat_id):
    """Envia mensagem principal para o grupo Telegram"""

    import unicodedata

    def remove_accents(text):
        if not isinstance(text, str):
            return ""
        return ''.join(
            c for c in unicodedata.normalize('NFD', text)
            if unicodedata.category(c) != 'Mn'
        )

    def format_currency(value):
        return f"R$ {value:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')

    SEPARATOR = "━━━━━━━━━━━━━━━━━━━━━━━"

    timestamp = datetime.now().strftime('%d/%m/%Y %H:%M')

    # ================= HEADER =================
    message = (
        f"{SEPARATOR}\n"
        f"📊 *RELATÓRIO DE ORDENS DE COMPRA E DESPESAS DE CARTÃO* 📊\n"
        f"{SEPARATOR}\n\n"
        f"🗓️ *Data/Hora*: {timestamp}\n"
        f"📎 *Arquivo*: `{os.path.basename(output_file_path)}`\n\n"
    )

    # ================= TOTAIS POR EMPRESA =================
    message += f"{SEPARATOR}\n"
    message += "🏢 *TOTAIS ORDENS DE COMPRA POR EMPRESA* 💰\n"
    message += f"{SEPARATOR}\n"
    for _, row in df_totals.iterrows():
        message += f"• {row['Empresa']}\n"
        message += f"  ➜ {format_currency(row['Total Empresa'])}\n"

    # ================= NORMALIZAR EMPRESA =================
    df_totals['Empresa_Norm'] = df_totals['Empresa'].apply(lambda x: remove_accents(x).upper())
    caicara_total_row = df_totals[df_totals['Empresa_Norm'].str.contains('CAICARA')]
    total_caicara = caicara_total_row.iloc[0]['Total Empresa'] if not caicara_total_row.empty else 0
    chesf_mais_caicara = total_chesf + total_caicara

    # ================= DESPESAS DE CARTÃO SOMADAS POR EMPRESA =================
    df_cartoes = (
        df_financeiro
        .groupby(['Empresa', 'Obra'], as_index=False)
        .agg({'Bruto': 'sum'})
    )
    df_cartoes['Empresa_Norm'] = df_cartoes['Empresa'].apply(lambda x: remove_accents(x).upper())

    cartao_caicara = df_cartoes[df_cartoes['Empresa_Norm'].str.contains('CAICARA')]['Bruto'].sum()
    cartao_pallas = df_cartoes[df_cartoes['Empresa_Norm'].str.contains('PALLAS')]['Bruto'].sum()

    # TOTAL PALLAS ORDENS
    pallas_total_row = df_totals[df_totals['Empresa_Norm'].str.contains('PALLAS')]
    total_pallas = pallas_total_row.iloc[0]['Total Empresa'] if not pallas_total_row.empty else 0

    # ================= NOVO AJUSTADO: CHESF - (CAICARA + CARTÃO) =================
    total_chesf_ajustado = total_chesf - (total_caicara + cartao_caicara)

    # ================= RESUMO GERAL =================
    message += f"\n{SEPARATOR}\n"
    message += "📈 *RESUMO GERAL* 📈\n"
    message += f"{SEPARATOR}\n"
    message += f"💼 TOTAL CHESF ➜ {format_currency(total_chesf)}\n"
    message += f"💰 TOTAL CAIÇARA - CHESF (CAIÇARA + CARTÃO) ➜ {format_currency(total_chesf_ajustado)}\n"
    message += f"➕ CHESF + TOTAL CAIÇARA ENGENHARIA ➜ {format_currency(chesf_mais_caicara)}\n"

    # ================= TOTAIS COMPLETOS POR EMPRESA =================
    total_completo_caicara = chesf_mais_caicara + cartao_caicara
    total_completo_pallas = total_pallas + cartao_pallas

    message += f"\n{SEPARATOR}\n"
    message += "🧮 *TOTAIS COMPLETOS POR EMPRESA* 🧮\n"
    message += f"{SEPARATOR}\n"
    message += f"➕ CAIÇARA (OBRAS + CHESF + CARTÃO): {format_currency(total_completo_caicara)}\n"
    message += f"➕ PALLAS (OBRAS + CARTÃO): {format_currency(total_completo_pallas)}\n"

    # ================= TOP 5 OBRAS =================
    message += f"\n{SEPARATOR}\n"
    message += "🏗️ *TOP 5 OBRAS POR VALOR* 🏆\n"
    message += f"{SEPARATOR}\n"
    top_obras = df_grouped.sort_values('Valor Previsto', ascending=False).head(5)
    for i, (_, row) in enumerate(top_obras.iterrows(), 1):
        message += (
            f"\n{i}️⃣ *{row['Obra.Projeto']}*\n"
            f"   🏢 Empresa: {row['Empresa']}\n"
            f"   💰 Valor Previsto: {format_currency(row['Valor Previsto'])}\n"
        )

    # ================= DESPESAS DE CARTÃO DETALHADAS =================
    message += f"\n\n{SEPARATOR}\n"
    message += "💳 *DESPESAS DE CARTÃO DE CRÉDITO* 💳\n"
    message += f"{SEPARATOR}\n"
    for empresa in df_cartoes['Empresa'].unique():
        message += f"\n🏢 {empresa.upper()}\n"
        sub = df_cartoes[df_cartoes['Empresa'] == empresa]
        for _, row in sub.iterrows():
            message += f"   • {row['Obra']} ➜ {format_currency(row['Bruto'])}\n"

    # ================= FOOTER =================
    message += f"\n\n{SEPARATOR}\n"
    message += "📤 *Anexando arquivo completo...*"

    # ================= SEND =================
    if send_telegram_message(bot_token, main_chat_id, message):
        send_telegram_file(bot_token, main_chat_id, output_file_path, "Planilha completa de pedidos e contas a pagar")
    else:
        print("[Relatório Principal] Falha ao enviar relatório principal.")


   

    
# =============================
# Alerta de pendentes (Telegram)
# =============================

def send_rejected_orders_alert(df_orders, bot_token, alert_chat_id, alert_chat_id_secundario):
    """Envia alerta de ordens não aprovadas para os dois canais, com aviso mesmo se vazio."""
    timestamp = datetime.now().strftime('%d/%m/%Y %H:%M')

    # Filtrar apenas ordens não aprovadas
    rejected_orders = df_orders[df_orders['Situação'].str.contains('Em Aprovação|Reprovado|Pendente', case=False, na=False)].copy()

    # Separar por valor
    high_value = rejected_orders[rejected_orders['Valor'] > 20001]
    low_value = rejected_orders[rejected_orders['Valor'] <= 20001]

    def build_message(df, faixa_texto):
        if df.empty:
            return f"⚠️ *{faixa_texto}*\n\n📅 {timestamp}\n\n⚠️ Por agora não temos ordens para aprovar nessa faixa."
        
        msg = f"⚠️ *{faixa_texto}* ⚠️\n\n📅 {timestamp}\n\n"
        msg += f"🔴 *Total de itens*: {len(df)}\n\n"
        for _, row in df.head(10).iterrows():
            msg += (
                f"• *Obra*: {row['Obra.Projeto']}\n"
                f"  *Fornecedor*: {row['Fornecedor']}\n"
                f"  *Valor*: R$ {row['Valor']:,.2f}\n"
                f"  *Ordem de Compra*: {int(row['Ordem de Compra (Nº)'])}\n"
                f"  *Situação*: {row['Situação']}\n\n"
            )
        if len(df) > 10:
            msg += f"⚠ *Mais {len(df) - 10} itens não listados*"
        return msg

    # Sempre envia algo em ambos os canais
    msg_high = build_message(high_value, "ORDENS ACIMA DE 20 MIL")
    send_telegram_message(bot_token, alert_chat_id, msg_high)
    print("[Alerta] Mensagem enviada para canal principal.")

    msg_low = build_message(low_value, "ORDENS ATÉ 20 MIL")
    send_telegram_message(bot_token, alert_chat_id_secundario, msg_low)
    print("[Alerta] Mensagem enviada para canal secundário.")

# =============================
# Função principal
# =============================

def export_full_report(orders_csv_path, finance_csv_path, output_excel_path):
    print("[Início] Processando relatórios...")

    # Configuração Telegram
    BOT_TOKEN = '8019760498:AAFc9Ro215yCe2c4hFcOJX3UXLJwPFGLNmU'
    MAIN_REPORT_CHAT_ID = '-4972134508'
    ALERT_CHAT_ID = '-1002708991795'
    ALERT_CHAT_ID_SECUNDARIO = '-4892707841'  # <= Substitua pelo ID real

    # -- Ler Ordens de Compra
    df_orders = pd.read_csv(orders_csv_path, encoding='latin-1', sep=';', skiprows=6)
    df_orders.rename(columns={'Obra': 'Obra.Projeto', 'Núm': 'Ordem de Compra (Nº)', 'Total (R$)': 'Valor', 'Desc.': 'Referência'}, inplace=True)
    df_orders['Valor'] = df_orders['Valor'].astype(str).str.replace('.', '', regex=False).str.replace(',', '.', regex=False)
    df_orders['Valor'] = pd.to_numeric(df_orders['Valor'], errors='coerce').fillna(0)
    df_orders['OBS'] = ''
    df_orders['RETORNO AO SOLICITANTE'] = ''

    df_orders_full = df_orders[['Empresa', 'Obra.Projeto', 'Valor', 'Ordem de Compra (Nº)', 'Fornecedor', 'Referência', 'Situação', 'OBS', 'RETORNO AO SOLICITANTE']].copy()
    df_orders_grouped = df_orders_full.groupby(['Empresa', 'Obra.Projeto'], as_index=False).agg({'Valor': 'sum'}).rename(columns={'Valor': 'Valor Previsto'})
    df_orders_totals = df_orders_grouped.groupby('Empresa', as_index=False).agg({'Valor Previsto': 'sum'}).rename(columns={'Valor Previsto': 'Total Empresa'})
    total_geral = df_orders_totals['Total Empresa'].sum()
    chesf_obras = [
        'PE - MANUTENCAO ELETROBRAS',
        'CENTRAL ELETROBRAS',
        'PB - MANUTENCAO ELETROBRAS',
        'RN - MANUTENCAO ELETROBRAS',
        'RN - MANUTENCAO ELETROBRAS,PB - MANUTENCAO ELETROBRAS,PE - MANUTENCAO ELETROBRAS'
    ]
    mask_chesf = df_orders_grouped['Obra.Projeto'].isin(chesf_obras)
    total_chesf = df_orders_grouped.loc[mask_chesf, 'Valor Previsto'].sum()

    # -- Ler Documentos Financeiros
    df_fin = pd.read_csv(finance_csv_path, encoding='latin-1', sep=';', skiprows=6)
    df_fin['Bruto (R$)'] = df_fin['Bruto (R$)'].astype(str).str.replace('.', '', regex=False).str.replace(',', '.', regex=False)
    df_fin['Líquido (R$)'] = df_fin['Líquido (R$)'].astype(str).str.replace('.', '', regex=False).str.replace(',', '.', regex=False)
    df_fin['Bruto'] = pd.to_numeric(df_fin['Bruto (R$)'], errors='coerce').fillna(0)
    df_fin['Liquido'] = pd.to_numeric(df_fin['Líquido (R$)'], errors='coerce').fillna(0)

    # -- Gerar Excel
    if os.path.exists(output_excel_path):
        try:
            os.remove(output_excel_path)
            print(f"[Arquivo] Removido existente: {output_excel_path}")
        except PermissionError:
            print(f"⚠ Não foi possível remover: {output_excel_path} (arquivo aberto)")
            return

    with pd.ExcelWriter(output_excel_path, engine='xlsxwriter') as writer:
        df_orders_full.to_excel(writer, sheet_name='Dados', index=False)
        df_orders_grouped.to_excel(writer, sheet_name='Por Obra', index=False)
        df_orders_totals.to_excel(writer, sheet_name='Totais por Empresa', index=False)
        df_fin.to_excel(writer, sheet_name='Documentos a Pagar', index=False)
    print(f"✔ Planilha gerada: {output_excel_path}")

    # -- Enviar para Telegram
    send_main_report(df_orders_grouped, df_orders_totals, total_geral, total_chesf,
                     df_fin, output_excel_path, BOT_TOKEN, MAIN_REPORT_CHAT_ID)
    send_rejected_orders_alert(df_orders_full, BOT_TOKEN, ALERT_CHAT_ID, ALERT_CHAT_ID_SECUNDARIO)


if __name__ == '__main__':
    base_dir = os.path.dirname(os.path.abspath(__file__))
    orders_csv = os.path.join(base_dir, 'ordens_de_compra.csv')
    finance_csv = os.path.join(base_dir, 'documentos_financeiro.csv')
    output_file = os.path.join(base_dir, 'Planilha_Completa_Relatorio.xlsx')
    export_full_report(orders_csv, finance_csv, output_file)
