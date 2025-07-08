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
                     df_financeiro_totals, output_file_path,
                     bot_token, main_chat_id):
    """Envia mensagem principal para o grupo Telegram"""
    def format_currency(value):
        return f"R$ {value:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')

    timestamp = datetime.now().strftime('%d/%m/%Y %H:%M')
    message = "📊 *RELATÓRIO DE ORDENS DE COMPRA E DOCUMENTOS A PAGAR* 📊\n\n"
    message += f"⏱ *Data/Hora*: {timestamp}\n"
    message += f"📂 *Arquivo*: `{os.path.basename(output_file_path)}`\n\n"

    # Ordens de Compra - Totais
    message += "🏢 *TOTAIS ORDENS DE COMPRA POR EMPRESA* 💰\n"
    for _, row in df_totals.iterrows():
        message += f"• *{row['Empresa']}*: {format_currency(row['Total Empresa'])}\n"

    message += f"\n💼 *TOTAL CHESF*: *{format_currency(total_chesf)}*\n"
    message += f"💰 *TOTAL GERAL AJUSTADO*: *{format_currency(total_geral - total_chesf)}*\n\n"

    # Top 5 Obras
    message += "🏗 *TOP 5 OBRAS* 🏆\n"
    top_obras = df_grouped.sort_values('Valor Previsto', ascending=False).head(5)
    for i, (_, row) in enumerate(top_obras.iterrows(), 1):
        message += f"{i}. *{row['Obra.Projeto']}* ({row['Empresa']}): {format_currency(row['Valor Previsto'])}\n"

    message += "\n🗂 *DOCUMENTOS A PAGAR POR EMPRESA* 💳\n"
    for _, row in df_financeiro_totals.iterrows():
        message += f"• *{row['Empresa']}*: Bruto {format_currency(row['Bruto'])}, Líquido {format_currency(row['Liquido'])}\n"

    message += "\n📤 *Anexando arquivo completo...*"
    if send_telegram_message(bot_token, main_chat_id, message):
        send_telegram_file(bot_token, main_chat_id, output_file_path, "Planilha completa de pedidos e contas a pagar")
    else:
        print("[Relatório Principal] Falha ao enviar relatório principal.")

# =============================
# Alerta de pendentes (Telegram)
# =============================

def send_rejected_orders_alert(df_orders, df_financeiro, bot_token, alert_chat_id):
    """Envia alerta para compras e documentos não aprovados"""
    timestamp = datetime.now().strftime('%d/%m/%Y %H:%M')
    message = f"⚠️ *ALERTA DE ITENS NÃO APROVADOS/PENDENTES* ⚠️\n\n📅 {timestamp}\n\n"

    # -- Ordens de Compra
    rejected_orders = df_orders[df_orders['Situação'].str.contains('Em Aprovação|Reprovado|Pendente', case=False, na=False)]
    if not rejected_orders.empty:
        message += f"🔴 *Ordens de Compra Não Aprovadas*: {len(rejected_orders)} itens\n\n"
        for _, row in rejected_orders.head(10).iterrows():
            message += (
                f"• *Obra*: {row['Obra.Projeto']}\n"
                f"  *Fornecedor*: {row['Fornecedor']}\n"
                f"  *Valor*: R$ {row['Valor']:,.2f}\n"
                f"  *Ordem de Compra*: {int(row['Ordem de Compra (Nº)'])}\n"
                f"  *Situação*: {row['Situação']}\n\n"
            )

    # -- Documentos a Pagar
    rejected_fin = df_financeiro[df_financeiro['Situação'].str.contains('Em Aprovação|Reprovado|Pendente', case=False, na=False)]
    if not rejected_fin.empty:
        message += f"🔴 *Documentos a Pagar Não Aprovados*: {len(rejected_fin)} itens\n\n"
        for _, row in rejected_fin.head(10).iterrows():
            message += (
                f"• *Empresa*: {row['Empresa']}\n"
                f"  *Obra*: {row['Obra']}\n"
                f"  *Fornecedor*: {row['Fornecedor']}\n"
                f"  *Bruto*: R$ {row['Bruto']:.2f}\n"
                f"  *Vencimento*: {row['Venc.']}\n"
                f"  *Situação*: {row['Situação']}\n\n"
            )

    if rejected_orders.empty and rejected_fin.empty:
        print("[Alerta] Nenhuma compra ou documento não aprovado encontrado.")
        return

    send_telegram_message(bot_token, alert_chat_id, message)
    print("[Alerta] Mensagem de alerta enviada.")


# =============================
# Função principal
# =============================

def export_full_report(orders_csv_path, finance_csv_path, output_excel_path):
    print("[Início] Processando relatórios...")

    # Configuração Telegram
    BOT_TOKEN = '8019760498:AAFc9Ro215yCe2c4hFcOJX3UXLJwPFGLNmU'
    MAIN_REPORT_CHAT_ID = '-4972134508'
    ALERT_CHAT_ID = '-1002708991795'

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

    df_financeiro_totals = df_fin.groupby('Empresa', as_index=False).agg({'Bruto': 'sum', 'Liquido': 'sum'})

    # -- Remover arquivo antigo
    if os.path.exists(output_excel_path):
        try:
            os.remove(output_excel_path)
            print(f"[Arquivo] Removido existente: {output_excel_path}")
        except PermissionError:
            print(f"⚠ Não foi possível remover: {output_excel_path} (arquivo aberto)")
            return

    # -- Gerar Excel
    with pd.ExcelWriter(output_excel_path, engine='xlsxwriter') as writer:
        df_orders_full.to_excel(writer, sheet_name='Dados', index=False)
        df_orders_grouped.to_excel(writer, sheet_name='Por Obra', index=False)
        df_orders_totals.to_excel(writer, sheet_name='Totais por Empresa', index=False)
        df_fin.to_excel(writer, sheet_name='Documentos a Pagar', index=False)
    print(f"✔ Planilha gerada: {output_excel_path}")

    # -- Enviar para Telegram
    send_main_report(df_orders_grouped, df_orders_totals, total_geral, total_chesf,
                     df_financeiro_totals, output_excel_path, BOT_TOKEN, MAIN_REPORT_CHAT_ID)
    send_rejected_orders_alert(df_orders_full, df_fin, BOT_TOKEN, ALERT_CHAT_ID)


if __name__ == '__main__':
    base_dir = os.path.dirname(os.path.abspath(__file__))
    orders_csv = os.path.join(base_dir, 'ordens_de_compra.csv')
    finance_csv = os.path.join(base_dir, 'documentos_financeiro.csv')
    output_file = os.path.join(base_dir, 'Planilha_Completa_Relatorio.xlsx')
    export_full_report(orders_csv, finance_csv, output_file)
