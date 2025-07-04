import pandas as pd
import os
import requests
from datetime import datetime


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


def send_main_report(df_grouped, df_totals, total_geral, total_chesf, output_file_path, bot_token, main_chat_id):
    """Envia relatório principal para o grupo"""
    print("[Relatório Principal] Preparando mensagem...")
    def format_currency(value):
        return f"R$ {value:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')

    timestamp = datetime.now().strftime('%d/%m/%Y %H:%M')
    message = "📊 *RELATÓRIO DE ORDENS DE COMPRA* 📊\n\n"
    message += f"⏱ *Data/Hora*: {timestamp}\n"
    message += f"📂 *Arquivo*: `{os.path.basename(output_file_path)}`\n\n"

    message += "🏢 *TOTAIS POR EMPRESA* 💰\n"
    for _, row in df_totals.iterrows():
        message += f"• *{row['Empresa']}*: {format_currency(row['Total Empresa'])}\n"

    # Totais de CHESF e ajuste de total geral
    message += f"\n💼 *TOTAL CHESF*: *{format_currency(total_chesf)}*\n"
    message += f"💰 *TOTAL GERAL AJUSTADO*: *{format_currency(total_geral - total_chesf)}* 💰\n\n"

    message += "🏗 *TOP 5 OBRAS* 🏆\n"
    top_obras = df_grouped.sort_values('Valor Previsto', ascending=False).head(5)
    for i, (_, row) in enumerate(top_obras.iterrows(), 1):
        message += f"{i}. *{row['Obra.Projeto']}* ({row['Empresa']}): {format_currency(row['Valor Previsto'])}\n"

    message += "\n📤 *Anexando arquivo completo...*"
    if send_telegram_message(bot_token, main_chat_id, message):
        send_telegram_file(bot_token, main_chat_id, output_file_path, "Planilha completa de pedidos")
    else:
        print("[Relatório Principal] Falha ao enviar relatório principal.")


def send_rejected_orders_alert(df, bot_token, alert_chat_id):
    """Envia alerta de compras não aprovadas para o grupo de alertas"""
    print("[Alerta] Verificando compras não aprovadas...")
    rejected = df[df['Situação'].str.contains('Em Aprovação|Reprovado|Pendente', case=False, na=False)].copy()
    rejected = rejected.sort_values('Valor', ascending=False)
    if rejected.empty:
        print("[Alerta] Nenhuma compra não aprovada encontrada.")
        return

    message = "⚠️ *ALERTA DE COMPRAS NÃO APROVADAS* ⚠️\n\n"
    message += f"📅 {datetime.now().strftime('%d/%m/%Y %H:%M')}\n\n"
    message += f"🔴 *Total de itens não aprovados*: {len(rejected)}\n\n"
    for _, row in rejected.head(10).iterrows():
        message += (
            f"• *Obra*: {row['Obra.Projeto']}\n"
            f"  *Fornecedor*: {row['Fornecedor']}\n"
            f"  *Valor*: R$ {row['Valor']:,.2f}\n"
            f"  *Ordem de Compra*: {int(row['Ordem de Compra (Nº)'])}\n"  
            f"  *Situação*: {row['Situação']}\n\n"
        )
    if len(rejected) > 10:
        message += f"⚠ *Mais {len(rejected)-10} itens não aprovados não listados*"

    send_telegram_message(bot_token, alert_chat_id, message)
    print("[Alerta] Mensagem de alerta de compras não aprovadas enviada.")


def export_full_and_by_obras(csv_path, output_excel_path):
    print(f"[Início] Processando CSV: {csv_path}")
    BOT_TOKEN = '8019760498:AAFc9Ro215yCe2c4hFcOJX3UXLJwPFGLNmU'
    MAIN_REPORT_CHAT_ID = '-4972134508'
    ALERT_CHAT_ID = '-1002708991795'

    df = pd.read_csv(csv_path, encoding='latin-1', sep=';', skiprows=6)
    print(f"[Dados] Linhas lidas: {len(df)}")

    df.rename(columns={'Obra': 'Obra.Projeto', 'Núm': 'Ordem de Compra (Nº)',
                       'Total (R$)': 'Valor', 'Desc.': 'Referência'}, inplace=True)
    df['Valor'] = (df['Valor'].astype(str).str.replace('.', '', regex=False)
                                    .str.replace(',', '.', regex=False))
    df['Valor'] = pd.to_numeric(df['Valor'], errors='coerce').fillna(0)
    df['OBS'] = ''
    df['RETORNO AO SOLICITANTE'] = ''

    df_full = df[['Empresa', 'Obra.Projeto', 'Valor', 'Ordem de Compra (Nº)',
                  'Fornecedor', 'Referência', 'Situação', 'OBS', 'RETORNO AO SOLICITANTE']].copy()
    df_grouped = (df_full.groupby(['Empresa', 'Obra.Projeto'], as_index=False)
                          .agg({'Valor': 'sum'})
                          .rename(columns={'Valor': 'Valor Previsto'}))
    df_totals = (df_grouped.groupby('Empresa', as_index=False)
                           .agg({'Valor Previsto': 'sum'})
                           .rename(columns={'Valor Previsto': 'Total Empresa'}))
    total_geral = df_totals['Total Empresa'].sum()

        # Calcular total CHESF (obras específicas e agrupamentos múltiplos)
    chesf_obras = [
        'PE - MANUTENCAO ELETROBRAS',
       
        'PB - MANUTENCAO ELETROBRAS',
        'RN - MANUTENCAO ELETROBRAS',
        'RN - MANUTENCAO ELETROBRAS,PB - MANUTENCAO ELETROBRAS,PE - MANUTENCAO ELETROBRAS'
    ]
    mask_chesf = df_grouped['Obra.Projeto'].isin(chesf_obras)
    total_chesf = df_grouped.loc[mask_chesf, 'Valor Previsto'].sum()

    # Preparar linhas de total geral e CHESF por empresa e CHESF por empresa
    df_total_rows = pd.DataFrame([{
        'Empresa': row['Empresa'], 'Obra.Projeto': 'TOTAL GERAL',
        'Valor Previsto': row['Total Empresa']
    } for _, row in df_totals.iterrows()])
    # Adicionar linha TOTAL CHESF
    df_total_rows = pd.concat([
        df_total_rows,
        pd.DataFrame([{
            'Empresa': '', 'Obra.Projeto': 'TOTAL CHESF',
            'Valor Previsto': total_chesf
        }])
    ], ignore_index=True)

    # Remover arquivo existente
    if os.path.exists(output_excel_path):
        try:
            os.remove(output_excel_path)
            print(f"[Arquivo] Removido existente: {output_excel_path}")
        except PermissionError:
            print(f"⚠ Não foi possível remover: {output_excel_path} (arquivo aberto)")
            return

    # Exportar para Excel
    with pd.ExcelWriter(output_excel_path, engine='xlsxwriter') as writer:
        df_full.to_excel(writer, sheet_name='Dados', index=False)
        df_grouped.to_excel(writer, sheet_name='Por Obra', index=False)
        df_total_rows.to_excel(writer, sheet_name='Totais por Empresa', index=False)
    print(f"✔ Planilha gerada: {output_excel_path}")

    send_main_report(df_grouped, df_totals, total_geral, total_chesf,
                     output_excel_path, BOT_TOKEN, MAIN_REPORT_CHAT_ID)
    send_rejected_orders_alert(df_full, BOT_TOKEN, ALERT_CHAT_ID)


if __name__ == '__main__':
    base_dir = os.path.dirname(os.path.abspath(__file__))
    csv_file = os.path.join(base_dir, 'ordens_de_compra.csv')
    output_file = os.path.join(base_dir, 'Planilha_Pedidos_Completos.xlsx')
    export_full_and_by_obras(csv_file, output_file)
