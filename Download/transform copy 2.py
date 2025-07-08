import pandas as pd
import os
import requests
from datetime import datetime, timedelta
import plotly.express as px
import jinja2
import subprocess
from pathlib import Path
import time

# =============================
# Configurações Globais
# =============================
class Config:
    # Telegram
    BOT_TOKEN = '8019760498:AAFc9Ro215yCe2c4hFcOJX3UXLJwPFGLNmU'
    MAIN_REPORT_CHAT_ID = '-4813455628'
    
    # Netlify
    NETLIFY_SITE_NAME = "meu-dashboard-financeiro"
    NETLIFY_DIR = "netlify"
    NETLIFY_DEPLOY_CMD = "netlify deploy --prod --json"  # Adicionado --json para melhor parsing
    
    # Templates
    TEMPLATE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "templates")
    HTML_TEMPLATE = "dashboard_template.html"
    
    # Validade do link (horas)
    LINK_VALIDITY_HOURS = 24

# =============================
# Helpers
# =============================
class Helpers:
    @staticmethod
    def format_currency(value):
        """Formata valores monetários"""
        if pd.isna(value):
            return "R$ 0,00"
        return f"R$ {value:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
    
    @staticmethod
    def generate_filename(prefix="dashboard"):
        """Gera um nome de arquivo único com timestamp"""
        timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
        return f"{prefix}-{timestamp}.html"
    
    @staticmethod
    def create_directory(path):
        """Cria diretório se não existir"""
        Path(path).mkdir(parents=True, exist_ok=True)
    
    @staticmethod
    def debug_paths():
        """Debug: mostra os caminhos importantes"""
        print("\n[DEBUG] Caminhos importantes:")
        print(f"Diretório do script: {os.path.dirname(os.path.abspath(__file__))}")
        print(f"Pasta templates: {Config.TEMPLATE_DIR}")
        print(f"Template existe?: {os.path.exists(os.path.join(Config.TEMPLATE_DIR, Config.HTML_TEMPLATE))}")
        print(f"Conteúdo da pasta templates: {os.listdir(Config.TEMPLATE_DIR) if os.path.exists(Config.TEMPLATE_DIR) else 'Pasta não existe'}\n")

# =============================
# Processamento de Dados
# =============================
class DataProcessor:
    def __init__(self, orders_csv_path, finance_csv_path):
        self.orders_csv_path = orders_csv_path
        self.finance_csv_path = finance_csv_path
        self.df_orders = None
        self.df_finance = None
        self.process_data()
    
    def process_data(self):
        """Processa os CSVs e prepara os DataFrames"""
        # Processar ordens de compra
        self.df_orders = pd.read_csv(self.orders_csv_path, encoding='latin-1', sep=';', skiprows=6)
        self._clean_orders_data()
        
        # Processar despesas de cartão
        self.df_finance = pd.read_csv(self.finance_csv_path, encoding='latin-1', sep=';', skiprows=6)
        self._clean_finance_data()
    
    def _clean_orders_data(self):
        """Limpa e formata os dados de ordens de compra"""
        # Renomear colunas
        self.df_orders.rename(columns={
            'Obra': 'Obra.Projeto', 
            'Núm': 'Ordem de Compra (Nº)', 
            'Total (R$)': 'Valor', 
            'Desc.': 'Referência'
        }, inplace=True)
        
        # Converter valores
        self.df_orders['Valor'] = (
            self.df_orders['Valor']
            .astype(str)
            .str.replace('.', '', regex=False)
            .str.replace(',', '.', regex=False)
        )
        self.df_orders['Valor'] = pd.to_numeric(self.df_orders['Valor'], errors='coerce').fillna(0)
        
        # Adicionar colunas extras
        self.df_orders['OBS'] = ''
        self.df_orders['RETORNO AO SOLICITANTE'] = ''
    
    def _clean_finance_data(self):
        """Limpa e formata os dados financeiros"""
        # Converter valores
        for col in ['Bruto (R$)', 'Líquido (R$)']:
            self.df_finance[col] = (
                self.df_finance[col]
                .astype(str)
                .str.replace('.', '', regex=False)
                .str.replace(',', '.', regex=False)
            )
            self.df_finance[col.replace(' (R$)', '')] = pd.to_numeric(self.df_finance[col], errors='coerce').fillna(0)
    
    def get_orders_by_company(self):
        """Agrupa ordens por empresa"""
        return (
            self.df_orders
            .groupby(['Empresa', 'Obra.Projeto'], as_index=False)
            .agg({'Valor': 'sum'})
            .rename(columns={'Valor': 'Valor Previsto'})
        )
    
    def get_totals_by_company(self):
        """Calcula totais por empresa"""
        df_grouped = self.get_orders_by_company()
        return (
            df_grouped
            .groupby('Empresa', as_index=False)
            .agg({'Valor Previsto': 'sum'})
            .rename(columns={'Valor Previsto': 'Total Empresa'})
        )
    
    def get_top_projects(self, n=5):
        """Retorna os top N projetos por valor"""
        df_grouped = self.get_orders_by_company()
        return (
            df_grouped
            .sort_values('Valor Previsto', ascending=False)
            .head(n)
        )
    
    def get_card_expenses(self):
        """Agrupa despesas de cartão por empresa"""
        return (
            self.df_finance
            .groupby(['Empresa', 'Obra'], as_index=False)
            .agg({'Bruto': 'sum'})
        )

# =============================
# Geração de Gráficos
# =============================
class ChartGenerator:
    @staticmethod
    def create_company_totals_chart(df_totals):
        """Cria gráfico de barras com totais por empresa"""
        fig = px.bar(
            df_totals,
            x='Empresa',
            y='Total Empresa',
            title='Totais por Empresa',
            color='Empresa',
            text_auto='.2s',
            labels={'Total Empresa': 'Valor Total (R$)'}
        )
        fig.update_traces(textfont_size=12, textangle=0, textposition="outside", cliponaxis=False)
        return fig.to_html(full_html=False)
    
    @staticmethod
    def create_top_projects_chart(df_top_projects):
        """Cria gráfico de barras dos top projetos"""
        fig = px.bar(
            df_top_projects,
            x='Obra.Projeto',
            y='Valor Previsto',
            title='Top 5 Obras por Valor',
            color='Empresa',
            text_auto='.2s',
            labels={'Valor Previsto': 'Valor (R$)'}
        )
        fig.update_layout(xaxis={'categoryorder': 'total descending'})
        return fig.to_html(full_html=False)
    
    @staticmethod
    def create_card_expenses_chart(df_card_expenses):
        """Cria gráfico de pizza com despesas de cartão por empresa"""
        fig = px.pie(
            df_card_expenses,
            names='Empresa',
            values='Bruto',
            title='Despesas de Cartão por Empresa',
            hole=0.3
        )
        fig.update_traces(textposition='inside', textinfo='percent+label')
        return fig.to_html(full_html=False)

# =============================
# Geração do HTML
# =============================
class HTMLGenerator:
    def __init__(self):
        # Debug: verificar caminhos antes de carregar
        Helpers.debug_paths()
        
        try:
            self.template_loader = jinja2.FileSystemLoader(searchpath=Config.TEMPLATE_DIR)
            self.template_env = jinja2.Environment(loader=self.template_loader)
            print("Template environment criado com sucesso!")
        except Exception as e:
            print(f"Erro ao criar template environment: {e}")
            raise
    
    def generate_dashboard(self, data_processor, output_path):
        """Gera o dashboard HTML completo"""
        try:
            # Preparar dados
            df_totals = data_processor.get_totals_by_company()
            df_top_projects = data_processor.get_top_projects()
            df_card_expenses = data_processor.get_card_expenses()
            
            # Gerar gráficos
            company_totals_chart = ChartGenerator.create_company_totals_chart(df_totals)
            top_projects_chart = ChartGenerator.create_top_projects_chart(df_top_projects)
            card_expenses_chart = ChartGenerator.create_card_expenses_chart(df_card_expenses)
            
            # Preparar contexto para o template
            context = {
                'report_date': datetime.now().strftime('%d/%m/%Y %H:%M'),
                'company_totals': df_totals.to_dict('records'),
                'top_projects': df_top_projects.to_dict('records'),
                'card_expenses': df_card_expenses.to_dict('records'),
                'company_totals_chart': company_totals_chart,
                'top_projects_chart': top_projects_chart,
                'card_expenses_chart': card_expenses_chart,
                'format_currency': Helpers.format_currency
            }
            
            # Carregar template e renderizar
            template_path = os.path.join(Config.TEMPLATE_DIR, Config.HTML_TEMPLATE)
            print(f"Tentando carregar template de: {template_path}")
            
            template = self.template_env.get_template(Config.HTML_TEMPLATE)
            html_output = template.render(context)
            
            # Salvar arquivo HTML
            Helpers.create_directory(os.path.dirname(output_path))
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(html_output)
            
            print(f"Dashboard gerado com sucesso em: {output_path}")
            return output_path
            
        except Exception as e:
            print(f"Erro ao gerar dashboard: {e}")
            raise

# =============================
# Integração com Netlify - Versão Aprimorada
# =============================
class NetlifyDeployer:
    @staticmethod
    def check_netlify_cli():
        """Verifica se o Netlify CLI está instalado"""
        try:
            result = subprocess.run(
                "netlify --version",
                shell=True,
                capture_output=True,
                text=True
            )
            if result.returncode == 0:
                print(f"Netlify CLI encontrado. Versão: {result.stdout.strip()}")
                return True
            print("Netlify CLI não encontrado ou não configurado corretamente.")
            return False
        except Exception as e:
            print(f"Erro ao verificar Netlify CLI: {e}")
            return False
    
    @staticmethod
    def deploy_to_netlify(html_file_path):
        """Versão robusta para deploy no Netlify"""
        try:
            print("\n[NETLIFY] Preparando deploy...")
            
            # 1. Verificar CLI
            if not NetlifyDeployer.check_netlify_cli():
                print("Instale o Netlify CLI com: npm install -g netlify-cli")
                return None
            
            # 2. Preparar pasta netlify
            netlify_dir = os.path.join(os.path.dirname(html_file_path), Config.NETLIFY_DIR)
            Helpers.create_directory(netlify_dir)
            
            # 3. Copiar arquivo para a pasta netlify
            dest_path = os.path.join(netlify_dir, os.path.basename(html_file_path))
            
            if os.path.exists(dest_path):
                os.remove(dest_path)
            os.rename(html_file_path, dest_path)
            
            print(f"[NETLIFY] Arquivo copiado para: {dest_path}")
            
            # 4. Executar deploy com timeout
            print("[NETLIFY] Iniciando deploy (pode levar alguns minutos)...")
            start_time = time.time()
            
            result = subprocess.run(
                Config.NETLIFY_DEPLOY_CMD,
                cwd=netlify_dir,
                shell=True,
                capture_output=True,
                text=True,
                timeout=300  # 5 minutos de timeout
            )
            
            print(f"[NETLIFY] Deploy concluído em {time.time() - start_time:.2f} segundos")
            
            # Processar saída JSON
            if result.returncode == 0:
                try:
                    import json
                    output = json.loads(result.stdout)
                    if 'deploy_url' in output:
                        url = output['deploy_url']
                        print(f"[NETLIFY] Dashboard publicado em: {url}")
                        return url
                except json.JSONDecodeError:
                    # Fallback para parsing manual se JSON falhar
                    for line in result.stdout.split('\n'):
                        if 'Website URL:' in line:
                            url = line.split('Website URL:')[1].strip()
                            print(f"[NETLIFY] Dashboard publicado em: {url}")
                            return url
            
            print("[NETLIFY] Falha ao obter URL do deploy")
            print(f"Saída do comando:\n{result.stdout}")
            if result.stderr:
                print(f"Erros:\n{result.stderr}")
            return None
            
        except subprocess.TimeoutExpired:
            print("[NETLIFY] Erro: Timeout excedido (5 minutos)")
            return None
        except Exception as e:
            print(f"[NETLIFY] Erro inesperado: {e}")
            return None

# =============================
# Integração com Telegram
# =============================
class TelegramNotifier:
    @staticmethod
    def send_dashboard_link(chat_id, dashboard_url):
        """Envia mensagem com link do dashboard para o Telegram"""
        expiration = (datetime.now() + timedelta(hours=Config.LINK_VALIDITY_HOURS)).strftime('%d/%m/%Y %H:%M')
        
        message = (
            f"📊 *Seu Relatório Online está pronto!*\n\n"
            f"🌐 [Clique aqui para acessar o Dashboard]({dashboard_url})\n\n"
            f"⏱️ *Link válido até*: {expiration}"
        )
        
        url = f"https://api.telegram.org/bot{Config.BOT_TOKEN}/sendMessage"
        params = {
            'chat_id': chat_id,
            'text': message,
            'parse_mode': 'Markdown',
            'disable_web_page_preview': False
        }
        
        try:
            response = requests.post(url, params=params)
            response.raise_for_status()
            return True
        except Exception as e:
            print(f"Erro ao enviar mensagem para o Telegram: {e}")
            return False

# =============================
# Função Principal
# =============================
def main():
    print("=== Iniciando geração do dashboard ===")
    
    try:
        # 1. Configurar caminhos dos arquivos
        base_dir = os.path.dirname(os.path.abspath(__file__))
        orders_csv = os.path.join(base_dir, 'ordens_de_compra.csv')
        finance_csv = os.path.join(base_dir, 'documentos_financeiro.csv')
        output_html = os.path.join(base_dir, Helpers.generate_filename())
        
        print(f"Diretório base: {base_dir}")
        print(f"Arquivo de ordens: {orders_csv}")
        print(f"Arquivo financeiro: {finance_csv}")
        
        # 2. Processar dados
        print("\n[1/4] Processando dados...")
        data_processor = DataProcessor(orders_csv, finance_csv)
        
        # 3. Gerar HTML
        print("\n[2/4] Gerando dashboard HTML...")
        html_generator = HTMLGenerator()
        html_path = html_generator.generate_dashboard(data_processor, output_html)
        
        # 4. Fazer deploy no Netlify
        print("\n[3/4] Fazendo deploy no Netlify...")
        dashboard_url = NetlifyDeployer.deploy_to_netlify(html_path)
        
        if dashboard_url:
            print(f"\n[SUCESSO] Dashboard publicado em: {dashboard_url}")
            
            # 5. Enviar link via Telegram
            print("\n[4/4] Enviando notificação para o Telegram...")
            if TelegramNotifier.send_dashboard_link(Config.MAIN_REPORT_CHAT_ID, dashboard_url):
                print("Notificação enviada com sucesso!")
            else:
                print("Falha ao enviar notificação")
        else:
            print("\n[ERRO] Falha ao publicar no Netlify")
            print("Verifique se:")
            print("1. O Netlify CLI está instalado (npm install -g netlify-cli)")
            print("2. Você está autenticado (netlify login)")
            print("3. Sua conexão com a internet está estável")
    
    except Exception as e:
        print(f"\n[ERRO CRÍTICO] Ocorreu um erro: {e}")
        print("Detalhes do erro:")
        import traceback
        traceback.print_exc()

if __name__ == '__main__':
    main()