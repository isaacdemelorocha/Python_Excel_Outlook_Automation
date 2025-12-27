import pandas as pd
import warnings
from pathlib import Path
from appscript import app, k, mactypes

# Silenciar avisos de formatação do Excel
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")


class AnalisadorAderenciaMac:
    """
    Consolida dados de treinamentos e automatiza o envio via Microsoft Outlook (macOS).
    Interface e dados 100% traduzidos para Português.
    """

    def __init__(self, caminho_arquivo: str):
        # --- CENTRALIZAÇÃO DE VARIÁVEIS E MAPEAMENTO ---
        self.caminho = Path(caminho_arquivo)
        self.df_resultado = None

        # Mapeamento de Colunas Originais do Excel
        self.COL_MISSAO = 'Mission Definitions Name'
        self.COL_STATUS = 'Mission Enrollments Status'

        # Dicionário de Tradução dos Status
        self.TRADUCAO_STATUS = {
            'COMPLETED': 'Concluído',
            'IN_PROGRESS': 'Em Andamento',
            'NOT_STARTED': 'Não Iniciado'
        }

    def processar(self) -> pd.DataFrame:
        """Lê o Excel, traduz os termos técnicos e calcula os KPIs."""
        if not self.caminho.exists():
            raise FileNotFoundError(f"Arquivo não localizado: {self.caminho}")

        # Leitura do arquivo
        df = pd.read_excel(self.caminho)

        # 1. Agrupamento matricial (Pivot Table)
        resumo = (
            df.groupby([self.COL_MISSAO, self.COL_STATUS])
            .size()
            .unstack(fill_value=0)
        )

        # 2. TRADUÇÃO DOS CABEÇALHOS
        # Traduz as colunas (Status)
        resumo = resumo.rename(columns=self.TRADUCAO_STATUS)
        # Traduz o título do agrupamento de colunas
        resumo.columns.name = 'Status de Conclusão'
        # Traduz o título das linhas
        resumo.index.name = 'Nome do Curso'

        # 3. CÁLCULO DA TAXA DE ADERÊNCIA
        total_por_missao = resumo.sum(axis=1)
        col_concluido = self.TRADUCAO_STATUS['COMPLETED']

        resumo['Aderência (%)'] = (
                resumo.get(col_concluido, 0) / total_por_missao * 100
        ).round(2)

        # Ordenação: Do maior percentual para o menor
        self.df_resultado = resumo.sort_values(by='Aderência (%)', ascending=False)
        return self.df_resultado

    def gerar_corpo_html(self) -> str:
        """Gera o template HTML/CSS moderno com experiência de utilizador (UX) premium."""
        tabela_html = self.df_resultado.to_html(classes='kpi-table', border=0)

        return f"""
        <html>
        <head>
            <style>
                .email-body {{ background-color: #f0f2f5; padding: 20px; font-family: 'Segoe UI', Helvetica, Arial, sans-serif; }}
                .card {{ max-width: 650px; margin: 0 auto; background: #ffffff; border-radius: 12px; border: 1px solid #e1e4e8; overflow: hidden; box-shadow: 0 4px 12px rgba(0,0,0,0.05); }}
                .brand-bar {{ background: #0078d4; height: 6px; }}
                .header {{ padding: 30px; text-align: center; color: #1a1a1a; }}
                .content {{ padding: 0 30px 30px 30px; color: #444; line-height: 1.6; }}
                .kpi-table {{ width: 100%; border-collapse: collapse; margin-top: 20px; }}
                .kpi-table th {{ background-color: #f8f9fa; color: #0078d4; text-align: left; padding: 12px; border-bottom: 2px solid #0078d4; font-size: 11px; text-transform: uppercase; }}
                .kpi-table td {{ padding: 10px 12px; border-bottom: 1px solid #eee; font-size: 13px; }}
                .footer {{ background: #f8f9fa; padding: 15px; text-align: center; font-size: 11px; color: #888; border-top: 1px solid #eee; }}
                .highlight {{ font-weight: bold; color: #0078d4; }}
            </style>
        </head>
        <body class="email-body">
            <div class="card">
                <div class="brand-bar"></div>
                <div class="header">
                    <h2 style="margin:0;">Dashboard de Aderência dos Cursos</h2>
                    <p style="margin:5px 0 0; color: #666; font-size: 14px;">Monitoramento dos Cursos</p>
                </div>
                <div class="content">
                    <p>Olá,</p>
                    <p>Os indicadores dos <span class="highlight">Cursos</span> foram atualizados. Abaixo, encontra o resumo consolidado por matéria:</p>

                    <div style="overflow-x:auto;">
                        {tabela_html}
                    </div>

                    <p style="margin-top:25px; font-size: 12px; color: #777;">
                        * O ficheiro detalhado original foi anexado para consulta de dados individuais.
                    </p>
                </div>
                <div class="footer">
                    <p>Automated Reporting System | <strong>2025</strong></p>
                </div>
            </div>
        </body>
        </html>
        """

    def enviar_email(self, destinatario: str):
        """Dispara o e-mail através do Microsoft Outlook do macOS."""
        if self.df_resultado is None:
            raise ValueError("Erro: Processe os dados antes de tentar enviar o e-mail.")

        assunto = f"📊 KPI: Aderência dos Cursos - {self.caminho.stem}"
        corpo = self.gerar_corpo_html()

        # Interface com o Microsoft Outlook
        outlook = app('Microsoft Outlook')

        # 1. Cria a nova mensagem
        msg = outlook.make(
            new=k.outgoing_message,
            with_properties={k.subject: assunto, k.content: corpo}
        )

        # 2. Define o destinatário
        msg.make(new=k.recipient, with_properties={k.email_address: {k.address: destinatario}})

        # 3. Anexa o ficheiro (Uso de Alias para evitar erros de permissão no Mac)
        caminho_abs = str(self.caminho.absolute())
        arquivo_alias = mactypes.Alias(caminho_abs)
        msg.make(new=k.attachment, with_properties={k.file: arquivo_alias})

        # 4. Enviar
        msg.send()


# --- BLOCO PRINCIPAL DE EXECUÇÃO ---
if __name__ == "__main__":
    # Configurações de Entrada
    DADOS_CONFIG = {
        "CAMINHO_EXCEL": "Caminho/Para/Seu/Arquivo.xlsx",
        "DESTINO": "emaildedestino@exemplo.com"
    }

    try:
        # Instanciar e executar
        servico = AnalisadorAderenciaMac(DADOS_CONFIG["CAMINHO_EXCEL"])

        print(f"[*] A processar ficheiro: {Path(DADOS_CONFIG['CAMINHO_EXCEL']).name}")
        servico.processar()

        print(f"[*] A gerar interface e a enviar via Outlook para {DADOS_CONFIG['DESTINO']}...")
        servico.enviar_email(DADOS_CONFIG["DESTINO"])

        print("\n[SUCESSO] Relatório traduzido e enviado com sucesso!")

    except Exception as e:
        print(f"\n[ERRO CRÍTICO]: {e}")
