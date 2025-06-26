import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

def calcular_balanceamento_por_obra(df, data_col, up_col, data_estrutura_col, coef_quartil_col, coef_mediana_col, coef_real_col, obra_col, data_referencia):
    """
    Calcula o balanceamento com as seguintes regras:
    1. Para meses anteriores à data_inicio_ponderacao: mantém valor original da UP
    2. Para meses a partir da data_inicio_ponderacao:
       - Se data_forecast >= data_estrutura: 3 próximos meses usam mediana/real, demais quartil
       - Se data_forecast < data_estrutura: todos usam quartil
    3. Cria coluna Data_rebalanc. estendendo/retraindo datas conforme meses a esticar
    4. Distribui a diferença de UPs uniformemente nos meses necessários
    """
    df = df.copy()

    # Converte datas para datetime
    df['data_mensal'] = pd.to_datetime(df[data_col], errors='coerce', dayfirst=True)
    df['data_estrutura'] = pd.to_datetime(df[data_estrutura_col], errors='coerce', dayfirst=True)
    data_forecast = pd.to_datetime(data_referencia, dayfirst=True)

    # Inicializa colunas de resultado
    df['data_inicio_ponderacao'] = np.nan
    df['C1: UP balanc. 1ºQ'] = df[up_col]  # Inicia com valores originais
    df['C2: UP balanc. Med.'] = df[up_col]  # Inicia com valores originais
    df['C3: 3m Med. + 1ºQ'] = df[up_col]   # Inicia com valores originais
    df['C4: 3m Real. + 1ºQ'] = df[up_col]  # Nova coluna para coeficiente real
    df['C3: tipo_ponderacao'] = "UP Original (Forecast)"
    df['C4: tipo_ponderacao'] = "UP Original (Forecast)"
    
    # Novas colunas de médias e totais
    df['C1: Média UP'] = np.nan
    df['C2: Média UP'] = np.nan
    df['C3: Média UP'] = np.nan
    df['C4: Média UP'] = np.nan
    df['C1: Total UP'] = np.nan
    df['C2: Total UP'] = np.nan
    df['C3: Total UP'] = np.nan
    df['C4: Total UP'] = np.nan
    
    # Novas colunas de diferença
    df['C1: Diferença UP'] = np.nan
    df['C2: Diferença UP'] = np.nan
    df['C3: Diferença UP'] = np.nan
    df['C4: Diferença UP'] = np.nan
    
    df['C1 - Meses a esticar'] = np.nan
    df['C2 - Meses a esticar'] = np.nan
    df['C3 - Meses a esticar'] = np.nan
    df['C4 - Meses a esticar'] = np.nan

    df['Unidade Totais'] = np.nan
    df['%AMP quartil'] = np.nan
    df['%AMP mediana'] = np.nan
    df['%AMP real'] = np.nan
    df['VP Bruta quartil'] = np.nan
    df['VP Bruta mediana'] = np.nan
    df['VP Bruta real'] = np.nan

    # Lista para armazenar os DataFrames de cada obra após rebalanceamento
    dfs_rebalanceados = []

    # Processa cada obra individualmente
    for obra, grupo in df.groupby(obra_col):
        # Armazena valores das colunas regionais
        regional_producao = grupo['Regional Produção'].iloc[0]
        abertura_regional = grupo['Abertura Regional'].iloc[0]
        
        coef_quartil = grupo[coef_quartil_col].iloc[0]
        coef_mediana = grupo[coef_mediana_col].iloc[0]
        coef_real = grupo[coef_real_col].iloc[0]
        data_estrutura_obra = grupo['data_estrutura'].iloc[0]

        # Determina data de início (maior entre forecast e estrutura)
        data_inicio = max(data_forecast, data_estrutura_obra)
        grupo['data_inicio_ponderacao'] = data_inicio

        # Ordena por data para garantir cronologia
        grupo_ordenado = grupo.sort_values('data_mensal')
        
        # Identifica meses a partir da data de início
        meses_posteriores = grupo_ordenado[grupo_ordenado['data_mensal'] >= data_inicio]
        
        if len(meses_posteriores) > 0:
            # Aplica coeficientes nos meses posteriores
            grupo.loc[meses_posteriores.index, 'C1: UP balanc. 1ºQ'] = meses_posteriores[up_col] * coef_quartil
            grupo.loc[meses_posteriores.index, 'C2: UP balanc. Med.'] = meses_posteriores[up_col] * coef_mediana
            grupo.loc[meses_posteriores.index, 'UP balanceada real'] = meses_posteriores[up_col] * coef_real
            
            # Verifica condição para aplicar mediana/real
            if data_forecast >= data_estrutura_obra:
                primeiros_meses = meses_posteriores.head(3)
                
                # Para C3 (Mediana)
                grupo.loc[primeiros_meses.index, 'C3: 3m Med. + 1ºQ'] = grupo.loc[primeiros_meses.index, 'C2: UP balanc. Med.']
                grupo.loc[primeiros_meses.index, 'C3: tipo_ponderacao'] = "Mediana (3 próximos meses)"
                
                # Para C4 (Real)
                grupo.loc[primeiros_meses.index, 'C4: 3m Real. + 1ºQ'] = grupo.loc[primeiros_meses.index, 'UP balanceada real']
                grupo.loc[primeiros_meses.index, 'C4: tipo_ponderacao'] = "Real (3 próximos meses)"
                
                meses_restantes = meses_posteriores.iloc[3:] if len(meses_posteriores) > 3 else pd.DataFrame()
                if len(meses_restantes) > 0:
                    grupo.loc[meses_restantes.index, 'C3: 3m Med. + 1ºQ'] = grupo.loc[meses_restantes.index, 'C1: UP balanc. 1ºQ']
                    grupo.loc[meses_restantes.index, 'C3: tipo_ponderacao'] = "Quartil (meses restantes)"
                    
                    grupo.loc[meses_restantes.index, 'C4: 3m Real. + 1ºQ'] = grupo.loc[meses_restantes.index, 'C1: UP balanc. 1ºQ']
                    grupo.loc[meses_restantes.index, 'C4: tipo_ponderacao'] = "Quartil (meses restantes)"
            else:
                grupo.loc[meses_posteriores.index, 'C3: 3m Med. + 1ºQ'] = grupo.loc[meses_posteriores.index, 'C1: UP balanc. 1ºQ']
                grupo.loc[meses_posteriores.index, 'C3: tipo_ponderacao'] = "Quartil (data_forecast < estrutura)"
                
                grupo.loc[meses_posteriores.index, 'C4: 3m Real. + 1ºQ'] = grupo.loc[meses_posteriores.index, 'C1: UP balanc. 1ºQ']
                grupo.loc[meses_posteriores.index, 'C4: tipo_ponderacao'] = "Quartil (data_forecast < estrutura)"

        # Calcula métricas agregadas
        total_up_original = grupo[up_col].sum()
        total_up_c1 = grupo['C1: UP balanc. 1ºQ'].sum()
        total_up_c2 = grupo['C2: UP balanc. Med.'].sum()
        total_up_c3 = grupo['C3: 3m Med. + 1ºQ'].sum()
        total_up_c4 = grupo['C4: 3m Real. + 1ºQ'].sum()
        
        # Calcula médias, ignorando valores 0
        media_c1 = grupo.loc[grupo['C1: UP balanc. 1ºQ'] != 0, 'C1: UP balanc. 1ºQ'].mean()
        media_c2 = grupo.loc[grupo['C2: UP balanc. Med.'] != 0, 'C2: UP balanc. Med.'].mean()
        media_c3 = grupo.loc[grupo['C3: 3m Med. + 1ºQ'] != 0, 'C3: 3m Med. + 1ºQ'].mean()
        media_c4 = grupo.loc[grupo['C4: 3m Real. + 1ºQ'] != 0, 'C4: 3m Real. + 1ºQ'].mean()
        
        # Atribui totais e médias (valores fixos para cada obra)
        grupo['C1: Total UP'] = total_up_original
        grupo['C2: Total UP'] = total_up_original
        grupo['C3: Total UP'] = total_up_original
        grupo['C4: Total UP'] = total_up_original
        
        grupo['C1: Média UP'] = media_c1 if not pd.isna(media_c1) else 0
        grupo['C2: Média UP'] = media_c2 if not pd.isna(media_c2) else 0
        grupo['C3: Média UP'] = media_c3 if not pd.isna(media_c3) else 0
        grupo['C4: Média UP'] = media_c4 if not pd.isna(media_c4) else 0
        
        grupo['Unidade Totais'] = total_up_original
        
        # Calcula diferenças (agora sempre em relação ao total original)
        grupo['C1: Diferença UP'] = total_up_c1 - total_up_original
        grupo['C2: Diferença UP'] = total_up_c2 - total_up_original
        grupo['C3: Diferença UP'] = total_up_c3 - total_up_original
        grupo['C4: Diferença UP'] = total_up_c4 - total_up_original
        
        # Cálculo dos meses a esticar (diferença / média UP)
        grupo['C1 - Meses a esticar'] = np.ceil(abs(grupo['C1: Diferença UP'].iloc[0]) / grupo['C1: Média UP'].iloc[0]) if grupo['C1: Média UP'].iloc[0] != 0 else 0
        grupo['C2 - Meses a esticar'] = np.ceil(abs(grupo['C2: Diferença UP'].iloc[0]) / grupo['C2: Média UP'].iloc[0]) if grupo['C2: Média UP'].iloc[0] != 0 else 0
        grupo['C3 - Meses a esticar'] = np.ceil(abs(grupo['C3: Diferença UP'].iloc[0]) / grupo['C3: Média UP'].iloc[0]) if grupo['C3: Média UP'].iloc[0] != 0 else 0
        grupo['C4 - Meses a esticar'] = np.ceil(abs(grupo['C4: Diferença UP'].iloc[0]) / grupo['C4: Média UP'].iloc[0]) if grupo['C4: Média UP'].iloc[0] != 0 else 0
        
        # Determina o número de meses a esticar (o maior entre os cenários)
        meses_esticar = max(
            grupo['C1 - Meses a esticar'].iloc[0],
            grupo['C2 - Meses a esticar'].iloc[0],
            grupo['C3 - Meses a esticar'].iloc[0],
            grupo['C4 - Meses a esticar'].iloc[0]
        )
        
        # Cria a coluna Data_rebalanc. com as datas originais
        grupo['Data_rebalanc.'] = grupo['data_mensal']
        
        # Se precisar esticar (adicionar meses)
        if meses_esticar > 0:
            ultima_data = grupo['data_mensal'].max()
            novos_meses = []
            
            # Adiciona novos meses após a última data
            for i in range(1, int(meses_esticar) + 1):
                nova_data = ultima_data + pd.DateOffset(months=i)
                novos_meses.append(nova_data)
            
            # Cria novas linhas para os meses adicionais
            if novos_meses:
                novas_linhas = pd.DataFrame({
                    'Regional Produção': regional_producao,
                    'Abertura Regional': abertura_regional,
                    obra_col: obra,
                    'Data_rebalanc.': novos_meses,
                    'data_mensal': pd.NaT,  # Mantém como NaT para indicar que é mês adicional
                    'data_estrutura': grupo['data_estrutura'].iloc[0],
                    'data_inicio_ponderacao': grupo['data_inicio_ponderacao'].iloc[0],
                    'C3: tipo_ponderacao': 'Mês adicional',
                    'C4: tipo_ponderacao': 'Mês adicional',
                    # Outras colunas permanecem com valores padrão ou NaN
                })
                
                # Concatena as novas linhas ao grupo original
                grupo = pd.concat([grupo, novas_linhas], ignore_index=True)
                
                # Calcula o valor a distribuir para cada cenário (divisão uniforme)
                diff_c1 = abs(grupo['C1: Diferença UP'].iloc[0])
                diff_c2 = abs(grupo['C2: Diferença UP'].iloc[0])
                diff_c3 = abs(grupo['C3: Diferença UP'].iloc[0])
                diff_c4 = abs(grupo['C4: Diferença UP'].iloc[0])
                
                # Calcula o valor uniforme para cada cenário
                valor_uniforme_c1 = diff_c1 / grupo['C1 - Meses a esticar'].iloc[0] if grupo['C1 - Meses a esticar'].iloc[0] > 0 else 0
                valor_uniforme_c2 = diff_c2 / grupo['C2 - Meses a esticar'].iloc[0] if grupo['C2 - Meses a esticar'].iloc[0] > 0 else 0
                valor_uniforme_c3 = diff_c3 / grupo['C3 - Meses a esticar'].iloc[0] if grupo['C3 - Meses a esticar'].iloc[0] > 0 else 0
                valor_uniforme_c4 = diff_c4 / grupo['C4 - Meses a esticar'].iloc[0] if grupo['C4 - Meses a esticar'].iloc[0] > 0 else 0
                
                # Índices das novas linhas
                idx_novas_linhas = grupo.index[-len(novos_meses):]
                
                # Aplica o valor uniforme nos meses necessários para cada cenário
                if grupo['C1 - Meses a esticar'].iloc[0] > 0:
                    grupo.loc[idx_novas_linhas[:int(grupo['C1 - Meses a esticar'].iloc[0])], 'C1: UP balanc. 1ºQ'] = valor_uniforme_c1
                
                if grupo['C2 - Meses a esticar'].iloc[0] > 0:
                    grupo.loc[idx_novas_linhas[:int(grupo['C2 - Meses a esticar'].iloc[0])], 'C2: UP balanc. Med.'] = valor_uniforme_c2
                
                if grupo['C3 - Meses a esticar'].iloc[0] > 0:
                    grupo.loc[idx_novas_linhas[:int(grupo['C3 - Meses a esticar'].iloc[0])], 'C3: 3m Med. + 1ºQ'] = valor_uniforme_c3
                
                if grupo['C4 - Meses a esticar'].iloc[0] > 0:
                    grupo.loc[idx_novas_linhas[:int(grupo['C4 - Meses a esticar'].iloc[0])], 'C4: 3m Real. + 1ºQ'] = valor_uniforme_c4
        
        # Se precisar retrair (remover meses)
        elif meses_esticar < 0:
            # Remove os últimos X meses
            grupo = grupo.sort_values('data_mensal').iloc[:-int(abs(meses_esticar))]
        
        # Recalcula os totais após o rebalanceamento para garantir que estão corretos
        grupo['C1: Total UP'] = total_up_original
        grupo['C2: Total UP'] = total_up_original
        grupo['C3: Total UP'] = total_up_original
        grupo['C4: Total UP'] = total_up_original
        
        # Adiciona o grupo rebalanceado à lista
        dfs_rebalanceados.append(grupo)
    
    # Concatena todos os grupos rebalanceados
    df_rebalanceado = pd.concat(dfs_rebalanceados, ignore_index=True)
    
    # Recalcula as porcentagens acumuladas e variações após o rebalanceamento
    for obra, grupo in df_rebalanceado.groupby(obra_col):
        # Cálculos de porcentagem acumulada
        total_up_c3 = grupo['C3: 3m Med. + 1ºQ'].sum()
        total_up_c4 = grupo['C4: 3m Real. + 1ºQ'].sum()
        
        if total_up_c3 > 0:
            df_rebalanceado.loc[grupo.index, '%AMP quartil'] = (grupo['C1: UP balanc. 1ºQ'] / total_up_c3).cumsum()
            df_rebalanceado.loc[grupo.index, '%AMP mediana'] = (grupo['C2: UP balanc. Med.'] / total_up_c3).cumsum()
        
        if total_up_c4 > 0:
            df_rebalanceado.loc[grupo.index, '%AMP real'] = (grupo['UP balanceada real'] / total_up_c4).cumsum()
        
        # Cálculos de variação percentual
        df_rebalanceado.loc[grupo.index, 'VP Bruta quartil'] = df_rebalanceado.loc[grupo.index, '%AMP quartil'].diff()
        df_rebalanceado.loc[grupo.index, 'VP Bruta mediana'] = df_rebalanceado.loc[grupo.index, '%AMP mediana'].diff()
        df_rebalanceado.loc[grupo.index, 'VP Bruta real'] = df_rebalanceado.loc[grupo.index, '%AMP real'].diff()

    return df_rebalanceado

def salvar_em_xlsx(df):
    """Salva o DataFrame como arquivo Excel em memória"""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Curva Rebalanceada')
    output.seek(0)
    return output

def main():
    """Função principal da aplicação Streamlit"""
    st.title("Rebalanceamento de Curvas de Produção (Obra a Obra)")

    # Upload do arquivo
    uploaded_file = st.file_uploader("Escolha o arquivo XLSX (Base Principal)", type=["xlsx"])

    if uploaded_file is not None:
        try:
            df = pd.read_excel(uploaded_file, sheet_name="Base Principal", engine="openpyxl")
        except Exception as e:
            st.error(f"Erro ao ler a aba 'Base Principal': {e}")
            return

        st.success("Arquivo carregado com sucesso!")

        # Input da data de referência
        data_referencia = st.date_input(
            "Selecione a data de referência do forecast",
            value=pd.to_datetime("06/01/2025"),
            min_value=pd.to_datetime("01/01/2000"),
            max_value=pd.to_datetime("31/12/2030"),
            help="Data base para cálculo do rebalanceamento"
        )

        # Processamento dos dados
        df_resultado = calcular_balanceamento_por_obra(
            df,
            data_col="Mensal",
            up_col="UP",
            data_estrutura_col="Início Estrutura",
            coef_quartil_col="Coeficiente Quartil",
            coef_mediana_col="Coeficiente Mediana",
            coef_real_col="Coeficiente Real",
            obra_col="Obra",
            data_referencia=data_referencia
        )

        # Colunas para exibição (com nova ordem)
        colunas_saida = [
            "Regional Produção", "Abertura Regional", "Obra", "Mensal", "Data_rebalanc.", "UP", "Unidade Totais",
            "Início Estrutura", "Recurso CEI016", "data_inicio_ponderacao",
            "C1: UP balanc. 1ºQ", "C1: Média UP", "C1: Total UP", "C1: Diferença UP", "C1 - Meses a esticar",
            "C2: UP balanc. Med.", "C2: Média UP", "C2: Total UP", "C2: Diferença UP", "C2 - Meses a esticar",
            "C3: 3m Med. + 1ºQ", "C3: Média UP", "C3: Total UP", "C3: Diferença UP", "C3 - Meses a esticar", "C3: tipo_ponderacao",
            "C4: 3m Real. + 1ºQ", "C4: Média UP", "C4: Total UP", "C4: Diferença UP", "C4 - Meses a esticar", "C4: tipo_ponderacao"
            #,"%AMP quartil", "%AMP mediana", "%AMP real",
            #"VP Bruta quartil", "VP Bruta mediana", "VP Bruta real"
        ]
        
        # Garante que todas as colunas existam no DataFrame
        colunas_existentes = [col for col in colunas_saida if col in df_resultado.columns]
        df_saida = df_resultado[colunas_existentes].copy()

        # Exibição dos resultados
        st.write("Resultado Final:")
        st.dataframe(df_saida)

        # Download do arquivo processado
        excel_bytes = salvar_em_xlsx(df_saida)
        st.download_button(
            label="📥 Baixar Excel Rebalanceado",
            data=excel_bytes,
            file_name="curva_rebalanceada.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

if __name__ == "__main__":
    main()
