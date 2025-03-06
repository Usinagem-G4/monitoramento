import streamlit as st
import pandas as pd
import datetime
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.formatting.rule import CellIsRule
from zoneinfo import ZoneInfo
import time

def calcular_tempo(arquivo_excel):
    """Calcula o tempo restante e excedente na planilha."""
    
    def formatar_tempo(delta):
        total_segundos = int(delta.total_seconds())
        horas, resto = divmod(total_segundos, 3600)
        minutos, segundos = divmod(resto, 60)
        return f"{horas:02d}:{minutos:02d}:{segundos:02d}"
    
    try:
        df = pd.read_excel(arquivo_excel)
    except FileNotFoundError:
        df = pd.DataFrame(columns=['Item', 'Operador', 'Termino', 'Tempo restante', 'Tempo excedente'])
        df.to_excel(arquivo_excel, index=False)
        st.warning(f"Arquivo {arquivo_excel} criado com colunas padrão.")

    if 'Tempo restante' not in df.columns:
        df['Tempo restante'] = ''
    if 'Tempo excedente' not in df.columns:
        df['Tempo excedente'] = ''

    for index, row in df.iterrows():
        if pd.notna(row.get('Termino')):
            termino_str = str(row['Termino'])
            try:
                termino = datetime.datetime.strptime(termino_str, '%d/%m/%Y %H:%M:%S')
            except ValueError:
                try:
                    hora = datetime.datetime.strptime(termino_str, '%H:%M:%S').time()
                    data_atual = datetime.date.today()
                    termino = datetime.datetime.combine(data_atual, hora)
                except ValueError:
                    st.error(f"Formato inválido na linha {index + 2}: {termino_str}")
                    continue

            agora = datetime.datetime.now(ZoneInfo("America/Sao_Paulo"))
            termino = termino.replace(tzinfo=ZoneInfo("America/Sao_Paulo"))

            if termino > agora:
                tempo_restante = termino - agora
                df.at[index, 'Tempo restante'] = formatar_tempo(tempo_restante)
                df.at[index, 'Tempo excedente'] = 'Dentro do tempo'
            else:
                tempo_excedente = agora - termino
                df.at[index, 'Tempo restante'] = 'Expirado'
                df.at[index, 'Tempo excedente'] = formatar_tempo(tempo_excedente)
        else:
            df.at[index, 'Tempo restante'] = ''
            df.at[index, 'Tempo excedente'] = ''

    df.to_excel(arquivo_excel, index=False)
    
    # Aplicar formatação condicional
    wb = load_workbook(arquivo_excel)
    ws = wb.active
    red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
    rule = CellIsRule(operator="lessThan", formula=['0'], stopIfTrue=True, fill=red_fill)
    ws.conditional_formatting.add(f"D2:D{len(df) + 1}", rule)
    wb.save(arquivo_excel)
    
    return df

def main():
    st.title("Monitoramento de Tempo em Tempo Real")
    
    # Configurações da página
    st.sidebar.header("Configurações")
    refresh_rate = st.sidebar.selectbox("Atualização automática:", 
                                      [60, 300, 600], 
                                      index=0,
                                      help="Intervalo de atualização em segundos")
    
    # Botão de upload principal
    st.subheader("Carregar Planilha")
    uploaded_file = st.file_uploader(
        "Selecione o arquivo monitoramento.xlsx",
        type=["xlsx"],
        accept_multiple_files=False,
        key="file_uploader"
    )
    
    # Verifica se foi feito upload de arquivo
    if uploaded_file is not None:
        # Salva o arquivo carregado
        with open("monitoramento.xlsx", "wb") as f:
            f.write(uploaded_file.getbuffer())
        st.success("Arquivo carregado com sucesso!")
    
    # Usa o arquivo padrão se não houver upload
    arquivo_excel = "monitoramento.xlsx"
    
    # Seção de atualização automática
    st.sidebar.markdown("---")
    if st.sidebar.button("Forçar Atualização"):
        st.rerun()
    
    # Atualização automática
    if 'last_refresh' not in st.session_state:
        st.session_state.last_refresh = time.time()
    
    if time.time() - st.session_state.last_refresh > refresh_rate:
        st.session_state.last_refresh = time.time()
        st.rerun()

    # Processamento principal
    try:
        df = calcular_tempo(arquivo_excel)
        
        # Exibição dos dados
        st.subheader("Dados Atualizados")
        st.dataframe(df.style.applymap(lambda x: 'background-color: #ff0000' if x == 'Expirado' else '', 
                                     subset=['Tempo restante']))
        
        # Botão de download
        with open(arquivo_excel, "rb") as file:
            st.download_button(
                label="Baixar Planilha Atualizada",
                data=file,
                file_name="monitoramento_atualizado.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        
    except Exception as e:
        st.error(f"Erro ao processar o arquivo: {str(e)}")
        st.info("Certifique-se de que o arquivo tem o formato correto com as colunas: Item, Operador, Termino")

if __name__ == '__main__':
    main()
