import pandas as pd


def formatar_numero(valor):
    # Verificar se o valor é numérico
    if isinstance(valor, (int, float)):
        # Converter números para o formato brasileiro
        return '{:,.2f}'.format(valor).replace('.', '*').replace(',', '.').replace('*', ',')
    else:
        return valor

def transformar_planilha(input_file, output_file):
    # Carregar a planilha Excel em um DataFrame usando pandas
    df = pd.read_excel(input_file)
    
    # Extrair apenas a primeira palavra da coluna
    df['Produto'] = df['Produto'].str.split().str[0]
    
    # Identificar e guardar registros de "Restituição de Capital"
    mask_restituicao = (df['Tipo de Evento'].str.lower().str.strip() == 'restituição de capital')
    restituicao_capital = df[mask_restituicao].copy()
    
    # Filtrar e remover registros de "Restituição de Capital"
    df = df[~mask_restituicao]
    
    # Remover as últimas 3 linhas ANTES de verificar eventos desconhecidos
    df = df.drop(df.tail(3).index)
    
    # Identificar tipos de evento conhecidos
    tipos_conhecidos = [
        'juros sobre capital próprio',
        'reembolso - dividendo',
        'reembolso',
        'reembolso - juros sobre capital próprio',
        'rendimento',
        'dividendo'
    ]
    
    # Identificar registros com tipos de evento não tratados
    mask_desconhecidos = ~df['Tipo de Evento'].str.lower().str.strip().isin(tipos_conhecidos)
    eventos_desconhecidos = df[mask_desconhecidos][['Produto', 'Tipo de Evento']].copy()
    
    # Extrair apenas a primeira palavra da coluna Instituição
    df['Instituição'] = df['Instituição'].str.split().str[0]

    df.loc[df['Tipo de Evento'].str.lower() == 'juros sobre capital próprio', 'Tipo de Evento'] = 'JSCP'
    df.loc[df['Tipo de Evento'].str.lower() == 'reembolso - dividendo', 'Tipo de Evento'] = 'REEMBOLSO(DIV)'
    df.loc[df['Tipo de Evento'].str.lower() == 'reembolso', 'Tipo de Evento'] = 'REEMBOLSO(DIV)'
    df.loc[df['Tipo de Evento'].str.lower() == 'reembolso - juros sobre capital próprio', 'Tipo de Evento'] = 'REEMBOLSO(JSCP)'    
    df.loc[df['Tipo de Evento'].str.lower() == 'rendimento', 'Tipo de Evento'] = 'RENDIMENTO'    
    df.loc[df['Tipo de Evento'].str.lower() == 'dividendo', 'Tipo de Evento'] = 'DIVIDENDO'

    # Excluir as colunas 5 e 6 (Coluna E e Coluna F)
    df = df.drop(columns=['Quantidade', 'Preço unitário'])

    df['Instituição'], df['Valor líquido'] = df['Valor líquido'], df['Instituição']

    df.insert(4, 'Coluna 5', '')
    df['Coluna 6'] = 'BRL'

    df['Valor líquido'], df['Coluna 6'] = df['Coluna 6'], df['Valor líquido']

    df['Instituição'] = df['Instituição'].apply(formatar_numero)

    df.rename(columns={'Produto': 'ativo'}, inplace=True)
    df.rename(columns={'Pagamento': 'date'}, inplace=True)
    df.rename(columns={'Tipo de Evento': 'evento'}, inplace=True)
    df.rename(columns={'Instituição': 'valor'}, inplace=True)
    df.rename(columns={'Coluna 5': 'irrf'}, inplace=True)
    df.rename(columns={'Valor líquido': 'moeda'}, inplace=True)
    df.rename(columns={'Coluna 6': 'corretora'}, inplace=True)
    df = df.drop(df.tail(3).index)

    # Salvar a planilha transformada em um novo arquivo
    df.to_csv(output_file, index=False, sep=';', encoding='utf-8')
    
    # Exibir mensagem sobre Restituição de Capital, se houver registros
    if not restituicao_capital.empty:
        ativos_restituicao = restituicao_capital['Produto'].unique()
        
        print("\n" + "="*80)
        print("Os ativos abaixo devem ser lançados apenas na aba Operações:")
        print()
        for ativo in sorted(ativos_restituicao):
            print(f"  • {ativo}")
        print()
        print("A Restituição de Capital é um valor retornado ao acionista após um evento")
        print("de redução de capital. Esse pagamento é um rendimento isento e não tributável")
        print("e deve ser abatido do seu preço médio. Por exemplo, ocorreu recentemente com BBSE3.")
        print()
        print("A lógica é idêntica a amortização que aplicamos aos FIIs. Vejamos como preencher")
        print("na planilha:")
        print("  Aba = Operações")
        print("  Ativo = AÇÃO (BBSE3, por exemplo)")
        print("  Data = Data da restituição, conforme fato relevante (10/01/20, por exemplo)")
        print("  Evento = R.CAP")
        print("  Quantidade = 0")
        print("  Preço = -Valor restituído (-135,22)")
        print("  Taxas = 0,00")
        print("  Corretora = A sua")
        print("="*80 + "\n")
    
    # Exibir mensagem sobre eventos desconhecidos, se houver registros
    if not eventos_desconhecidos.empty:
        print("\n" + "="*80)
        print("Não existe tipo de evento compatível para os ativos abaixo. Conferir na Dlombello o nome exato do tipo de evento e fazer a conversão:")
        print()
        for _, row in eventos_desconhecidos.iterrows():
            print(f"  • {row['Produto']} - {row['Tipo de Evento']}")
        print("="*80 + "\n")

if __name__ == "__main__":
    # Defina o nome do arquivo de entrada e saída
    arquivo_entrada = '/files_vol/proventos.xlsx'
    arquivo_saida = '/files_vol/prov.csv'
    
    # Chame a função para realizar a transformação
    transformar_planilha(arquivo_entrada, arquivo_saida)
