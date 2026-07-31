import pandas as pd
import numpy as np
import json
import os

def generate_dashboard():
    # Load Data
    # --- LOAD CRM DATA ---
    try:
        df_crm = pd.read_excel(r"C:\Users\Isac\Desktop\Dados_CRM_Atividade.xlsx")
        df_crm['Vendedor'] = df_crm['Vendedor'].astype(str).str.upper().str.strip()
        crm_grouped = df_crm.groupby('Vendedor')[['Novos_Cadastros', 'Cotacoes', 'Interacoes_CRM']].sum().reset_index()
    except:
        crm_grouped = pd.DataFrame(columns=['Vendedor', 'Novos_Cadastros', 'Cotacoes', 'Interacoes_CRM'])

    clientes_path = r'C:\Users\Isac\Desktop\Planilhas geradas pela IA\Clientes_Detalhes_Extras.xlsx'
    interacoes_path = r'C:\Users\Isac\Desktop\Planilhas geradas pela IA\Interacoes_Sistema.xlsx'
    propostas_path = r"C:\Users\Isac\Desktop\Planilhas geradas pela IA\Propostas_Sistema.xlsx"
    nfe_path = r"C:\Users\Isac\Desktop\Planilhas geradas pela IA\Notas_Fiscais_Sistema.xlsx"
    
    df_clientes = pd.read_excel(clientes_path)
    df_interacoes = pd.read_excel(interacoes_path)
    df_propostas = pd.read_excel(propostas_path)
    df_nfe = pd.read_excel(nfe_path, engine='openpyxl')

    try:
        cfop_counts = df_nfe['CFOP'].value_counts().head(5)
        top_cfop_labels = [str(x)[:40] + ('...' if len(str(x)) > 40 else '') for x in cfop_counts.index.tolist()]
        top_cfop_data = [int(x) for x in cfop_counts.values.tolist()]
    except Exception as e:
        print("Erro CFOP:", e)
        top_cfop_labels = []
        top_cfop_data = []

    try:
        cfop_counts = df_nfe['CFOP'].value_counts().head(5)
        top_cfop_labels = [str(x)[:40] + ('...' if len(str(x)) > 40 else '') for x in cfop_counts.index.tolist()]
        top_cfop_data = [int(x) for x in cfop_counts.values.tolist()]
    except Exception as e:
        print("Erro CFOP:", e)
        top_cfop_labels = []
        top_cfop_data = []


    try:
        pot_clientes_df = pd.read_excel(r'C:\Users\Isac\Desktop\Planilhas geradas pela IA\Potenciais_Clientes_150.xlsx')
        pot_clientes_df = pot_clientes_df.fillna('')
        pot_table_rows = ''
        for _, row in pot_clientes_df.iterrows():
            match_color = '#2e7d32' if row['Grau de Afinidade'] == 'Alto' else '#ff9800'
            pot_table_rows += f"""<tr>
                <td>{row.get('Razão Social / Nome Fantasia', '')}</td>
                <td>{row.get('CNAE Principal', '')}</td>
                <td>{row.get('Segmento', '')}</td>
                <td>{row.get('Estado', '')} - {row.get('Cidade', '')}</td>
                <td><a href='http://{row.get('Site da Empresa', '')}' target='_blank' style='color:#007bff;'>{row.get('Site da Empresa', '')}</a></td>
                <td><span style='color: white; background-color: {match_color}; padding: 3px 8px; border-radius: 12px; font-size: 11px;'>{row.get('Grau de Afinidade', '')}</span></td>
            </tr>"""
    except Exception as e:
        pot_table_rows = "<tr><td colspan='6'>Nenhum dado encontrado</td></tr>"




    
    # Clean Data Clientes
    df_clientes['Ult_Data_Venda'] = pd.to_datetime(df_clientes.get('Ult_Data_Venda'), errors='coerce')
    df_clientes['Ult_Data_Cotacao'] = pd.to_datetime(df_clientes.get('Ult_Data_Cotacao'), errors='coerce')
    df_clientes['Ticket_Medio_Num'] = pd.to_numeric(df_clientes.get('Ticket_Medio', pd.Series(np.zeros(len(df_clientes)))), errors='coerce').fillna(0)
    df_clientes['Ticket_Total_Num'] = pd.to_numeric(df_clientes.get('Ticket_Total', pd.Series(np.zeros(len(df_clientes)))), errors='coerce').fillna(0)
    
    # Clean Data Propostas
    df_propostas['ValorTotalNum'] = pd.to_numeric(df_propostas.get('ValorTotal', pd.Series(np.zeros(len(df_propostas)))), errors='coerce').fillna(0)
    df_propostas['Status'] = df_propostas.get('Status', pd.Series([""]*len(df_propostas))).astype(str).str.strip()
    df_propostas['Vendedor'] = df_propostas.get('Vendedor', pd.Series([""]*len(df_propostas))).astype(str).str.strip()
    
    # Clean Data NFe
    df_nfe['DataEmissao'] = pd.to_datetime(df_nfe['DataEmissao'], errors='coerce')
    df_nfe['Ano'] = df_nfe['DataEmissao'].dt.year
    df_nfe['MesAno'] = df_nfe['DataEmissao'].dt.to_period('M').astype(str)
    df_nfe['ValorTotalNF'] = pd.to_numeric(df_nfe.get('ValorTotalNF', pd.Series(np.zeros(len(df_nfe)))), errors='coerce').fillna(0)
    df_nfe['StatusDesc'] = df_nfe.get('StatusDesc', pd.Series([""]*len(df_nfe))).astype(str).str.strip()
    df_nfe['Vendedor'] = df_nfe.get('Vendedor', pd.Series([""]*len(df_nfe))).astype(str).str.strip()
    
    # Base valid invoices (for all analyses)
    df_nfe_valid = df_nfe[df_nfe['StatusDesc'].isin(['EMITIDA', 'AUTORIZADA', ''])].copy()
    
    # Categorize CFOPs
    def categorize_cfop(val):
        val_str = str(val).upper()
        if 'VENDA' in val_str: return 'Venda de Material'
        elif 'TRANSFERENCIA' in val_str: return 'Transferências'
        elif 'CONSERTO' in val_str or 'RETORNO' in val_str or 'DEVOLUCAO' in val_str or 'DEVOLUÇÃO' in val_str: return 'Retornos e Consertos'
        elif 'INDUSTRIALIZACAO' in val_str or 'INDUSTRIALIZAÇÃO' in val_str: return 'Serviço Prestado (Industrialização)'
        else: return 'Outros'

    if 'CFOP' in df_nfe_valid.columns:
        df_nfe_valid['CategoriaCFOP'] = df_nfe_valid['CFOP'].apply(categorize_cfop)
    else:
        df_nfe_valid['CategoriaCFOP'] = 'Venda de Material'
        
    # Global 'Sales Only' mask (for Seller Performance and General KPI) - includes Materials + Services
    mask_sales_only = df_nfe_valid['CategoriaCFOP'].isin(['Venda de Material', 'Serviço Prestado (Industrialização)'])
    df_nfe_sales = df_nfe_valid[mask_sales_only].copy()
    # Helpers
    def cur_format_py(val):
        return f"R$ {val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        
    def date_format_py(val):
        return pd.to_datetime(val).strftime('%d/%m/%Y') if pd.notna(val) else 'N/I'

    total_fat_real = cur_format_py(df_nfe_valid['ValorTotalNF'].sum())
    total_prop_abertas = cur_format_py(df_propostas[df_propostas['Status'].isin(['ENVIADA', 'NAO ENVIADA', 'STANDBY'])]['ValorTotalNum'].sum())


    # -----------------------------------------------------
    # PREPARE JSON DATA FOR INTERACTIVE CFOP FILTERING
    # -----------------------------------------------------
    cfop_data = {}
    categorias_cfop = ['TODOS', 'Materiais + Serviços', 'Venda de Material', 'Serviço Prestado (Industrialização)', 'Transferências', 'Retornos e Consertos']
    
    for cat in categorias_cfop:
        if cat == 'TODOS':
            df_subset = df_nfe_valid
        elif cat == 'Materiais + Serviços':
            df_subset = df_nfe_sales
        else:
            df_subset = df_nfe_valid[df_nfe_valid['CategoriaCFOP'] == cat]
            
        total_val = df_subset['ValorTotalNF'].sum()
        
        fat_por_ano = df_subset[df_subset['Ano'].notna()].groupby('Ano')['ValorTotalNF'].sum().sort_index()
        fat_por_mes = df_subset[df_subset['MesAno'].notna() & (df_subset['MesAno'] != 'NaT')].groupby('MesAno')['ValorTotalNF'].sum().sort_index()
        top_clientes_fat = df_subset.groupby('ClienteNome')['ValorTotalNF'].sum().sort_values(ascending=False).head(10)
        top_20_clientes_fat = df_subset.groupby('ClienteNome')['ValorTotalNF'].sum().sort_values(ascending=False).head(20)
        rows_html = ""
        for emp, val in top_20_clientes_fat.items():
            empresa = str(emp)
            valor_fat = cur_format_py(val)
            rows_html += f"<tr><td>{empresa}</td><td style='font-weight:bold; color:#0288d1;'>{valor_fat}</td></tr>"
            
        cfop_data[cat] = {
            'total_str': cur_format_py(total_val),
            'ano_labels': [str(int(x)) for x in fat_por_ano.index],
            'ano_data': [float(x) for x in fat_por_ano.values],
            'mes_labels': [str(x) for x in fat_por_mes.index],
            'mes_data': [float(x) for x in fat_por_mes.values],
            'cli_labels': [(str(x)[:20] + '...') if len(str(x)) > 20 else str(x) for x in top_clientes_fat.index],
            'cli_data': [float(x) for x in top_clientes_fat.values],
            'table_rows': rows_html
        }
    cfop_json_str = json.dumps(cfop_data)

    # -----------------------------------------------------
    # PREPARE JSON DATA FOR INTERACTIVE SELLER FILTERING
    # -----------------------------------------------------
    # Identify Top 7 sellers from NFe Sales
    top_sellers_series = df_nfe_sales.groupby('Vendedor')['ValorTotalNF'].sum().sort_values(ascending=False)
    top_7_sellers = top_sellers_series.head(7).index.tolist()
    
    def categorize_seller(v):
        v = str(v).strip()
        if v in top_7_sellers: return v
        if v == '' or v == 'nan': return 'N/I'
        return 'Outros'

    df_nfe_sales['VendedorGrp'] = df_nfe_sales['Vendedor'].apply(categorize_seller)
    df_propostas['VendedorGrp'] = df_propostas['Vendedor'].apply(categorize_seller)
    
    # Active open proposals for table
    df_abertas = df_propostas[df_propostas['Status'].isin(['ENVIADA', 'NAO ENVIADA', 'STANDBY'])]
    
    seller_categories = ['TODOS'] + top_7_sellers + ['Outros']
    seller_data = {}
    
    for s_cat in seller_categories:
        # NFe subset
        df_sub_nfe = df_nfe_sales if s_cat == 'TODOS' else df_nfe_sales[df_nfe_sales['VendedorGrp'] == s_cat]
        # Propostas subset
        df_sub_prop = df_abertas if s_cat == 'TODOS' else df_abertas[df_abertas['VendedorGrp'] == s_cat]
        
        # 1. Total Revenue (NFe)
        total_val = df_sub_nfe['ValorTotalNF'].sum()
        
        # 2. Year by Year (NFe)
        fat_por_ano = df_sub_nfe[df_sub_nfe['Ano'].notna()].groupby('Ano')['ValorTotalNF'].sum().sort_index()
        
        # 2b. Month by Month (NFe)
        fat_por_mes = df_sub_nfe[df_sub_nfe['MesAno'].notna() & (df_sub_nfe['MesAno'] != 'NaT')].groupby('MesAno')['ValorTotalNF'].sum().sort_index()
        
        # 3. Bar Chart logic: if ALL, show top 10 sellers. If specific seller, show only them.
        if s_cat == 'TODOS':
            vend_bar = df_sub_nfe.groupby('Vendedor')['ValorTotalNF'].sum().sort_values(ascending=False).head(10)
        else:
            vend_bar = df_sub_nfe.groupby('Vendedor')['ValorTotalNF'].sum().sort_values(ascending=False).head(1)
            
        # 4. Open Proposals Table
        top_abertas = df_sub_prop.sort_values(by='ValorTotalNum', ascending=False).head(50)
        rows_html = ""
        for _, row in top_abertas.iterrows():
            empresa = str(row.get('Nome') or row.get('Razao'))
            valor_prop = cur_format_py(row['ValorTotalNum'])
            vendedor_n = str(row.get('Vendedor', 'N/I'))
            status = str(row.get('Status', 'N/I'))
            rows_html += f"<tr><td>{empresa}</td><td style='font-weight:bold; color:var(--primary-red);'>{valor_prop}</td><td>{vendedor_n}</td><td><span class='badge bad'>{status}</span></td></tr>"
            
        seller_data[s_cat] = {
            'total_str': cur_format_py(total_val),
            'total_val': float(total_val),
            'ano_labels': [str(int(x)) for x in fat_por_ano.index],
            'ano_data': [float(x) for x in fat_por_ano.values],
            'mes_labels': [str(x) for x in fat_por_mes.index],
            'mes_data': [float(x) for x in fat_por_mes.values],
            'mes_labels': [str(x) for x in fat_por_mes.index],
            'mes_data': [float(x) for x in fat_por_mes.values],
            'bar_labels': [str(x) for x in vend_bar.index],
            'bar_data': [float(x) for x in vend_bar.values],
            'table_rows': rows_html
        }
    seller_json_str = json.dumps(seller_data)

    # --- PROCESS PERFORMANCE 360 ---
    # Aggregate NFe by seller for 360 Tab
    vend_nfe_360 = df_nfe_valid.groupby('Vendedor')['ValorTotalNF'].sum().reset_index()
    vend_nfe_360.rename(columns={'ValorTotalNF': 'Faturamento_NFe'}, inplace=True)
    vend_nfe_360['Vendedor'] = vend_nfe_360['Vendedor'].astype(str).str.upper().str.strip()
    
    df_360 = pd.merge(crm_grouped, vend_nfe_360, on='Vendedor', how='outer')
    df_360 = df_360[(df_360['Faturamento_NFe'] > 0) | (df_360['Interacoes_CRM'] > 0)]
    df_360['Faturamento_NFe'] = df_360['Faturamento_NFe'].fillna(0)
    df_360['Interacoes_CRM'] = df_360['Interacoes_CRM'].fillna(0)
    df_360['Cotacoes'] = df_360['Cotacoes'].fillna(0)
    df_360['Novos_Cadastros'] = df_360['Novos_Cadastros'].fillna(0)
    
    # Calculate KPIs
    df_360['Conv_Cotacao'] = df_360.apply(lambda row: row['Faturamento_NFe'] / row['Cotacoes'] if row['Cotacoes'] > 0 else 0, axis=1)
    df_360['Rend_Interacao'] = df_360.apply(lambda row: row['Faturamento_NFe'] / row['Interacoes_CRM'] if row['Interacoes_CRM'] > 0 else 0, axis=1)
    
    # Sort for table
    df_360_sorted = df_360.sort_values('Faturamento_NFe', ascending=False)
    
    rows_360 = ""
    for idx, row in df_360_sorted.iterrows():
        rows_360 += f"<tr><td>{row['Vendedor']}</td><td style='color:var(--primary-red); font-weight:bold;'>{cur_format_py(row['Faturamento_NFe'])}</td><td>{int(row['Interacoes_CRM'])}</td><td>{int(row['Cotacoes'])}</td><td>{int(row['Novos_Cadastros'])}</td><td>{cur_format_py(row['Conv_Cotacao'])}</td><td>{cur_format_py(row['Rend_Interacao'])}</td></tr>"
        
    # Chart Data (Top 10 by Interacoes for Scatter/Bar)
    top_360 = df_360.sort_values('Interacoes_CRM', ascending=False).head(15)
    
    perf360_json_data = {
        'labels': top_360['Vendedor'].tolist(),
        'faturamento': top_360['Faturamento_NFe'].tolist(),
        'interacoes': top_360['Interacoes_CRM'].tolist()
    }
    perf360_json_str = json.dumps(perf360_json_data)

    # -----------------------------------------------------
    # KPIs and Variables for other tabs
    # -----------------------------------------------------
    faturamento_total = df_nfe_sales['ValorTotalNF'].sum()
    faturamento_str = cur_format_py(faturamento_total)
    
    total_clientes = len(df_clientes)
    df_cotaram_sem_venda = df_clientes[
        (df_clientes['Ult_Data_Cotacao'].notna()) &
        ((df_clientes['Ult_Data_Venda'].isna()) | (df_clientes['Ult_Data_Venda'] < df_clientes['Ult_Data_Cotacao']))
    ]
    total_pipeline = len(df_cotaram_sem_venda)
    valor_pipeline = df_abertas['ValorTotalNum'].sum()
    valor_pipeline_str = cur_format_py(valor_pipeline)
    
    # Top Representantes (Clientes)
    if 'Representante' in df_clientes.columns:
        representantes = df_clientes['Representante'].value_counts().head(10)
        bar_labels = [str(x) for x in representantes.index]
        bar_data = [int(x) for x in representantes.values]
    else: bar_labels = []; bar_data = []

    if 'Estado' in df_clientes.columns:
        estados = df_clientes['Estado'].value_counts().head(5)
        estados_labels = [str(x) for x in estados.index]
        estados_data = [int(x) for x in estados.values]
    else: estados_labels = []; estados_data = []
        
    top_pipeline_html = df_cotaram_sem_venda.sort_values(by='Ult_Data_Cotacao', ascending=False).head(50)
    pipeline_rows = ""
    for _, row in top_pipeline_html.iterrows():
        pipeline_rows += f"<tr><td>{str(row.get('Nome') or row.get('Razao_Social'))}</td><td>{date_format_py(row.get('Ult_Data_Cotacao'))}</td><td>{str(row.get('Contato', 'N/I'))}</td><td><span class='badge bad'>Pendente</span></td></tr>"
        
    maiores_tickets = df_clientes[df_clientes['Ticket_Medio_Num'] > 0].sort_values(by='Ticket_Medio_Num', ascending=False).head(50)
    tickets_rows = ""
    for _, row in maiores_tickets.iterrows():
        tickets_rows += f"<tr><td>{str(row.get('Nome') or row.get('Razao_Social'))}</td><td style='font-weight:bold; color:var(--primary-red);'>{cur_format_py(row['Ticket_Medio_Num'])}</td><td>{date_format_py(row.get('Ult_Data_Venda'))}</td><td><span class='badge good'>VIP</span></td></tr>"

    maiores_totais = df_clientes[df_clientes['Ticket_Total_Num'] > 0].sort_values(by='Ticket_Total_Num', ascending=False).head(50)
    recompras_rows = ""
    for _, row in maiores_totais.iterrows():
        recompras_rows += f"<tr><td>{str(row.get('Nome') or row.get('Razao_Social'))}</td><td style='font-weight:bold;'>{cur_format_py(row['Ticket_Total_Num'])}</td><td>{date_format_py(row.get('Ult_Data_Venda'))}</td><td><span class='badge good'>Fidelizado</span></td></tr>"

    # HTML Buttons for Sellers
    seller_buttons_html = ""
    for s in seller_categories:
        disp_name = s.title() if s not in ['TODOS', 'N/I'] else ('Visão Geral' if s == 'TODOS' else s)
        active_cls = 'active' if s == 'TODOS' else ''
        # Use single quotes around string argument for JS call
        seller_buttons_html += f'<button class="filter-btn {active_cls}" onclick="filterVendedor(\'{s}\')">{disp_name}</button>\n'

    # HTML Template

    # --- PROCESS MAP DATA ---
    import unicodedata
    def remove_accents(s):
        if pd.isna(s): return ""
        nfkd = unicodedata.normalize('NFKD', str(s))
        return u"".join([c for c in nfkd if not unicodedata.combining(c)]).upper().strip()
        
    df_clientes['Cidade_Match'] = df_clientes['Cidade'].apply(lambda x: remove_accents(str(x).split(' - ')[0]))
    try:
        df_mun = pd.read_csv(r'C:\Users\Isac\Desktop\Planilhas geradas pela IA\municipios.csv')
        df_mun['Cidade_Match'] = df_mun['nome'].apply(remove_accents)
        # Drop duplicates by city name to prevent explosion (using first occurrence)
        df_mun = df_mun.drop_duplicates(subset=['Cidade_Match'])
        
        # Merge clients with municipalities
        df_map = pd.merge(df_clientes, df_mun[['Cidade_Match', 'latitude', 'longitude']], on='Cidade_Match', how='left')
        
        # Filter valid coords
        df_map = df_map[df_map['latitude'].notna() & df_map['longitude'].notna()]
        
        # Deduplicate to prevent massive repetition in the same city (e.g. WEG in Jaragua)
        df_map['Nome_Clean'] = df_map['Razao_Social'].fillna(df_map['Nome']).astype(str).str.upper().str.strip()
        df_map = df_map.drop_duplicates(subset=['Nome_Clean', 'Cidade_Match'])
        
        # Prepare list for JSON
        map_points = []
        import random
        for _, row in df_map.iterrows():
            # Add small natural random jitter (Gaussian) so points don't form a square
            lat_jitter = random.gauss(0, 0.005)
            lon_jitter = random.gauss(0, 0.005)
            
            nome = str(row.get('Razao_Social', ''))
            if nome == 'nan' or nome == '':
                nome = str(row.get('Nome', 'Cliente Desconhecido'))
                
            cidade = str(row.get('Cidade', ''))
            estado = str(row.get('Estado', ''))
            ticket = cur_format_py(row.get('Ticket_Medio_Num', 0))
            
            map_points.append({
                'lat': float(row['latitude']) + lat_jitter,
                'lon': float(row['longitude']) + lon_jitter,
                'nome': nome,
                'cidade': cidade,
                'estado': estado,
                'ticket': ticket,
                'tipo': 'cliente'
            })
        

        # Add International Clients
        international_clients = [
            {'lat': -17.7833, 'lon': -63.1821, 'nome': 'MANUFACTURAS TECNICAS MATEC SA', 'cidade': 'SANTA CRUZ DE LA SIERRA', 'estado': 'BOLIVIA', 'ticket': 'N/A', 'tipo': 'cliente'},
            {'lat': -17.7833, 'lon': -63.1821, 'nome': 'EDUARDO S.A', 'cidade': 'SANTA CRUZ DE LA SIERRA', 'estado': 'BOLIVIA', 'ticket': 'N/A', 'tipo': 'cliente'},
            {'lat': -19.5836, 'lon': -65.7531, 'nome': 'TEC FUCCA SRL', 'cidade': 'POTOSI', 'estado': 'BOLIVIA', 'ticket': 'N/A', 'tipo': 'cliente'},
            {'lat': -16.5000, 'lon': -68.1500, 'nome': 'TOMAS MOLLER CHINO', 'cidade': 'EL ALTO', 'estado': 'BOLIVIA', 'ticket': 'N/A', 'tipo': 'cliente'},
            {'lat': 19.4326, 'lon': -99.1332, 'nome': 'REGAL REXNORD MEXICO', 'cidade': 'MEXICO CITY', 'estado': 'MEXICO', 'ticket': 'N/A', 'tipo': 'cliente'},
            {'lat': 19.2826, 'lon': -99.6557, 'nome': 'MARTIN SPROCKET MEXICO', 'cidade': 'TOLUCA', 'estado': 'MEXICO', 'ticket': 'N/A', 'tipo': 'cliente'},
            {'lat': 19.8333, 'lon': -99.2000, 'nome': 'WEG MEXICO S.A. DE C.V.', 'cidade': 'HUEHUETOCA', 'estado': 'MEXICO', 'ticket': 'N/A', 'tipo': 'cliente'},
            {'lat': 38.7223, 'lon': -9.1393, 'nome': 'VLB TEC SA - PORTUGAL', 'cidade': 'LISBOA', 'estado': 'PORTUGAL', 'ticket': 'N/A', 'tipo': 'cliente'},
            {'lat': 41.2279, 'lon': -8.6210, 'nome': 'WEG PORTUGAL', 'cidade': 'MAIA', 'estado': 'PORTUGAL', 'ticket': 'N/A', 'tipo': 'cliente'},
            {'lat': -31.4280, 'lon': -62.0827, 'nome': 'WEG EQUIPAMIENTOS ELECTRICOS S.A.', 'cidade': 'SAN FRANCISCO', 'estado': 'ARGENTINA', 'ticket': 'N/A', 'tipo': 'cliente'},
            {'lat': 3.4516, 'lon': -76.5320, 'nome': 'RESORTES HERCULES S.A.S.', 'cidade': 'CALI', 'estado': 'COLOMBIA', 'ticket': 'N/A', 'tipo': 'cliente'},
            {'lat': -34.9011, 'lon': -56.1645, 'nome': 'MATEBOM SRL', 'cidade': 'MONTEVIDEO', 'estado': 'URUGUAI', 'ticket': 'N/A', 'tipo': 'cliente'}
        ]
        
        map_points.extend(international_clients)
        
        map_json_str = json.dumps(map_points)
    except Exception as e:
        print("Map error:", e)
        map_json_str = "[]"



    # -- NOVO: DADOS PARA TAB SERVICOS DE TEMPERA --
    try:
        df_tempera = pd.read_excel(r'Tempera_Itens.xlsx', engine='openpyxl')
        df_tempera['ValorTotalNF'] = pd.to_numeric(df_tempera['ValorTotalNF'], errors='coerce').fillna(0)
        df_tempera['DataEmissao'] = pd.to_datetime(df_tempera['DataEmissao'], errors='coerce')
        df_tempera = df_tempera.dropna(subset=['DataEmissao']).copy()
        
        if not df_tempera.empty:
            df_tempera['Ano'] = df_tempera['DataEmissao'].dt.year
            df_tempera['Mes'] = df_tempera['DataEmissao'].dt.month
            
            # Monthly Revenue grouped by year
            tempera_anos = df_tempera['Ano'].unique()
            tempera_monthly_dict = {str(int(ano)): [0.0]*12 for ano in tempera_anos}

            energy_costs_dict = {
                "2025": [0, 0, 0, 3437.94, 3317.45, 3831.39, 4429.39, 3830.65, 3364.3, 3832.14, 2982.75, 2835.48],
                "2026": [2294.01, 3609.01, 3513.06, 3007.29, 3007.29, 2892.75, 0, 0, 0, 0, 0, 0]
            }

            
            grp = df_tempera.groupby(['Ano', 'Mes'])['ValorTotalNF'].sum().reset_index()
            for _, row in grp.iterrows():
                ano = str(int(row['Ano']))
                mes = int(row['Mes']) - 1 # 0-indexed
                tempera_monthly_dict[ano][mes] = float(row['ValorTotalNF'])
            
            
            tempera_anos_list = sorted(list(tempera_monthly_dict.keys()), reverse=True)
            tempera_year_options_html = "".join([f'<option value="{ano}">{ano}</option>' for ano in tempera_anos_list])
            
            # --- NOVO: Tabela Anual ---
            tempera_yearly_tables_dict = {}
            for ano in tempera_anos:
                df_ano = df_tempera[df_tempera['Ano'] == ano].sort_values(by='DataEmissao', ascending=False)
                rows = []
                for _, row in df_ano.iterrows():
                    d = row['DataEmissao'].strftime('%d/%m/%Y') if pd.notnull(row['DataEmissao']) else '-'
                    n = str(row['NumeroDaNota']) if 'NumeroDaNota' in df_ano.columns else '-'
                    c = str(row['ClienteNome'])
                    c = (c[:40] + '...') if len(c) > 40 else c
                    v = cur_format_py(row['ValorTotalNF'])
                    rows.append(f"<tr><td>{d}</td><td>{n}</td><td>{c}</td><td style='font-weight:bold; color:var(--primary-red);'>{v}</td></tr>")
                tempera_yearly_tables_dict[str(int(ano))] = "".join(rows) if rows else "<tr><td colspan='4'>Nenhuma nota encontrada</td></tr>"

            
            # Top clients
            tempera_top_clients = df_tempera.groupby('ClienteNome')['ValorTotalNF'].sum().sort_values(ascending=False).head(5).reset_index()
            tempera_total_rev = df_tempera['ValorTotalNF'].sum()
            tempera_total_str = f"R$ {tempera_total_rev:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            
            tempera_clients_labels = [str(x)[:40] + ('...' if len(str(x)) > 40 else '') for x in tempera_top_clients['ClienteNome'].tolist()]
            tempera_clients_data = [float(x) for x in tempera_top_clients['ValorTotalNF'].tolist()]
        else:
            tempera_monthly_dict = {}
            tempera_yearly_tables_dict = {}
            tempera_year_options_html = '<option value="">Sem dados</option>'
            tempera_clients_labels = []
            tempera_clients_data = []
            tempera_total_str = "R$ 0,00"
            
    except Exception as e:
        print("Erro em tempera:", e)
        tempera_labels = []
        tempera_data = []
        tempera_yearly_labels = []
        tempera_yearly_data = []
        tempera_clients_labels = []
        tempera_clients_data = []
        tempera_total_str = "R$ 0,00"

    # -- NOVO: TABELA DE ULTIMAS NOTAS --
    try:
        df_ultimas = df_nfe.head(10).copy() # A planilha ja vem ordenada do sistema da mais nova para a mais velha
        df_ultimas['DataEmissao'] = df_ultimas['DataEmissao'].dt.strftime('%d/%m/%Y')
        df_ultimas['ValorTotalNF'] = df_ultimas['ValorTotalNF'].apply(cur_format_py)
        table_ultimas_nfe = df_ultimas[['DataEmissao', 'ClienteNome', 'ValorTotalNF', 'StatusDesc', 'Vendedor']].to_html(index=False, classes='dashboard-table', border=0)
    except Exception as e:
        print("Erro ultimas nfe:", e)
        table_ultimas_nfe = "<p>Erro ao carregar últimas notas.</p>"


    # -- NOVO: LOGO DA 3MI --
    logo_base64 = ""
    logo_path = r'C:\Users\Isac\Desktop\logo_3mi.png'
    try:
        import base64
        if os.path.exists(logo_path):
            with open(logo_path, 'rb') as lf:
                logo_base64 = base64.b64encode(lf.read()).decode('utf-8')
    except Exception as e:
        print("Erro ao carregar logo:", e)

    html_content = f"""<!DOCTYPE html>


<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Dashboard Comercial - 3MI</title>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;600;800&display=swap" rel="stylesheet">
    <style>
        :root {{
            --primary-red: #A6192E;
            --dark-red: #7a1020;
            --bg-color: #f4f7f6;
            --card-bg: #ffffff;
            --text-main: #333333;
            --text-muted: #777777;
            --sidebar-width: 250px;
        }}
        body {{ font-family: 'Inter', sans-serif; background-color: var(--bg-color); margin: 0; padding: 0; color: var(--text-main); display: flex; height: 100vh; overflow: hidden; }}
        .sidebar {{ width: var(--sidebar-width); background-color: #1a1a1a; color: white; display: flex; flex-direction: column; box-shadow: 2px 0 10px rgba(0,0,0,0.1); z-index: 10; }}
        .sidebar-header {{ padding: 0; background: white; text-align: center; display: flex; align-items: center; justify-content: center; height: 120px; }}
        .sidebar-header h1 {{ margin: 0; font-size: 32px; font-style: italic; letter-spacing: -2px; color: white; }}
        .menu {{ list-style: none; padding: 0; margin: 20px 0; flex: 1; }}
        .menu-item {{ padding: 15px 25px; cursor: pointer; font-size: 14px; font-weight: 600; border-left: 4px solid transparent; transition: all 0.2s; color: #bbb; }}
        .menu-item:hover {{ background-color: #333; color: white; }}
        .menu-item.active {{ background-color: #2a2a2a; border-left: 4px solid var(--primary-red); color: white; }}
        .main-wrapper {{ flex: 1; display: flex; flex-direction: column; overflow-y: auto; }}
        .topbar {{ background: white; padding: 20px 40px; box-shadow: 0 2px 10px rgba(0,0,0,0.05); display: flex; justify-content: space-between; align-items: center; }}
        .topbar h2 {{ margin: 0; font-weight: 600; }}
        .container {{ padding: 40px; max-width: 1400px; margin: 0 auto; width: 100%; box-sizing: border-box; }}
        .tab-content {{ display: none; animation: fadeIn 0.3s; }}
        .tab-content.active {{ display: block; }}
        @keyframes fadeIn {{ from {{ opacity: 0; transform: translateY(10px); }} to {{ opacity: 1; transform: translateY(0); }} }}
        .kpi-grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(300px, 1fr)); gap: 20px; margin-bottom: 40px; }}
        .kpi-card {{ background: var(--card-bg); padding: 25px; border-radius: 12px; box-shadow: 0 5px 15px rgba(0,0,0,0.03); border-top: 4px solid var(--primary-red); }}
        .kpi-title {{ font-size: 13px; color: var(--text-muted); text-transform: uppercase; font-weight: 600; margin-bottom: 10px; }}
        .kpi-value {{ font-size: 32px; font-weight: 800; color: var(--text-main); margin: 0; }}
        .kpi-value.highlight {{ color: var(--primary-red); }}
        .kpi-value.success {{ color: #2e7d32; }}
        .charts-grid {{ display: grid; grid-template-columns: 1fr 1fr; gap: 30px; margin-bottom: 40px; }}
        .chart-card {{ background: var(--card-bg); padding: 25px; border-radius: 12px; box-shadow: 0 5px 15px rgba(0,0,0,0.03); margin-bottom: 30px; }}
        .chart-card h3 {{ margin-top: 0; margin-bottom: 20px; color: var(--text-main); font-weight: 600; }}
        table {{ width: 100%; border-collapse: collapse; margin-top: 10px; font-size: 14px; }}
        th {{ text-align: left; padding: 12px; background-color: #f9f9f9; color: var(--text-muted); font-weight: 600; border-bottom: 2px solid #ddd; }}
        td {{ padding: 12px; border-bottom: 1px solid #eee; }}
        .badge {{ padding: 4px 8px; border-radius: 4px; font-size: 12px; font-weight: bold; }}
        .badge.bad {{ background: #ffebee; color: #b71c1c; border: 1px solid #ffcdd2; }}
        .badge.good {{ background: #e8f5e9; color: #2e7d32; border: 1px solid #c8e6c9; }}
        .filter-container {{ margin-bottom: 20px; display: flex; gap: 10px; flex-wrap: wrap; }}
        .filter-btn {{ padding: 8px 16px; border-radius: 20px; font-size: 13px; font-weight: 600; cursor: pointer; border: 1px solid #ddd; background: white; color: var(--text-main); transition: all 0.2s; }}
        .filter-btn:hover {{ background: #f0f0f0; }}
        .filter-btn.active {{ background: var(--primary-red); color: white; border-color: var(--primary-red); }}
    </style>

    <!-- Leaflet CSS and JS -->
    <link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css" />
    <script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>

</head>
<body>

    <div class="sidebar">
        <div class="sidebar-header">
            <img src="data:image/png;base64,{logo_base64}" alt="3MI Logo" style="width: 100%; height: 100%; object-fit: cover; display: {'block' if logo_base64 else 'none'};">
        </div>
        <ul class="menu">
            <li class="menu-item active" onclick="switchTab('dashboard', this)">📊 Visão Geral</li>
            <li class="menu-item" onclick="switchTab('faturamento', this)">💰 Faturamento Real (NFe)</li>
            <li class="menu-item" onclick="switchTab('propostas', this)">📈 Funil de Propostas</li>
            <li class="menu-item" onclick="switchTab('360', this)">🎯 Performance 360º</li>
            <li class="menu-item" onclick="switchTab('mapa', this)">🌍 Presença Global</li>
            <li class="menu-item" onclick="switchTab('tempera', this)">🔥 Serviços de Têmpera</li>
            <li class="menu-item" onclick="switchTab('ultimas_nfe', this)">🧾 Últimas 10 Notas Fiscais</li>
        </ul>
    </div>

    <div class="main-wrapper">
                <div class="topbar">
            <h2 id="pageTitle">Visão Geral - Inteligência Comercial</h2>
            <div style="display: flex; align-items: center; gap: 20px;">
                <div style="font-size: 14px; color: var(--text-muted); text-align: right;">Atualizado via Exportação CRM/NFe</div>
                <img src="data:image/png;base64,{logo_base64}" alt="3MI Logo" style="height: 80px; display: {'block' if logo_base64 else 'none'};">
            </div>
        </div>


        <div class="container">
            <!-- TAB 1: VISÃO GERAL -->
            <div id="tab-dashboard" class="tab-content" style="display: none;">
                <div class="kpi-grid">
                    <div class="kpi-card" style="border-top-color: #2e7d32; position: relative;">
                        <div class="kpi-title" style="display: flex; justify-content: space-between; align-items: center;">
                            Faturamento (Materiais + Serviços)
                            <span style="cursor: pointer; font-size: 16px;" onclick="toggleVisibility('valFaturamento', this)">👁️</span>
                        </div>
                        <p class="kpi-value success" id="valFaturamento" data-value="{faturamento_str}">R$ •••••••</p>
                    </div>
                    <div class="kpi-card">
                        <div class="kpi-title">Total de Clientes Mapeados</div>
                        <p class="kpi-value">{total_clientes}</p>
                    </div>
                    <div class="kpi-card" style="border-top-color: #ff9800; position: relative;">
                        <div class="kpi-title" style="display: flex; justify-content: space-between; align-items: center;">
                            Pipeline Aberto (Dinheiro na Mesa)
                            <span style="cursor: pointer; font-size: 16px;" onclick="toggleVisibility('valPipeline', this)">👁️</span>
                        </div>
                        <p class="kpi-value highlight" style="color:#ff9800;" id="valPipeline" data-value="{valor_pipeline_str}">R$ •••••••</p>
                    </div>
                </div>

                <div class="charts-grid">
                    <div class="chart-card">
                        <h3>Top 10 Representantes (Volume de Clientes)</h3>
                        <canvas id="barChart" height="200"></canvas>
                    </div>
                    <div class="chart-card">
                        <h3>Top 5 Estados (Clientes)</h3>
                        <canvas id="pieChart" height="200"></canvas>
                    </div>
                </div>

                  <div class="charts-grid" style="grid-template-columns: 1fr; margin-top: 30px;">
                      <div class="chart-card">
                          <h3>Distribuição por Natureza da Operação (Top 5 CFOPs)</h3>
                          <canvas id="cfopPieChart" height="60"></canvas>
                      </div>
                  </div>
            </div>
            
            <!-- TAB 6: FATURAMENTO REAL -->
            <div id="tab-faturamento" class="tab-content" style="display: none;">
                  <div class="filter-container" style="margin-bottom: 20px;">
                      <button class="filter-btn active" onclick="updateCFOPChart('clientes', this)">Top Clientes</button>
                      <button class="filter-btn" onclick="updateCFOPChart('mes', this)">Mês a Mês</button>
                      <button class="filter-btn" onclick="updateCFOPChart('ano', this)">Ano a Ano</button>
                  </div>
                  <div class="kpi-grid" style="grid-template-columns: 2fr 1fr;">
                      <div class="chart-card">
                          <h3 id="cfopBarTitle">Faturamento</h3>
                          <div id="containerBarFat" style="height: 350px; display: grid;">
                              <canvas id="barFatChart"></canvas>
                          </div>
                          <div id="containerLineFatAno" style="height: 350px; display: none;">
                              <canvas id="lineFatAnoChart"></canvas>
                          </div>
                      </div>
                      <div class="chart-card" style="max-height: 400px; overflow-y: auto;">
                          <h3>Top 20 Clientes (Faturamento R$)</h3>
                          <table class="data-table">
                              <thead>
                                  <tr><th>Cliente</th><th>Valor (R$)</th></tr>
                              </thead>
                              <tbody id="cfop-table-body">
                              </tbody>
                          </table>
                      </div>
                </div>
            </div>

            <!-- TAB: PROPOSTAS -->
            <div id="tab-propostas" class="tab-content" style="display: none;">
                                  <div class="filter-container" style="margin-bottom: 20px;">
                      <button class="filter-btn active" onclick="updateSellerChart('vendedores', this)">Top Vendedores</button>
                      <button class="filter-btn" onclick="updateSellerChart('mes', this)">Mês a Mês</button>
                      <button class="filter-btn" onclick="updateSellerChart('ano', this)">Ano a Ano</button>
                  </div>
                <div class="kpi-grid" style="grid-template-columns: 2fr 1fr;">
                                          <div class="chart-card">
                          <h3 id="sellerBarTitle">Top 10 Vendedores (Receita Convertida R$)</h3>
                          <div id="containerBarVendas" style="height: 350px; display: grid;">
                              <canvas id="barVendasChart"></canvas>
                          </div>
                          <div id="containerLineAno" style="height: 350px; display: none;">
                              <canvas id="lineSellerAnoChart"></canvas>
                          </div>
                      </div>
                                          <div class="chart-card">
                          <h3>Propostas em Aberto (Maiores Valores)</h3>
                          <table class="data-table">
                              <thead>
                                  <tr><th>Cliente</th><th>Valor (R$)</th><th>Vendedor</th><th>Status</th></tr>
                              </thead>
                              <tbody id="seller-table-body">
                              </tbody>
                          </table>
                      </div>
                </div>
            </div>

            <!-- TAB: MAPA -->
            <div id="tab-mapa" class="tab-content" style="display: none;">
                <div class="section-title">Presença Global (Inteligência Comercial)</div>
                <div class="chart-card" style="padding: 10px;">
                    <div style="margin-bottom: 10px; display: flex; align-items: center; gap: 10px;">
                        <label for="mapClientFilter" style="font-weight: bold;">Pesquisar Empresa no Mapa:</label>
                        <select id="mapClientFilter" style="padding: 5px; flex: 1; max-width: 400px; border-radius: 4px; border: 1px solid #ccc;">
                            <option value="">Todas as Empresas</option>
                        </select>
                    </div>
                    <div id="map_container" style="height: 600px; width: 100%; z-index:1;"></div>
                </div>
            </div>

            <!-- TAB: TEMPERA -->
            <div id="tab-tempera" class="tab-content active" style="display: block;">
                <div class="filter-container" style="margin-bottom: 20px;">
                    <label for="temperaYearFilter" style="font-weight: bold; margin-right: 10px;">Filtrar Ano:</label>
                    <select id="temperaYearFilter" onchange="updateTemperaYear(this.value)" style="padding: 5px; border-radius: 4px; border: 1px solid #ccc; font-size: 14px;">
                        {tempera_year_options_html}
                    </select>
                </div>
                <div class="kpi-grid" style="grid-template-columns: 2fr 1fr;">
                    <div class="chart-card">
                        <h3>Evolução de Faturamento</h3>
                        <div style="height: 350px;">
                            <canvas id="temperaRevChart"></canvas>
                        </div>
                    </div>
                    <div class="chart-card">
                        <h3>Top 5 Clientes</h3>
                        <div style="height: 350px; display: flex; justify-content: center; align-items: center;">
                            <canvas id="temperaClientsPieChart"></canvas>
                        </div>
                    </div>

                    <div class="chart-card" style="margin-top: 20px; grid-column: 1 / -1;">
                        <h3>Notas Fiscais do Ano</h3>
                        <div style="max-height: 400px; overflow-y: auto;">
                            <table class="data-table">
                                <thead>
                                    <tr><th>Data Emissão</th><th>Nota Fiscal</th><th>Cliente</th><th>Valor</th></tr>
                                </thead>
                                <tbody id="tempera-table-body">
                                </tbody>
                            </table>
                        </div>
                    </div>
                </div>
            </div>

            <!-- TAB: POTENCIAIS CLIENTES -->
            <div id="tab-pot" class="tab-content" style="display: none;">
                <div class="section-title">Lista de Potenciais Clientes (Inteligência Comercial)</div>
                <div class="chart-container" style="overflow-x: auto;">
                    <table class="data-table">
                        <thead>
                            <tr>
                                <th>Razão Social / Fantasia</th>
                                <th>CNAE Principal</th>
                                <th>Segmento</th>
                                <th>Localização</th>
                                <th>Site</th>
                                <th>Grau de Afinidade</th>
                            </tr>
                        </thead>
                        <tbody>
                            {pot_table_rows}
                        </tbody>
                    </table>
                </div>
            </div>

            <!-- TAB 6: PERFORMANCE 360 -->
            <div id="tab-360" class="tab-content" style="display: none;">
                <div class="charts-grid" style="grid-template-columns: 1fr;">
                    <div class="chart-card">
                        <h3>Produtividade vs Faturamento (Top 15 CRM)</h3>
                        <canvas id="perf360Chart" height="80"></canvas>
                    </div>
                </div>

                <div class="table-container" style="margin-top: 20px;">
                    <h3>Detalhamento por Vendedor</h3>
                    <table>
                        <thead>
                            <tr>
                                <th>Vendedor</th>
                                <th>Faturamento (NFe)</th>
                                <th>Interações CRM</th>
                                <th>Cotações (Propostas)</th>
                                <th>Novos Cadastros</th>
                                <th>Receita / Cotação</th>
                                <th>Receita / Interação (Eficiência)</th>
                            </tr>
                        </thead>
                        <tbody>
                            {rows_360}
                        </tbody>
                    </table>
                </div>
            </div>


            <!-- TAB: ULTIMAS 10 NFE -->
            <div id="tab-ultimas_nfe" class="tab-content" style="display: none;">
                <div class="kpi-card" style="margin-bottom: 20px;">
                    <h3>🧾 Últimas 10 Notas Fiscais Registradas</h3>
                    <div class="table-container">
                        {table_ultimas_nfe}
                    </div>
                </div>
            </div>

        </div>
    </div>

    <script>

        function switchTab(tabId, element) {{
            if(tabId === 'mapa' && typeof initMap === 'function') {{
                setTimeout(initMap, 100);
            }}
            
            document.querySelectorAll('.tab-content').forEach(tab => tab.classList.remove('active'));
            document.querySelectorAll('.menu-item').forEach(item => item.classList.remove('active'));
            
            document.getElementById('tab-' + tabId).classList.add('active');
            if (element) {{
                element.classList.add('active');
                
                // Limpar emojis do titulo removendo o primeiro caractere que e o emoji
                let cleanText = element.innerText;
                if(cleanText.length > 2) {{
                    cleanText = cleanText.substring(2).trim();
                }}
                document.getElementById('pageTitle').innerText = cleanText + " - Inteligência Comercial";
            }}
            
            if(tabId === 'faturamento') {{
                setTimeout(function(){{ filterCFOP('TODOS'); }}, 100);
            }}
            if(tabId === 'propostas') {{
                setTimeout(function(){{ filterVendedor('TODOS'); }}, 100);
            }}
            if(tabId === 'tempera') {{
                setTimeout(function(){{ 
                    if(temperaChartObj) temperaChartObj.resize();
                    if(window.temperaPieChartObj) window.temperaPieChartObj.resize();
                }}, 100);
            }}
        }}

        const curFormat = {{
            callback: function(value) {{
                if(value >= 1000000) return 'R$ ' + (value/1000000).toFixed(1) + 'M';
                if(value >= 1000) return 'R$ ' + (value/1000).toFixed(0) + 'k';
                return 'R$ ' + value;
            }}
        }};

        // Static Charts Tab 1
        new Chart(document.getElementById('barChart').getContext('2d'), {{
            type: 'bar',
            data: {{
                labels: {json.dumps(bar_labels)},
                datasets: [{{ label: 'Número de Clientes na Carteira', data: {json.dumps(bar_data)}, backgroundColor: '#A6192E', borderRadius: 6 }}]
            }},
            options: {{ responsive: true, plugins: {{ legend: {{ display: true, position: 'top' }} }} }}
        }});

        new Chart(document.getElementById('pieChart').getContext('2d'), {{
            type: 'doughnut',
            data: {{
                labels: {json.dumps(estados_labels)},
                datasets: [{{ data: {json.dumps(estados_data)}, backgroundColor: ['#A6192E', '#333333', '#777777', '#c0392b', '#e74c3c'], borderWidth: 0 }}]
            }},
            options: {{ responsive: true, plugins: {{ legend: {{ position: 'bottom' }} }}, cutout: '70%' }}
        }});
        

        new Chart(document.getElementById('cfopPieChart').getContext('2d'), {{
            type: 'bar',
            data: {{
                labels: {json.dumps(top_cfop_labels)},
                datasets: [{{ label: 'Qtd de Notas Fiscais', data: {json.dumps(top_cfop_data)}, backgroundColor: ['#A6192E', '#333333', '#777777', '#c0392b', '#e74c3c'], borderRadius: 6 }}]
            }},
            options: {{ responsive: true, plugins: {{ legend: {{ display: true, position: 'top' }} }}, indexAxis: 'y' }}
        }});

        // --- INTERACTIVE CFOP LOGIC ---
        const cfopData = {cfop_json_str};
        let lineFatAnoChart = null;
        let barFatChart = null;
        
        
        let currentCFOPPeriod = 'Mensal';
        
        function updateCFOPChart(period, btn) {{
            document.querySelectorAll('#tab-faturamento .filter-btn').forEach(b => b.classList.remove('active'));
            if(btn) btn.classList.add('active');
            
            let c1 = document.getElementById('containerBarFat');
            let c2 = document.getElementById('containerLineFatAno');
            
            if(period === 'clientes') {{
                if(c1) c1.style.display = 'grid';
                if(c2) c2.style.display = 'none';
                let t = document.getElementById('cfopBarTitle'); if(t) t.innerText = 'Top 10 Clientes (Faturamento R$)';
                // Find active CFOP category and re-render the bar chart
                let activeCat = 'TODOS';
                document.querySelectorAll('#cfop-filters .filter-btn').forEach(b => {{
                    if(b.classList.contains('active')) {{
                        if(b.innerText.includes('Tudo')) activeCat = 'TODOS';
                        else activeCat = b.innerText;
                    }}
                }});
                filterCFOP(activeCat);
            }} else {{
                if(c1) c1.style.display = 'none';
                if(c2) c2.style.display = 'grid';
                let t = document.getElementById('cfopBarTitle'); if(t) t.innerText = 'Evolução de Faturamento';
                currentCFOPPeriod = period === 'mes' ? 'Mensal' : 'Anual';
                renderCFOPChart();
            }}
        }}
        
        function renderCFOPChart() {{
            let activeCat = 'TODOS';
            document.querySelectorAll('#cfop-filters .filter-btn').forEach(btn => {{
                if(btn.classList.contains('active')) {{
                    if(btn.innerText.includes('Tudo')) activeCat = 'TODOS';
                    else activeCat = btn.innerText;
                }}
            }});
            const data = cfopData[activeCat];
            if(!data) return;
            
            const lineEl = document.getElementById('lineFatAnoChart'); if(!lineEl) return; const ctxLine = lineEl.getContext('2d');
            if(lineFatAnoChart) lineFatAnoChart.destroy();
            
            let labels = currentCFOPPeriod === 'Mensal' ? data.mes_labels : data.ano_labels;
            let chartData = currentCFOPPeriod === 'Mensal' ? data.mes_data : data.ano_data;
            
            lineFatAnoChart = new Chart(ctxLine, {{
                type: 'line',
                data: {{ labels: labels, datasets: [{{ label: 'Valor (R$)', data: chartData, borderColor: '#f39c12', backgroundColor: 'rgba(243, 156, 18, 0.1)', borderWidth: 3, fill: true, tension: 0.3 }}] }},
                options: {{ responsive: true, plugins: {{ legend: {{ display: true, position: 'top' }} }}, scales: {{ y: {{ ticks: curFormat }} }} }}
            }});
        }}
        
        function updateSellerChart(period, btn) {{
            document.querySelectorAll('#tab-propostas .filter-btn').forEach(b => b.classList.remove('active'));
            if(btn) btn.classList.add('active');
            
            let c1 = document.getElementById('containerBarVendas');
            let c2 = document.getElementById('containerLineAno');
            
            if(period === 'vendedores') {{
                if(c1) c1.style.display = 'grid';
                if(c2) c2.style.display = 'none';
                filterVendedor('TODOS');
            }} else {{
                if(c1) c1.style.display = 'none';
                if(c2) c2.style.display = 'grid';
                currentSeller = 'TODOS';
                currentSellerPeriod = (period === 'mes') ? 'Mensal' : 'Anual';
                renderSellerChart();
            }}
        }}

        function filterCFOP(categoria) {{
            document.querySelectorAll('#cfop-filters .filter-btn').forEach(btn => {{
                btn.classList.remove('active');
                if(btn.innerText.includes(categoria) || (categoria === 'TODOS' && btn.innerText.includes('Tudo'))) {{
                    btn.classList.add('active');
                }}
            }});
            const data = cfopData[categoria];
            if(!data) return;
            
            const cfopTotalEl = document.getElementById('cfop-total');
            if(cfopTotalEl) {{ cfopTotalEl.setAttribute('data-value', data.total_str); cfopTotalEl.innerText = data.total_str; }}
            
            const cfopTb = document.getElementById('cfop-table-body'); if(cfopTb) cfopTb.innerHTML = data.table_rows;
            
            // Draw bar chart
            const barEl = document.getElementById('barFatChart'); if(barEl) {{ const ctxBar = barEl.getContext('2d');
            if(barFatChart) barFatChart.destroy();
            barFatChart = new Chart(ctxBar, {{
                type: 'bar',
                data: {{ labels: data.cli_labels, datasets: [{{ label: 'Valor (R$)', data: data.cli_data, backgroundColor: '#0288d1', borderRadius: 6 }}] }},
                options: {{ responsive: true, plugins: {{ legend: {{ display: true, position: 'top' }} }}, scales: {{ y: {{ ticks: curFormat }} }} }}
            }});
            }}
            
            // Update line chart
            renderCFOPChart();
        }}
        filterCFOP('TODOS');

        // --- INTERACTIVE SELLER LOGIC ---
        const sellerData = {seller_json_str};
        let lineSellerAnoChart = null;
        let barVendasChart = null;
        let currentSellerPeriod = 'Mensal';
        let currentSeller = 'TODOS';
        
        function setSellerPeriod(period) {{
            currentSellerPeriod = period;
            document.querySelectorAll('#period-filters .filter-btn').forEach(btn => {{
                btn.classList.remove('active');
                if(btn.innerText === period) btn.classList.add('active');
            }});
            renderSellerChart();
        }}
        
        function renderSellerChart() {{
            const data = sellerData[currentSeller];
            if(!data) return;
            
            const lineSEl = document.getElementById('lineSellerAnoChart'); if(!lineSEl) return; const ctxLineS = lineSEl.getContext('2d');
            if(lineSellerAnoChart) lineSellerAnoChart.destroy();
            
            let labels = [];
            let chartData = [];
            let type = 'line';
            let bg = 'rgba(39, 174, 96, 0.1)';
            
            if(currentSellerPeriod === 'Mensal') {{
                labels = data.mes_labels;
                chartData = data.mes_data;
            }} else if (currentSellerPeriod === 'Anual') {{
                labels = data.ano_labels;
                chartData = data.ano_data;
            }}
            
            lineSellerAnoChart = new Chart(ctxLineS, {{
                type: type,
                data: {{ labels: labels, datasets: [{{ label: 'Valor (R$)', data: chartData, borderColor: '#27ae60', backgroundColor: bg, borderWidth: 3, fill: true, tension: 0.3 }}] }},
                options: {{ responsive: true, plugins: {{ legend: {{ display: true, position: 'top' }} }}, scales: {{ y: {{ ticks: curFormat }} }} }}
            }});
        }}
        
        function filterVendedor(nomeVendedor) {{
            currentSeller = nomeVendedor;
            document.querySelectorAll('#seller-filters .filter-btn').forEach(btn => {{
                btn.classList.remove('active');
                const btnText = btn.innerText;
                if(nomeVendedor === 'TODOS' && btnText.includes('Geral')) btn.classList.add('active');
                else if(btnText.toUpperCase() === nomeVendedor.toUpperCase()) btn.classList.add('active');
            }});
            const data = sellerData[nomeVendedor];
            if(!data) return;
            
            const sellerTotalEl = document.getElementById('seller-total');
              if(sellerTotalEl) {{ sellerTotalEl.setAttribute('data-value', data.total_str); sellerTotalEl.innerText = data.total_str; }}
            const selTb = document.getElementById('seller-table-body'); if(selTb) selTb.innerHTML = data.table_rows;
            
            if (currentSeller === 'TODOS') {{
                let c1 = document.getElementById('containerBarVendas'); if(c1) c1.style.display = 'grid';
                let c2 = document.getElementById('containerLineAno'); if(c2) c2.style.display = 'none';
            }} else {{
                let c3 = document.getElementById('containerBarVendas'); if(c3) c3.style.display = 'none';
                let c4 = document.getElementById('containerLineAno'); if(c4) c4.style.display = 'grid';
                renderSellerChart();
            }}
            
            let c5 = document.getElementById('sellerBarTitle'); if(c5) c5.innerText = nomeVendedor === 'TODOS' ? 'Top 10 Vendedores (Receita Convertida R$)' : `Receita Convertida R$ - ${{nomeVendedor}}`;
            
            const bvs = document.getElementById('barVendasChart'); if(!bvs) return; const ctxBarS = bvs.getContext('2d');
            if(barVendasChart) barVendasChart.destroy();
            barVendasChart = new Chart(ctxBarS, {{
                type: 'bar',
                data: {{ labels: data.bar_labels, datasets: [{{ label: 'Receita Gerada (R$)', data: data.bar_data, backgroundColor: '#2e7d32', borderRadius: 6 }}] }},
                options: {{ responsive: true, plugins: {{ legend: {{ display: true, position: 'top' }} }}, scales: {{ y: {{ ticks: curFormat }} }} }}
            }});
        }}
        filterVendedor('TODOS');
        

        function toggleVisibility(id, iconElement) {{
            const el = document.getElementById(id);
            if(el.innerText.includes('•') || el.innerText.includes('?')) {{
                el.innerText = el.getAttribute('data-value');
                iconElement.innerText = 'Ocultar';
            }} else {{
                el.innerText = 'R$ •••••••';
                iconElement.innerText = 'Mostrar';
            }}
        }}


        // --- PERFORMANCE 360 CHART ---
        const perfData = {perf360_json_str};
        const ctxPerf = document.getElementById('perf360Chart').getContext('2d');
        new Chart(ctxPerf, {{
            type: 'bar',
            data: {{
                labels: perfData.labels,
                datasets: [
                    {{
                        type: 'bar',
                        label: 'Interações (CRM)',
                        data: perfData.interacoes,
                        backgroundColor: '#3498db',
                        yAxisID: 'y'
                    }},
                    {{
                        type: 'line',
                        label: 'Faturamento NFe (R$)',
                        data: perfData.faturamento,
                        borderColor: '#2ecc71',
                        backgroundColor: 'transparent',
                        borderWidth: 3,
                        yAxisID: 'y1'
                    }}
                ]
            }},
            options: {{
                responsive: true,
                scales: {{
                    y: {{ type: 'linear', display: true, position: 'left', title: {{display: true, text: 'Qtd Interações'}} }},
                    y1: {{ type: 'linear', display: true, position: 'right', ticks: curFormat, grid: {{drawOnChartArea: false}}, title: {{display: true, text: 'Faturamento (R$)'}} }}
                }}
            }}
        }});
        
    
        
        // MAP INITIALIZATION
        let mapInitialized = false;
        function initMap() {{
            if (mapInitialized) return;
            
            var mapContainer = document.getElementById('map_container');
            if (!mapContainer) return;
            
            var map = L.map('map_container').setView([-15.7801, -47.9292], 4);
            L.tileLayer('https://{{s}}.basemaps.cartocdn.com/rastertiles/voyager/{{z}}/{{x}}/{{y}}{{r}}.png', {{
                attribution: '&copy; OpenStreetMap &copy; CARTO',
                subdomains: 'abcd',
                maxZoom: 20
            }}).addTo(map);

            var mapData = {map_json_str};
            
            var markers = L.featureGroup();
            
            mapData.forEach(function(point) {{
                var isLead = point.tipo === 'lead';
                var fillColor = isLead ? "#007bff" : "#A6192E";
                var titleColor = isLead ? "#0056b3" : "#A6192E";
                var typeLabel = isLead ? "LEAD" : "CLIENTE";
                var marker = L.circleMarker([point.lat, point.lon], {{
                    radius: isLead ? 5 : 7,
                    fillColor: fillColor,
                    color: "#fff",
                    weight: 1,
                    opacity: 1,
                    fillOpacity: 0.8
                }});
                
                var popupContent = "<div style='font-family:Inter,sans-serif;'>" +
                                   "<div style='font-size:10px; font-weight:bold; color:white; background:" + fillColor + "; padding:2px 5px; border-radius:3px; display:inline-block; margin-bottom:5px;'>" + typeLabel + "</div><br/>" +
                                   "<strong style='color:" + titleColor + "; font-size:14px;'>" + point.nome + "</strong><br/>" +
                                   "<span style='color:#666;'>" + point.cidade + " - " + point.estado + "</span><br/>" +
                                   "</div>";
                
                marker.bindTooltip(popupContent, {{direction: 'top', opacity: 0.95}});
                markers.addLayer(marker);
            }});
            
            markers.addTo(map);
            
            if(mapData.length > 0) {{
                map.fitBounds(markers.getBounds(), {{padding: [50, 50]}});
            }}
            
            var clientSelect = document.getElementById('mapClientFilter');
            if(clientSelect) {{
                var clientNames = [...new Set(mapData.map(item => item.cliente))].sort();
                clientNames.forEach(name => {{
                    var option = document.createElement('option');
                    option.value = name;
                    option.textContent = name;
                    clientSelect.appendChild(option);
                }});
                
                clientSelect.addEventListener('change', function(e) {{
                    var selectedClient = e.target.value;
                    markers.clearLayers();
                    
                    var filteredData = selectedClient ? mapData.filter(p => p.cliente === selectedClient) : mapData;
                    
                    filteredData.forEach(function(point) {{
                        var isLead = point.tipo === 'lead';
                        var fillColor = isLead ? "#007bff" : "#A6192E";
                        var titleColor = isLead ? "#0056b3" : "#A6192E";
                        var typeLabel = isLead ? "LEAD" : "CLIENTE";
                        var marker = L.circleMarker([point.lat, point.lon], {{
                            radius: isLead ? 5 : 7,
                            fillColor: fillColor,
                            color: "#fff",
                            weight: 1,
                            opacity: 1,
                            fillOpacity: 0.8
                        }});
                        
                        var valFormatted = "R$ " + (point.faturamento||0).toLocaleString('pt-BR', {{minimumFractionDigits: 2}});
                        var popupContent = "<div style='font-family:sans-serif; min-width:200px;'>" +
                                           "<div style='background-color:" + titleColor + "; color:white; padding:5px; border-radius:4px; margin-bottom:8px; font-weight:bold; font-size:12px; text-align:center;'>" + typeLabel + "</div>" +
                                           "<strong style='color:#333; font-size:14px;'>" + point.cliente + "</strong><br/>" +
                                           "<span style='color:#666;'>" + point.cidade + " - " + point.estado + "</span><br/>" +
                                           "</div>";
                        
                        marker.bindTooltip(popupContent, {{direction: 'top', opacity: 0.95}});
                        markers.addLayer(marker);
                    }});
                    
                    if (filteredData.length > 0) {{
                        if (selectedClient) {{
                            map.fitBounds(markers.getBounds(), {{padding: [50, 50], maxZoom: 12}});
                        }} else {{
                            map.fitBounds(markers.getBounds(), {{padding: [50, 50]}});
                        }}
                    }}
                }});
            }}
            
            mapInitialized = true;
        }}


        // CHART: TEMPERA
        let temperaChartObj = null;
        let temperaMonthlyDict = {json.dumps(tempera_monthly_dict)};
        let temperaEnergyDict = {json.dumps(energy_costs_dict)};
          let temperaYearlyTables = {json.dumps(tempera_yearly_tables_dict)};
        let temperaMonthLabels = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez'];

        if(document.getElementById('temperaRevChart')) {{
            var initialYearEl = document.getElementById('temperaYearFilter');
            var initialYear = initialYearEl ? initialYearEl.value : null;
            var initialData = (initialYear && temperaMonthlyDict[initialYear]) ? temperaMonthlyDict[initialYear] : [0,0,0,0,0,0,0,0,0,0,0,0];
              const tb = document.getElementById('tempera-table-body');
              if(tb && initialYear && temperaYearlyTables[initialYear]) {{ tb.innerHTML = temperaYearlyTables[initialYear]; }}
            
            temperaChartObj = new Chart(document.getElementById('temperaRevChart').getContext('2d'), {{
                type: 'line',
                data: {{
                    labels: temperaMonthLabels,
                    datasets: [
                          {{
                              label: 'Custo de Energia (R$)',
                              data: (initialYear && temperaEnergyDict[initialYear]) ? temperaEnergyDict[initialYear] : [0,0,0,0,0,0,0,0,0,0,0,0],
                              borderColor: '#333333',
                              backgroundColor: 'rgba(51, 51, 51, 0.1)',
                              borderWidth: 2,
                              fill: true,
                              tension: 0.3
                          }},
                          {{
                              label: 'Faturamento (R$)',
                              data: initialData,
                              borderColor: '#A6192E',
                              backgroundColor: 'rgba(166, 25, 46, 0.1)',
                              borderWidth: 2,
                              fill: true,
                              tension: 0.3
                          }}
                      ]
                }},
                options: {{
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {{ legend: {{ display: true, position: 'top' }} }},
                    scales: {{ 
                        y: {{ 
                            beginAtZero: true,
                            ticks: {{
                                callback: function(value) {{ return 'R$ ' + (value/1000).toLocaleString() + 'k'; }}
                            }}
                        }} 
                    }}
                }}
            }});
        }}

        if(document.getElementById('temperaClientsPieChart')) {{
            window.temperaPieChartObj = new Chart(document.getElementById('temperaClientsPieChart').getContext('2d'), {{
                type: 'doughnut',
                data: {{
                    labels: {json.dumps(tempera_clients_labels)},
                    datasets: [{{
                        label: 'Faturamento',
                        data: {json.dumps(tempera_clients_data)},
                        backgroundColor: ['#A6192E', '#2e7d32', '#fbc02d', '#1565c0', '#e65100']
                    }}]
                }},
                options: {{
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {{
                        legend: {{ position: 'bottom', labels: {{ boxWidth: 12, font: {{ size: 10 }} }} }}
                    }}
                }}
            }});
        }}

        window.updateTemperaYear = function(ano) {{
              if (!temperaChartObj || !temperaMonthlyDict[ano]) return;
              if (temperaEnergyDict[ano]) {{
                  temperaChartObj.data.datasets[0].data = temperaEnergyDict[ano];
              }} else {{
                  temperaChartObj.data.datasets[0].data = [0,0,0,0,0,0,0,0,0,0,0,0];
              }}
              temperaChartObj.data.datasets[1].data = temperaMonthlyDict[ano];
              temperaChartObj.update();
              
              const tb = document.getElementById('tempera-table-body');
              if(tb && temperaYearlyTables[ano]) {{
                  tb.innerHTML = temperaYearlyTables[ano];
              }}
          }};

    </script>
</body>
</html>
"""

    output_path = r"index.html"
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html_content)
        
    print(f"HTML Dashboard gerado em: {output_path}")

if __name__ == "__main__":
    generate_dashboard()

print("Reached the end of the script! name is:", __name__)
