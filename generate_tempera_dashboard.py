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
            <div id="tab-dashboard" class="tab-content active">
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
            <div id="tab-faturamento" class="tab-content">
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
            <div id="tab-propostas" class="tab-content">
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
            <div id="tab-mapa" class="tab-content">
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
            <div id="tab-tempera" class="tab-content">
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
            <div id="tab-pot" class="tab-content">
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
            <div id="tab-360" class="tab-content">
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
            <div id="tab-ultimas_nfe" class="tab-content">
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

    # --- ISAC'S STANDALONE EXTRACTOR & STYLIST ---
    from bs4 import BeautifulSoup
    print('Limpando outras abas, menu lateral e injetando estilo PREMIUM com LOGO MAIOR...')
    soup = BeautifulSoup(html_content, 'html.parser')
    
    # Hide sidebar
    sidebar = soup.find('div', class_='sidebar')
    if sidebar: sidebar['style'] = 'display: none !important;'
        
    menu_toggle = soup.find('button', class_='menu-toggle')
    if menu_toggle: menu_toggle['style'] = 'display: none !important;'
    
    # Process tabs
    main_content = soup.find('div', class_='container')
    if main_content:
        tabs = main_content.find_all('div', class_='tab-content')
        for tab in tabs:
            if tab.get('id') == 'tab-tempera':
                tab['class'] = [c for c in tab.get('class', []) if c != 'hidden']
                if 'active' not in tab['class']: tab['class'].append('active')
                tab['style'] = 'display: block !important;'
            else:
                tab['class'] = [c for c in tab.get('class', []) if c != 'active']
                tab['class'].append('hidden')
                tab['style'] = 'display: none !important; visibility: hidden !important; height: 0 !important; overflow: hidden !important; margin: 0 !important; padding: 0 !important;'
                
    # Update TOPBAR
    topbar = soup.find('div', class_='topbar')
    if topbar:
        h2 = topbar.find('h2', id='pageTitle')
        if h2: 
            h2.string = 'Faturamento - Serviços de tempera'
            h2['style'] = 'color: white !important; font-size: 32px !important; font-weight: 800 !important; font-style: normal !important; letter-spacing: -1px !important; text-shadow: 0 2px 10px rgba(0,0,0,0.2) !important;'
        
        # Apply premium gradient to topbar with subtle pattern
        topbar['style'] = 'background: linear-gradient(135deg, #8A1224 0%, #B81D33 100%) !important; border-bottom: none !important; box-shadow: 0 10px 30px rgba(138, 18, 36, 0.2) !important; padding: 15px 50px !important; display: flex; justify-content: space-between; align-items: center; position: relative; z-index: 10;'
        
        # Fix the text color for right side, but don't decompose the parent wrapper!
        for div in topbar.find_all('div', recursive=True):
            if div.get('style') and 'font-size: 14px' in div.get('style'):
                div.decompose() # Remove ONLY the specific text node div
                
        # Fix the image - make it much nicer!
        for img in topbar.find_all('img'):
            img['src'] = 'data:image/png;base64,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'
            img['style'] = 'height: 85px !important; width: auto !important; display: block !important; background: white !important; padding: 8px 15px !important; border-radius: 12px !important; box-shadow: 0 4px 15px rgba(0,0,0,0.15), inset 0 0 0 1px rgba(0,0,0,0.05) !important; transition: transform 0.3s ease !important;'

    # Wrap the filter in a nice floating panel
    filter_label = soup.find(string=lambda t: t and 'Filtrar Ano:' in t)
    if filter_label:
        parent_div = filter_label.parent
        if parent_div and parent_div.name == 'div':
            parent_div['style'] = 'background: white !important; padding: 15px 25px !important; border-radius: 12px !important; box-shadow: 0 4px 20px rgba(0,0,0,0.04) !important; display: inline-flex !important; align-items: center !important; gap: 15px !important; margin-bottom: 30px !important; font-weight: 600 !important; color: #334155 !important;'

    # Injecting modern styles
    style_tag = soup.new_tag('style')
    style_tag.string = '''
        @import url('https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;500;600;700;800&display=swap');
        
        :root {
            --bg-color: #f4f7f9;
            --primary-red: #A6192E;
        }
        body {
            background-color: var(--bg-color) !important;
            background-image: radial-gradient(#e2e8f0 1px, transparent 1px) !important;
            background-size: 20px 20px !important;
            font-family: 'Outfit', -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif !important;
            display: block !important;
            height: auto !important;
            overflow: auto !important;
        }
        
        .main-wrapper {
            margin-left: 0 !important;
            width: 100% !important;
            display: block !important;
        }
        
        .container { 
            margin-left: 0 !important; 
            max-width: 1500px !important; 
            padding: 40px 50px !important; 
            margin-inline: auto !important;
        }
        
        .sidebar { display: none !important; }
        
        .chart-card, .kpi-card {
            background: rgba(255, 255, 255, 0.95) !important;
            backdrop-filter: blur(10px) !important;
            border-radius: 20px !important;
            box-shadow: 0 15px 35px rgba(0,0,0,0.04), 0 5px 15px rgba(0,0,0,0.02) !important;
            border: 1px solid rgba(255,255,255,0.8) !important;
            transition: all 0.3s cubic-bezier(0.25, 0.8, 0.25, 1) !important;
            padding: 30px !important;
            border-top: 5px solid var(--primary-red) !important;
        }
        .chart-card:hover, .kpi-card:hover {
            transform: translateY(-5px) !important;
            box-shadow: 0 20px 40px rgba(166, 25, 46, 0.1), 0 8px 20px rgba(0,0,0,0.04) !important;
        }
        
        .chart-card h3, .section-title, .kpi-title {
            font-weight: 700 !important;
            color: #1e293b !important;
            margin-bottom: 25px !important;
            font-size: 1.3rem !important;
            text-transform: none !important;
            letter-spacing: -0.5px !important;
        }
        
        select, input, .filter-btn {
            border-radius: 10px !important;
            border: 1px solid #cbd5e1 !important;
            padding: 10px 20px !important;
            font-size: 15px !important;
            background-color: #f8fafc !important;
            color: #0f172a !important;
            font-weight: 600 !important;
            transition: all 0.2s !important;
            cursor: pointer;
            font-family: 'Outfit', sans-serif !important;
        }
        select:hover, input:hover, .filter-btn:hover {
            border-color: var(--primary-red) !important;
            background-color: #ffffff !important;
            box-shadow: 0 4px 10px rgba(0,0,0,0.05) !important;
        }
        select:focus, input:focus {
            outline: none !important;
            box-shadow: 0 0 0 3px rgba(166, 25, 46, 0.2) !important;
            border-color: var(--primary-red) !important;
        }
        
        /* Table styles to match the premium look */
        table {
            width: 100% !important;
            border-collapse: separate !important;
            border-spacing: 0 8px !important;
            margin-top: 20px !important;
        }
        th {
            background: transparent !important;
            color: #64748b !important;
            font-weight: 600 !important;
            font-size: 14px !important;
            text-transform: uppercase !important;
            letter-spacing: 0.5px !important;
            border: none !important;
            padding: 10px 20px !important;
        }
        td {
            background: white !important;
            padding: 16px 20px !important;
            border: none !important;
            font-size: 15px !important;
            color: #334155 !important;
            font-weight: 500 !important;
        }
        tr td:first-child { border-radius: 12px 0 0 12px !important; }
        tr td:last-child { border-radius: 0 12px 12px 0 !important; }
        tr { box-shadow: 0 2px 10px rgba(0,0,0,0.02) !important; transition: transform 0.2s !important; }
        tr:hover { transform: scale(1.01) !important; box-shadow: 0 5px 15px rgba(0,0,0,0.05) !important; }
        
        #tab-tempera .charts-grid {
            display: grid !important;
            grid-template-columns: 2fr 1fr !important;
            gap: 30px !important;
            margin-top: 10px !important;
        }
        
        @media (max-width: 1024px) {
            #tab-tempera .charts-grid {
                grid-template-columns: 1fr !important;
            }
        }
    '''
    soup.head.append(style_tag)

    html_content = str(soup)
    # -----------------------------------

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html_content)
        
    print(f"HTML Dashboard gerado em: {output_path}")

if __name__ == "__main__":
    generate_dashboard()

print("Reached the end of the script! name is:", __name__)
