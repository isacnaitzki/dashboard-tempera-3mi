import requests
import pandas as pd

def scrape_pedidos():
    url = "http://177.200.204.3:8090/r3/_connect.php"
    
    headers = {
        'Accept': '*/*',
        'Accept-Language': 'pt-BR,pt;q=0.9,en-US;q=0.8,en;q=0.7',
        'Connection': 'keep-alive',
        'Content-Type': 'application/x-www-form-urlencoded',
        'Cookie': 'usuarioid=PATRICK; id=47; dbpath=r3; senha=BOB',
        'Origin': 'http://177.200.204.3:8090',
        'Referer': 'http://177.200.204.3:8090/r3/_sistema.php',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
    }

    # Added &search=preventiva and limit=1000000
    payload = "codvalido=-1c-1c72c0c_uj8_2026_08_11_09_22_18_&route=_carregaplan.php&wr=r&tabelas=pedidos%2C%20perempinf%2C%20empresas%2C%20tipoexecucao%2C%20pedidosstatus%2C%20nfe%2C%20perrepresent&filtros=&searchfiltros=&listaRegistro=12&distinct=&linhasPorPagina=1000000&wheres=&searchcols=pedidos.id%5Enp%5E~CASE%20WHEN%20nfe.nnf%20IS%20NULL%20THEN%20NULL%20ELSE%20CONCATrpapapmnfe.nnf%3A%3Atext%2C%20rataptm-rataptm%2C%20nfe.serierpfppfm%20END%5Et%5E~CONCATrpapapmempresas.id%2C%20rataptm%20-%20rataptm%2C%20empresas.nomerpfppfm%5Et%5E~CONCATrpapapmempresas.id%2C%20rataptm%20-%20rataptm%2C%20empresas.razaorpfppfm%5Et%5E~pedidos.descricao%5Etx%5E~pedidos.nserie%5Eta%5E&search=preventiva&ordem=%20ORDER%20BY%20pedidos.id%20DESC%20NULLS%20LAST&wheremais=&veriUltTotReg=&tabctrrefresh=&colunas=pedidos.dtcadarpcmcperempinf.nomearpcmcperempinf.razaoarpcmcperempinf.cpfcnpjarpcmcpedidos.idarpcmcpedidos.propostatarpcmcpedidos.gantregarpcmcempresas.nomearpcmcempresas.razaoarpcmctipoexecucao.descricaoarpcmcpedidosstatus.statusarpcmcpedidos.descricaoarpcmcpedidos.entregaarpcmcpedidos.nseriearpcmcnfe.nffakearpcmcnfe.statusarpcmcperrepresent.nomearpcmcpedidos.valorosarpcmcpedidos.valorofarpcmcpedidos.vlfretearpcmcpedidos.totalarpcmcpedidos.recebidoarpcmcpedidos.arecebido&tipos=d%7Ct%7Ct%7Ct%7Cnp%7Ct%7Cb%7Ct%7Ct%7Ct%7Ct%7Ctx%7Cd%7Cta%7Ct%7Ct%7Ct%7Cf%7Cf%7Cf%7Cf%7Cf%7Cf&sobrecol=arpcmcarpcmcarpcmcarpcmcarpcmcarpcmcCASE%20WHEN%20rpapapmSELECT%20gantt.id%20FROM%20gantt%20WHERE%20gantt.ex%20IS%20NULL%20AND%20gantt.nivel%20riipgm%201%20AND%20gantt.of%20riipgm%20pedidos.id%20LIMIT%201rpfppfm%20IS%20NULL%20THEN%20rataptmrataptm%20ELSE%20rataptmrsmmpmimg%20srcriipgmrppspm..rbbpbm..rbbpbm_imgrbbpbmrealizado.pngrppspm%20data-titleriipgmrppspmOKrppspm%20widthriipgmrppspm16pxrppspm%20heightriipgmrppspm16pxrppspm%20rbbpbmrmmssmrataptm%20END%20AS%20perganttarpcmcCONCATrpapapmempresas.id%2C%20rataptm%20-%20rataptm%2C%20empresas.nomerpfppfm%20AS%20perclientearpcmcCONCATrpapapmempresas.id%2C%20rataptm%20-%20rataptm%2C%20empresas.razaorpfppfm%20AS%20perazaoarpcmcarpcmcarpcmcarpcmcarpcmcarpcmcCASE%20WHEN%20nfe.nnf%20IS%20NULL%20THEN%20NULL%20ELSE%20CONCATrpapapmnfe.nnf%3A%3Atext%2C%20rataptm-rataptm%2C%20nfe.serierpfppfm%20END%20AS%20fakenfearpcmcarpcmcarpcmcarpcmcarpcmcarpcmcCASE%20WHEN%20pedidos.valoros%20IS%20NULL%20AND%20pedidos.valorof%20IS%20NULL%20AND%20pedidos.vlfrete%20IS%20NULL%20THEN%20NULL%20ELSE%20COALESCErpapapmpedidos.valoros%2C0rpfppfm%20%2B%20COALESCErpapapmpedidos.valorof%2C0rpfppfm%20%2B%20COALESCErpapapmpedidos.vlfrete%2C0rpfppfm%20END%20AS%20pertotalarpcmcarpcmc&carrLinha=-1&refresh=0&groupBy=&rolagem=0&sqlWith=&sqlFrom=pedidos%20LEFT%20JOIN%20infoempresa%20ON%20infoempresa.id%20riipgm%20pedidos.r3idemp%20LEFT%20JOIN%20empresas%20AS%20perempinf%20ON%20perempinf.id%20riipgm%20infoempresa.assoccad%20LEFT%20JOIN%20pedidosstatus%20ON%20pedidos.status%20riipgm%20pedidosstatus.id%20LEFT%20JOIN%20empresas%20ON%20empresas.id%20riipgm%20pedidos.empresa%20LEFT%20JOIN%20tipoexecucao%20ON%20tipoexecucao.id%20riipgm%20pedidos.tipoexecucao%20LEFT%20JOIN%20contaspagar%20ON%20contaspagar.id%20riipgm%20pedidos.idconr%20LEFT%20JOIN%20empresas%20AS%20perrepresent%20ON%20perrepresent.id%20riipgm%20pedidos.representante%20LEFT%20JOIN%20nfe%20ON%20contaspagar.idnfe%20riipgm%20nfe.id%20AND%20nfe.ex%20IS%20NULL%20&rotina=_modrbbpbmperbbpbmpe_cad_ofs_c.php&idp=0&selLinhaPlan=0&selColPlan=1&naofoca=0&carrCol=0&transfeinf=&frame=0&bdpadrao=1&tabpri=pedidos"

    cols_names = [
        "DataCadastro", "NomeInfo", "RazaoInfo", "CNPJInfo", "PedidoID", "Proposta", 
        "Gantt", "ClienteNome", "ClienteRazao", "TipoExecucao", "StatusPedido", 
        "Descricao", "DataEntrega", "NumSerie", "NFe", "NFeStatus", 
        "Representante", "ValorOS", "ValorOF", "ValorFrete", "Total", 
        "Recebido", "AReceber"
    ]
    parsed_data = []

    print("Fetching Pedidos de Manutenção Preventiva...")
    
    try:
        response = requests.post(url, headers=headers, data=payload, timeout=60)
        if response.status_code != 200:
            print("Error fetching:", response.status_code)
            return
    except Exception as e:
        print("Erro de conexao:", e)
        return

    data_lines = response.text.split('|~||')
    if len(data_lines) <= 2:
        print("No valid data retrieved!")
        return

    for line in data_lines[2:]:
        if not line.strip():
            continue
        cells = line.split('|')
        row_data = {}
        for i, cell in enumerate(cells):
            if i < len(cols_names):
                if "^" in cell:
                    parts = cell.split('^')
                    val = parts[1] if len(parts) > 1 else parts[0]
                else:
                    val = cell
                row_data[cols_names[i]] = val
        parsed_data.append(row_data)

    if not parsed_data:
        print("Nenhum dado retornado.")
        return

    df = pd.DataFrame(parsed_data)
    
    # Exclude rows where Descricao contains 'preventivamente'
    if 'Descricao' in df.columns:
        mask = ~df['Descricao'].str.contains('preventivamente', case=False, na=False)
        df = df[mask]
    
    print(f"Total rows after filtering 'preventivamente': {len(df)}")
    
    # Clean up numerical columns
    for col in ["ValorOS", "Total"]:
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
    
    df['DataCadastro'] = pd.to_datetime(df['DataCadastro'], errors='coerce')
        
    df.to_excel("Pedidos_Manutencao.xlsx", index=False)
    print("Saved to Pedidos_Manutencao.xlsx")

if __name__ == "__main__":
    scrape_pedidos()
