import requests
import pandas as pd

def scrape_tempera():
    url = "http://177.200.204.3:8090/r3/_connect.php"

    payload = "codvalido=-1c-1c81c0c_3fo_2026_07_30_07_27_19_&route=_carregaplan.php&wr=r&tabelas=nfe%2C%20perempinf%2C%20nfetpnf%2C%20nfetipovenda%2C%20nfestatus%2C%20material%2C%20empresas%2C%20cidades%2C%20natnfecfop%2C%20nfecfop%2C%20nfeitens%2C%20pervendedor&filtros=%20AND%20rpapapm%20nfecfop.cdescricao%20IN%20rpapapmrataptm5.124%20-%20INDUSTRIALIZACAO%20EFETUADA%20PARA%20OUTRA%20EMPRESArataptm%2Crataptm6.124%20-%20INDUSTRIALIZACAO%20FEITA%20PARA%20OUTRA%20EMPRESArataptmrpfppfmrpfppfm&searchfiltros=&listaRegistro=12&distinct=&linhasPorPagina=1000000&wheres=&searchcols=material.mascara%5Et%5E~material.descricao%5Et%5E~CASE%20WHEN%20nfe.nnf%20IS%20NULL%20THEN%20rataptmrataptm%20ELSE%20CONCATrpapapmnfe.nnf%2Crataptm-rataptm%2Cnfe.serierpfppfm%20END%5Et%5E~CONCATrpapapmempresas.id%2C%20rataptm%20-%20rataptm%2C%20empresas.nomerpfppfm%5Et%5E~CONCATrpapapmempresas.id%2C%20rataptm%20-%20rataptm%2C%20empresas.razaorpfppfm%5Et%5E~nfe.chaveprotocolo%5Et%5E~material.codbarras%5Eean%5E&search=&ordem=%20ORDER%20BY%20nfecfop.cdescricao%20DESC%20NULLS%20LAST%2C%20nfe.demi%20DESC%20NULLS%20LAST%2C%20nfe.id%20DESC%20NULLS%20LAST%2C%20nfeitens.id%20DESC%20NULLS%20LAST&wheremais=%20AND%20nfeitens.nfe%20IS%20NOT%20NULL%20AND%20nfe.ex%20IS%20NULL&veriUltTotReg=&tabctrrefresh=&colunas=nfe.idarpcmcnfe.demiarpcmcperempinf.nomearpcmcperempinf.razaoarpcmcperempinf.cpfcnpjarpcmcnfetpnf.descricaoarpcmcnfetipovenda.tipoarpcmcnfe.nnffakearpcmcnfe.statusarpcmcnfestatus.statusarpcmcmaterial.mascaraarpcmcmaterial.descricaoarpcmcmaterial.codbarrasarpcmcmaterial.ncmarpcmcempresas.nomearpcmcempresas.razaoarpcmccidades.ufsiglaarpcmcnatnfecfop.cdescricaoarpcmcnfecfop.cdescricaoarpcmcnfe.chaveprotocoloarpcmcnfeitens.qcomarpcmcnfeitens.vuncomarpcmcnfeitens.vlprodarpcmcnfeitens.vbcicmsarpcmcnfeitens.picmsarpcmcnfeitens.vicmsarpcmcnfeitens.pipiarpcmcnfeitens.vipiarpcmcnfeitens.ppisarpcmcnfeitens.vpisarpcmcnfeitens.pcofinsarpcmcnfeitens.vcofinsarpcmcnfeitens.vicmsstarpcmcnfeitens.vtotalarpcmcnfeitens.vuncomfaturaarpcmcpervendedor.nome&tipos=np%7Cd%7Ct%7Ct%7Ct%7Ct%7Ct%7Ct%7Cnp%7Ct%7Ct%7Ct%7Cean%7Ct%7Ct%7Ct%7Ct%7Ct%7Ct%7Ct%7Cf%7Cf%7Cf%7Cf%7Cf%7Cf%7Cf%7Cf%7Cf%7Cf%7Cf%7Cf%7Cf%7Cf%7Cf%7Ct&sobrecol=arpcmcarpcmcarpcmcarpcmcarpcmcarpcmcarpcmcCASE%20WHEN%20nfe.nnf%20IS%20NULL%20THEN%20rataptmrataptm%20ELSE%20CONCATrpapapmnfe.nnf%2Crataptm-rataptm%2Cnfe.serierpfppfm%20END%20AS%20pernfearpcmcarpcmcarpcmcarpcmcarpcmcarpcmcCASE%20WHEN%20COALESCErpapapmnfeitens.ncm%2Crataptmrataptmrpfppfm%20riipgm%20rataptmrataptm%20THEN%20material.ncm%20ELSE%20nfeitens.ncm%20END%20AS%20pncmarpcmcCONCATrpapapmempresas.id%2C%20rataptm%20-%20rataptm%2C%20empresas.nomerpfppfm%20AS%20%20perclientearpcmcCONCATrpapapmempresas.id%2C%20rataptm%20-%20rataptm%2C%20empresas.razaorpfppfm%20AS%20perazaoarpcmcarpcmcarpcmcarpcmcarpcmcarpcmcarpcmcrpapapmrpapapmnfeitens.qcom%20*%20nfeitens.vuncomtrpfppfm%20-%20COALESCErpapapmnfeitens.tdesconto%2C0rpfppfmrpfppfm%20AS%20vlprodarpcmcarpcmcarpcmcarpcmcarpcmcarpcmcarpcmcarpcmcarpcmcarpcmcarpcmcROUNDrpapapmCASTrpapapmrpapapmrpapapmnfeitens.qcom%20*%20nfeitens.vuncomtrpfppfm%20-%20COALESCErpapapmnfeitens.tdesconto%2C0rpfppfm%20%2B%20COALESCErpapapmnfeitens.vicmsst%2C0rpfppfm%20%2B%20COALESCErpapapmnfeitens.vipi%2C0rpfppfmrpfppfm%20AS%20NUMERICrpfppfm%2C%202rpfppfm%20AS%20totalitemarpcmcarpcmc&carrLinha=-1&refresh=0&groupBy=&rolagem=0&sqlWith=&sqlFrom=nfeitens%20LEFT%20JOIN%20material%20ON%20nfeitens.prodid%20riipgm%20material.id%20LEFT%20JOIN%20nfecfop%20ON%20nfecfop.id%20riipgm%20nfeitens.idcfop%20LEFT%20JOIN%20nfe%20ON%20nfeitens.nfe%20riipgm%20nfe.id%20LEFT%20JOIN%20nfecfop%20AS%20natnfecfop%20ON%20natnfecfop.id%20riipgm%20nfe.natop%20LEFT%20JOIN%20empresas%20ON%20nfe.cliente%20riipgm%20empresas.id%20LEFT%20JOIN%20cidades%20ON%20empresas.cidade%20riipgm%20cidades.id%20LEFT%20JOIN%20nfestatus%20ON%20nfestatus.id%20riipgm%20nfe.status%20LEFT%20JOIN%20infoempresa%20ON%20infoempresa.id%20riipgm%20nfe.r3idemp%20LEFT%20JOIN%20empresas%20AS%20perempinf%20ON%20perempinf.id%20riipgm%20infoempresa.assoccad%20LEFT%20JOIN%20empresas%20AS%20pervendedor%20ON%20pervendedor.id%20riipgm%20nfe.representante%20LEFT%20JOIN%20nfetpnf%20ON%20nfecfop.tpnf%20riipgm%20nfetpnf.codigo%20LEFT%20JOIN%20nfetipovenda%20ON%20nfetipovenda.codigo%20riipgm%20nfe.nfetipovenda%20LEFT%20JOIN%20empresas%20AS%20transportadora%20ON%20transportadora.id%20riipgm%20nfe.transportadora%20LEFT%20JOIN%20nfemodfrete%20ON%20nfemodfrete.codigo%20riipgm%20nfe.modfrete%20LEFT%20JOIN%20nfemodfrete%20AS%20nfemodfretered%20ON%20nfemodfretered.codigo%20riipgm%20nfe.modredespacho%20LEFT%20JOIN%20pedidos%20ON%20pedidos.id%20riipgm%20nfe.idpedido%20&rotina=_modrbbpbmnferbbpbmnfe_ger_nfe_c.php&idp=1&selLinhaPlan=0&selColPlan=2&naofoca=0&carrCol=0&transfeinf=&frame=0&bdpadrao=1&tabpri=nfeitens"

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

    print("Fetching Tempera items...")
    response = requests.post(url, headers=headers, data=payload)
    if response.status_code != 200:
        print("Error fetching:", response.status_code)
        return

    # Parse response
    data_text = response.text
    # the response is separated by |~||
    data_lines = data_text.split('|~||')
    
    parsed_data = []
    cols = [
        "nfe_id", "DataEmissao", "perempinf_nome", "perempinf_razao", "perempinf_cpfcnpj",
        "nfetpnf_descricao", "nfetipovenda_tipo", "NumeroDaNota", "nfe_status", "StatusDesc",
        "material_mascara", "material_descricao", "material_codbarras", "material_ncm", 
        "ClienteNome", "empresas_razao", "cidades_ufsigla", "natnfecfop_cdescricao", 
        "CFOP", "nfe_chaveprotocolo", "nfeitens_qcom", "nfeitens_vuncom", "nfeitens_vlprod",
        "nfeitens_vbcicms", "nfeitens_picms", "nfeitens_vicms", "nfeitens_pipi", "nfeitens_vipi",
        "nfeitens_ppis", "nfeitens_vpis", "nfeitens_pcofins", "nfeitens_vcofins", "nfeitens_vicmsst",
        "ValorTotalNF", "nfeitens_vuncomfatura", "Vendedor"
    ]

    for i, line in enumerate(data_lines):
        if i == 0:
            # First line might contain header metadata before |~||
            # We can split by |~| or just take the end part
            if '|' in line:
                line = line.split('|')[-1]
                
        fields = line.split('|')
        
        row = []
        for field in fields:
            if '^' in field:
                val = field.split('^')[-1]
            else:
                val = field
            row.append(val)
            
        # Clean row: filter out the first column if it's empty from splitting |~||
        # Actually it's separated by | for fields
        if len(row) > 0 and row[0] == '':
            row = row[1:]

        if len(row) >= len(cols):
            parsed_data.append(row[:len(cols)])
        elif len(row) > 0:
            # pad with empty
            row += [''] * (len(cols) - len(row))
            parsed_data.append(row)

    df = pd.DataFrame(parsed_data, columns=cols)
    
    # Fix types
    try:
        df['ValorTotalNF'] = pd.to_numeric(df['ValorTotalNF'], errors='coerce').fillna(0)
        df['DataEmissao'] = pd.to_datetime(df['DataEmissao'], errors='coerce')
    except Exception as e:
        print("Warning type conv:", e)
        
    out_path = r"Tempera_Itens.xlsx"
    df.to_excel(out_path, index=False)
    print(f"Saved {len(df)} rows to {out_path}")

if __name__ == '__main__':
    scrape_tempera()
