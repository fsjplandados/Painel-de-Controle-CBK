"""
Pipeline de Consolidação de Chargebacks — SAP ↔ Adyen ↔ ClearSale
Desenvolvido para consolidar chargebacks originados do SAP com informações complementares.
"""
import os
import re
import sys
import glob
import json
import pandas as pd
import numpy as np
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed

# Garante que o console do Windows exiba caracteres acentuados corretamente
sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)


def extract_psp_ref(text):
    """
    Extrai o código de 16 caracteres alfanuméricos (Psp Reference) do texto.
    Exemplo: 'CHARGEBACK NSU CR8WJ3L83T24KLF6-18056' -> 'CR8WJ3L83T24KLF6'
    """
    if pd.isna(text):
        return None
    match = re.search(r'[A-Z0-9]{16}', str(text).upper())
    return match.group(0) if match else None


def extract_metadata(metadata_str):
    """
    Extrai orderId, paymentId e transactionId do campo Metadata JSON da Adyen.
    """
    if pd.isna(metadata_str) or not isinstance(metadata_str, str):
        return None, None, None
    try:
        data = json.loads(metadata_str)
        return (
            data.get("orderId"),
            data.get("paymentId"),
            data.get("transactionId")
        )
    except Exception:
        return None, None, None


def load_sap_data(sap_folder):
    """
    Carrega todos os arquivos Excel na pasta SAP e filtra por CHARGEBACK na Denominação.
    """
    print("\n[1/4] Carregando dados do SAP...")
    files = glob.glob(os.path.join(sap_folder, "*.xlsx"))
    files = [f for f in files if not os.path.basename(f).startswith("~$")]
    
    if not files:
        raise FileNotFoundError(f"Nenhum arquivo Excel (.xlsx) encontrado em: {sap_folder}")
    
    dfs = []
    for f in files:
        print(f"  → Lendo SAP: {os.path.basename(f)}")
        try:
            df = pd.read_excel(f)
            print(f"    - Total bruto: {len(df)} linhas")
            
            # Identificar coluna Denominação (independente de acentuação/case)
            denom_col = next((c for c in df.columns if "Denomina" in c and "objeto" not in c), None)
            
            if not denom_col:
                print(f"    ⚠️ Coluna 'Denominação' não encontrada em {os.path.basename(f)}. Pulando arquivo.")
                continue
                
            # Filtrar apenas linhas contendo CHARGEBACK na Denominação
            mask = df[denom_col].astype(str).str.upper().str.contains("CHARGEBACK", na=False)
            df_cb = df[mask].copy()
            
            # Extrair PSP Reference
            df_cb["extracted_psp_ref"] = df_cb[denom_col].apply(extract_psp_ref)
            df_cb["sap_source_file"] = os.path.basename(f)
            
            print(f"    - Total de Chargebacks extraídos: {len(df_cb)} linhas")
            dfs.append(df_cb)
        except Exception as e:
            print(f"    ❌ Erro ao ler {os.path.basename(f)}: {e}")
            
    if not dfs:
        raise ValueError("Nenhum dado de Chargeback pôde ser carregado do SAP.")
        
    df_sap_all = pd.concat(dfs, ignore_index=True)
    print(f"  Total consolidado SAP: {len(df_sap_all)} chargebacks.")
    return df_sap_all


def load_adyen_data(adyen_folder):
    """
    Carrega todos os relatórios CSV da Adyen e indexa-os para busca rápida.
    """
    print("\n[2/4] Carregando dados da Adyen...")
    files = sorted(glob.glob(os.path.join(adyen_folder, "*.csv")))
    
    if not files:
        raise FileNotFoundError(f"Nenhum arquivo CSV encontrado em: {adyen_folder}")
        
    dfs = []
    for f in files:
        # Detectar delimitador: lê a primeira linha e checa se tem ';' ou ','
        try:
            with open(f, "r", encoding="utf-8") as fh:
                first_line = fh.readline()
                sep = ";" if ";" in first_line else ","
        except Exception:
            sep = ","  # Fallback
            
        print(f"  → Lendo Adyen: {os.path.basename(f)} (delimitador: '{sep}')")
        try:
            df = pd.read_csv(f, sep=sep, on_bad_lines="skip", engine="python")
            # Padroniza colunas (remove espaços extras)
            df.columns = [c.strip() for c in df.columns]
            
            required = ["Psp Reference", "Record Type", "Dispute Amount"]
            missing = [r for r in required if r not in df.columns]
            if missing:
                print(f"    ⚠️ Colunas obrigatórias ausentes em {os.path.basename(f)}: {missing}. Pulando.")
                continue
                
            dfs.append(df)
        except Exception as e:
            print(f"    ❌ Erro ao ler {os.path.basename(f)}: {e}")
            
    if not dfs:
        raise ValueError("Nenhum dado pôde ser carregado da Adyen.")
        
    df_adyen_all = pd.concat(dfs, ignore_index=True)
    print(f"  Total carregado Adyen: {len(df_adyen_all)} registros.")
    
    # Extrair metadados JSON para colunas próprias para facilitar o join com ClearSale
    if "Metadata" in df_adyen_all.columns:
        print("  - Extraindo IDs de integração do Metadata Adyen...")
        meta_extracted = df_adyen_all["Metadata"].apply(extract_metadata)
        df_adyen_all["meta_order_id"] = [x[0] for x in meta_extracted]
        df_adyen_all["meta_payment_id"] = [x[1] for x in meta_extracted]
        df_adyen_all["meta_transaction_id"] = [x[2] for x in meta_extracted]
    else:
        df_adyen_all["meta_order_id"] = None
        df_adyen_all["meta_payment_id"] = None
        df_adyen_all["meta_transaction_id"] = None
        
    # Garantir chaves em String, maiúsculas e sem espaços extras para busca exata
    df_adyen_all["psp_key"] = df_adyen_all["Psp Reference"].astype(str).str.strip().str.upper()
    
    # Remover duplicatas no Psp Reference para o cruzamento, mantendo o mais detalhado
    df_adyen_unique = df_adyen_all.drop_duplicates(subset=["psp_key"], keep="last")
    print(f"  Total de Psp References únicos na Adyen: {len(df_adyen_unique)}")
    
    return df_adyen_unique


def load_clearsale_data(cs_folder):
    """
    Carrega todos os relatórios da ClearSale e indexa por prefixo do campo PEDIDO.
    """
    print("\n[3/4] Carregando dados da ClearSale...")
    files = glob.glob(os.path.join(cs_folder, "*.xls")) + glob.glob(os.path.join(cs_folder, "*.csv"))
    files = [f for f in files if not os.path.basename(f).startswith("~$")]
    
    if not files:
        print("  ⚠️ Nenhum arquivo ClearSale encontrado. O cruzamento com a ClearSale será ignorado.")
        return pd.DataFrame()
        
    dfs = []
    for f in files:
        print(f"  → Lendo ClearSale: {os.path.basename(f)}")
        df = None
        # Tenta ler como HTML disfarçado de XLS
        if f.lower().endswith(".xls"):
            try:
                tables = pd.read_html(f)
                if tables:
                    df = tables[0]
                    print("    - Detectado formato HTML Spreadsheet")
            except Exception:
                pass
                
        # Se não deu certo ou é CSV, tenta CSV com ponto e vírgula e codificações diferentes
        if df is None or "PEDIDO" not in df.columns:
            for enc in ["utf-8-sig", "latin1", "iso-8859-1"]:
                try:
                    df = pd.read_csv(f, sep=";", encoding=enc, on_bad_lines="skip")
                    if "PEDIDO" in df.columns:
                        print(f"    - Lido com sucesso usando codificação: {enc}")
                        break
                except Exception:
                    continue
                    
        if df is not None and "PEDIDO" in df.columns:
            # Padroniza colunas
            df.columns = [c.strip() for c in df.columns]
            
            # Criar chave de cruzamento baseada no prefixo de PEDIDO antes do hífen
            # Exemplo: '2AEC561E9FA441E8A30DB2D3EF9A8266-3337338' -> '2AEC561E9FA441E8A30DB2D3EF9A8266'
            df["cs_join_key"] = df["PEDIDO"].astype(str).str.split("-").str[0].str.strip().str.upper()
            df["cs_source_file"] = os.path.basename(f)
            
            dfs.append(df)
            print(f"    - Carregados {len(df)} registros")
        else:
            print(f"    ⚠️ Não foi possível determinar a estrutura do arquivo {os.path.basename(f)} ou coluna 'PEDIDO' ausente.")
            
    if not dfs:
        print("  ⚠️ Nenhum dado útil pôde ser carregado da ClearSale.")
        return pd.DataFrame()
        
    df_cs_all = pd.concat(dfs, ignore_index=True)
    
    # Remove duplicatas de chaves ClearSale mantendo o último registro
    df_cs_unique = df_cs_all.drop_duplicates(subset=["cs_join_key"], keep="last")
    print(f"  Total de chaves ClearSale únicas: {len(df_cs_unique)}")
    return df_cs_unique


def classify_delivery(d_val):
    if not d_val or pd.isna(d_val):
        return None
    d_lower = str(d_val).lower()
    if any(k in d_lower for k in ["retirada", "retire", "retira", "pickup", "retirar"]):
        return "Clique e Retire"
    elif any(k in d_lower for k in ["entrega", "delivery", "normal", "expressa", "sedex", "pac", "correios", "economic"]):
        return "Entrega em Casa"
    else:
        return str(d_val)


def enrich_vtex_data(df, cache_path="data/vtex_cache.json"):
    """
    Enriquece o DataFrame consolidado com informações adicionais buscadas na VTEX.
    Tenta primeiro buscar no cache local para evitar chamadas lentas de API.
    Se não encontrar no cache, faz a busca na API em paralelo e atualiza o cache.
    """
    print("\n[VTEX] Iniciando enriquecimento com a API/Cache da VTEX...")
    
    # 1. Carregar cache existente
    vtex_cache = {}
    if os.path.exists(cache_path):
        try:
            with open(cache_path, "r", encoding="utf-8") as f:
                vtex_cache = json.load(f)
            print(f"  → Cache carregado com {len(vtex_cache)} pedidos.")
        except Exception as e:
            print(f"  ⚠️ Erro ao carregar cache VTEX: {e}")
            
    # 2. Identificar pedidos a buscar (exclui nulos e strings vazias)
    if "adyen_vtex_order_id" not in df.columns:
        print("  ⚠️ Coluna 'adyen_vtex_order_id' não encontrada. Pulando enriquecimento VTEX.")
        return df
        
    mask_valid = df["adyen_vtex_order_id"].notna() & df["adyen_vtex_order_id"].astype(str).str.strip().ne("")
    order_ids = [str(oid).strip() for oid in df.loc[mask_valid, "adyen_vtex_order_id"].unique()]
    print(f"  → Total de {len(order_ids)} IDs de pedidos únicos para enriquecer.")
    
    # Descobrir quais não estão no cache (considerando possíveis sufixos -01, etc.)
    missing_oids = []
    for oid in order_ids:
        formatos = [oid] if "-" in oid else [f"{oid}-01", oid, f"{oid}-02"]
        found = False
        for fmt in formatos:
            if fmt in vtex_cache:
                found = True
                break
        if not found:
            missing_oids.append(oid)
            
    print(f"  → {len(order_ids) - len(missing_oids)} pedidos já estão no cache, {len(missing_oids)} pendentes para consulta na API.")
    
    # 3. Consultar API VTEX para os faltantes se houver
    if missing_oids:
        print("  → Iniciando consultas na API VTEX...")
        from vtex_extractor import VTEXExtractor
        from config.settings import VTEXConfig
        config = VTEXConfig()
        extractor = VTEXExtractor(config)
        
        # Testar conexão
        if not extractor.ping():
            print("  ❌ Não foi possível conectar à API VTEX (Ping falhou). Usando apenas dados do cache.")
        else:
            def fetch_one(oid):
                formatos = [oid] if "-" in oid else [f"{oid}-01", oid, f"{oid}-02"]
                for fmt in formatos:
                    try:
                        detail = extractor.get_order_detail(fmt)
                        flat = extractor._flatten_detail(detail)
                        return oid, flat
                    except KeyError:
                        continue
                    except Exception:
                        continue
                return oid, None

            completed = 0
            with ThreadPoolExecutor(max_workers=15) as executor_pool:
                futures = {executor_pool.submit(fetch_one, oid): oid for oid in missing_oids}
                for future in as_completed(futures):
                    oid, result = future.result()
                    if result:
                        vtex_cache[oid] = result
                    completed += 1
                    if completed % 50 == 0 or completed == len(missing_oids):
                        print(f"    - {completed}/{len(missing_oids)} pedidos consultados...", flush=True)
                        # Salvar cache parcial
                        try:
                            with open(cache_path, "w", encoding="utf-8") as f:
                                json.dump(vtex_cache, f, ensure_ascii=False, indent=2)
                        except Exception:
                            pass
                            
            # Salvar cache final
            try:
                with open(cache_path, "w", encoding="utf-8") as f:
                    json.dump(vtex_cache, f, ensure_ascii=False, indent=2)
                print(f"  → Cache final salvo com {len(vtex_cache)} pedidos.")
            except Exception as e:
                print(f"  ⚠️ Erro ao salvar cache VTEX: {e}")
                
    # 4. Mapear cache → DataFrame
    vtex_loja_list = []
    vtex_cidade_list = []
    vtex_uf_list = []
    vtex_forma_entrega_list = []
    vtex_tipo_produto_list = []
    vtex_categoria_list = []
    
    for idx, row in df.iterrows():
        oid = row.get("adyen_vtex_order_id")
        if pd.isna(oid) or not str(oid).strip():
            vtex_loja_list.append(None)
            vtex_cidade_list.append(None)
            vtex_uf_list.append(None)
            vtex_forma_entrega_list.append(None)
            vtex_tipo_produto_list.append(None)
            vtex_categoria_list.append(None)
            continue
            
        oid_str = str(oid).strip()
        formatos = [oid_str] if "-" in oid_str else [f"{oid_str}-01", oid_str, f"{oid_str}-02"]
        
        vtex_data = None
        for fmt in formatos:
            if fmt in vtex_cache:
                vtex_data = vtex_cache[fmt]
                break
                
        if vtex_data:
            vtex_loja_list.append(vtex_data.get("vtex_store"))
            vtex_cidade_list.append(vtex_data.get("vtex_cidade"))
            vtex_uf_list.append(vtex_data.get("vtex_uf"))
            
            # Classificar forma de entrega
            deliv_raw = vtex_data.get("vtex_delivery_type")
            vtex_forma_entrega_list.append(classify_delivery(deliv_raw))
            
            # Produtos e categorias
            vtex_tipo_produto_list.append(vtex_data.get("vtex_products"))
            vtex_categoria_list.append(vtex_data.get("vtex_categories"))
        else:
            vtex_loja_list.append(None)
            vtex_cidade_list.append(None)
            vtex_uf_list.append(None)
            vtex_forma_entrega_list.append(None)
            vtex_tipo_produto_list.append(None)
            vtex_categoria_list.append(None)
            
    df["vtex_loja"] = vtex_loja_list
    df["vtex_cidade"] = vtex_cidade_list
    df["vtex_uf"] = vtex_uf_list
    df["vtex_forma_entrega"] = vtex_forma_entrega_list
    df["vtex_tipo_produto"] = vtex_tipo_produto_list
    df["vtex_categoria"] = vtex_categoria_list
    
    match_count = df["vtex_loja"].notna().sum()
    print(f"  → Enriquecimento concluído! {match_count}/{len(df)} linhas enriquecidas com dados da VTEX.")
    return df


def consolidate_pipeline():
    """
    Executa todo o pipeline de carregamento, cruzamento e exportação.
    """
    start_time = datetime.now()
    print("=" * 70)
    print("  INICIANDO PIPELINE DE CONSOLIDAÇÃO DE CHARGEBACKS")
    print("=" * 70)
    
    SAP_FOLDER = "data/SAP"
    ADYEN_FOLDER = "data/adyen"
    CS_FOLDER = "data/ClearSale"
    OUTPUT_FILE = "relatorio_chargeback_consolidado.xlsx"
    
    # 1. Carrega Fontes
    df_sap = load_sap_data(SAP_FOLDER)
    df_adyen = load_adyen_data(ADYEN_FOLDER)
    df_cs = load_clearsale_data(CS_FOLDER)
    
    # Criar estruturas de dicionários rápidos para busca O(1)
    print("\n[4/4] Realizando o cruzamento inteligente de dados...")
    
    # Converter Adyen para dicionário baseado em 'psp_key'
    adyen_dict = df_adyen.set_index("psp_key").to_dict(orient="index")
    
    # Converter ClearSale para dicionário baseado em 'cs_join_key'
    cs_dict = {}
    if not df_cs.empty:
        cs_dict = df_cs.set_index("cs_join_key").to_dict(orient="index")
        
    # Inicializar listas para construir os DataFrames finais
    consolidated_rows = []
    unmatched_sap_rows = []
    
    match_adyen_count = 0
    match_cs_count = 0
    
    # Itera sobre cada chargeback do SAP
    for idx, row in df_sap.iterrows():
        # Copia todos os dados originais do SAP
        sap_data = row.to_dict()
        psp_ref = sap_data.get("extracted_psp_ref")
        
        # Cria um registro básico
        merged_row = sap_data.copy()
        
        # Colunas Adyen complementares padrão (None se não achar)
        adyen_cols = {
            "adyen_psp_reference": None,
            "adyen_record_type": None,
            "adyen_record_date": None,
            "adyen_payment_date": None,
            "adyen_payment_amount": None,
            "adyen_dispute_amount": None,
            "adyen_dispute_reason": None,
            "adyen_payment_method": None,
            "adyen_shopper_name": None,
            "adyen_shopper_email": None,
            "adyen_shopper_phone": None,
            "adyen_shopper_ip": None,
            "adyen_shopper_country": None,
            "adyen_risk_scoring": None,
            "adyen_vtex_order_id": None,
            "adyen_vtex_payment_id": None,
            "adyen_vtex_transaction_id": None,
            "adyen_source_file": None
        }
        
        # Colunas ClearSale complementares padrão
        cs_cols = {
            "cs_pedido": None,
            "cs_status_finalizacao": None,
            "cs_data_finalizacao": None,
            "cs_valor": None,
            "cs_score": None,
            "cs_item_principal": None,
            "cs_status_chargeback": None,
            "cs_email": None,
            "cs_alertas": None,
            "cs_source_file": None
        }
        
        found_adyen = False
        found_cs = False
        
        if psp_ref and psp_ref in adyen_dict:
            found_adyen = True
            match_adyen_count += 1
            a_record = adyen_dict[psp_ref]
            
            # Preencher campos Adyen
            adyen_cols["adyen_psp_reference"] = a_record.get("Psp Reference")
            adyen_cols["adyen_record_type"] = a_record.get("Record Type")
            adyen_cols["adyen_record_date"] = a_record.get("Record Date")
            adyen_cols["adyen_payment_date"] = a_record.get("Payment Date")
            adyen_cols["adyen_payment_amount"] = a_record.get("Payment Amount")
            adyen_cols["adyen_dispute_amount"] = a_record.get("Dispute Amount")
            adyen_cols["adyen_dispute_reason"] = a_record.get("Dispute Reason")
            adyen_cols["adyen_payment_method"] = a_record.get("Payment Method")
            adyen_cols["adyen_shopper_name"] = a_record.get("Shopper Name")
            adyen_cols["adyen_shopper_email"] = a_record.get("Shopper Email")
            adyen_cols["adyen_shopper_phone"] = a_record.get("Shopper Phone Number")
            adyen_cols["adyen_shopper_ip"] = a_record.get("Shopper IP")
            adyen_cols["adyen_shopper_country"] = a_record.get("Shopper Country")
            adyen_cols["adyen_risk_scoring"] = a_record.get("Risk Scoring")
            adyen_cols["adyen_vtex_order_id"] = a_record.get("meta_order_id")
            adyen_cols["adyen_vtex_payment_id"] = a_record.get("meta_payment_id")
            adyen_cols["adyen_vtex_transaction_id"] = a_record.get("meta_transaction_id")
            
            # Buscar na ClearSale usando as chaves de metadados extraídas da Adyen
            # Tenta combinar com transactionId, depois paymentId, depois orderId
            cs_match_key = None
            for key in ["meta_transaction_id", "meta_payment_id", "meta_order_id"]:
                val = a_record.get(key)
                if val:
                    val_str = str(val).strip().upper()
                    if val_str in cs_dict:
                        cs_match_key = val_str
                        break
            
            if cs_match_key:
                found_cs = True
                match_cs_count += 1
                cs_record = cs_dict[cs_match_key]
                
                # Preencher campos ClearSale
                cs_cols["cs_pedido"] = cs_record.get("PEDIDO")
                cs_cols["cs_status_finalizacao"] = cs_record.get("STATUS FINALIZACAO")
                cs_cols["cs_data_finalizacao"] = cs_record.get("DATA FINALIZACAO")
                cs_cols["cs_valor"] = cs_record.get("VALOR")
                cs_cols["cs_score"] = cs_record.get("SCORE")
                cs_cols["cs_item_principal"] = cs_record.get("ITEM PRINCIPAL")
                cs_cols["cs_status_chargeback"] = cs_record.get("STATUS DO CHARGEBACK")
                cs_cols["cs_email"] = cs_record.get("E-MAIL")
                cs_cols["cs_alertas"] = cs_record.get("ALERTAS")
                cs_cols["cs_source_file"] = cs_record.get("cs_source_file")
                
        # Junta todas as colunas
        merged_row.update(adyen_cols)
        merged_row.update(cs_cols)
        
        # Adiciona flag indicador de fonte
        if found_adyen and found_cs:
            merged_row["fonte_consolidada"] = "SAP + Adyen + ClearSale"
        elif found_adyen:
            merged_row["fonte_consolidada"] = "SAP + Adyen"
        else:
            merged_row["fonte_consolidada"] = "Apenas SAP"
            
        consolidated_rows.append(merged_row)
        
        # Se não encontrou na Adyen, vai também para a aba de Auditoria/Sem correspondência
        if not found_adyen:
            unmatched_sap_rows.append(sap_data)
            
    # Cria os DataFrames
    df_consolidado = pd.DataFrame(consolidated_rows)
    df_sem_match = pd.DataFrame(unmatched_sap_rows)
    
    # Enriquecer com dados da VTEX
    df_consolidado = enrich_vtex_data(df_consolidado)
    
    # 5. Organizar e Salvar Excel
    print(f"\n[5/5] Exportando dados para '{OUTPUT_FILE}'...")
    
    # Criar aba de estatísticas rápidas
    total_sap = len(df_sap)
    match_adyen_pct = (match_adyen_count / total_sap * 100) if total_sap > 0 else 0
    match_cs_pct = (match_cs_count / max(match_adyen_count, 1) * 100) if match_adyen_count > 0 else 0
    match_vtex_count = df_consolidado["vtex_loja"].notna().sum()
    match_vtex_pct = (match_vtex_count / max(match_adyen_count, 1) * 100) if match_adyen_count > 0 else 0
    
    df_metrics = pd.DataFrame([
        {"Métrica": "Total de Chargebacks no SAP", "Valor": total_sap},
        {"Métrica": "Cruzados com Adyen (Match PSP)", "Valor": match_adyen_count},
        {"Métrica": "Taxa de Cruzamento SAP → Adyen (%)", "Valor": round(match_adyen_pct, 2)},
        {"Métrica": "Cruzados com ClearSale (Match IDs)", "Valor": match_cs_count},
        {"Métrica": "Taxa de Enriquecimento Adyen → ClearSale (%)", "Valor": round(match_cs_pct, 2)},
        {"Métrica": "Enriquecidos com VTEX (Loja/Local/Entrega)", "Valor": match_vtex_count},
        {"Métrica": "Taxa de Enriquecimento Adyen → VTEX (%)", "Valor": round(match_vtex_pct, 2)},
        {"Métrica": "Não Encontrados na Adyen (Aba Sem Correspondência)", "Valor": len(df_sem_match)},
        {"Métrica": "Data de Geração", "Valor": datetime.now().strftime("%d/%m/%Y %H:%M:%S")}
    ])
    
    try:
        with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
            df_consolidado.to_excel(writer, sheet_name="Consolidado Completo", index=False)
            if not df_sem_match.empty:
                df_sem_match.to_excel(writer, sheet_name="Sem Correspondência (Adyen)", index=False)
            df_metrics.to_excel(writer, sheet_name="Métricas Gerais", index=False)
        output_path_final = os.path.abspath(OUTPUT_FILE)
    except PermissionError:
        fallback_file = "relatorio_chargeback_consolidado_novo.xlsx"
        print(f"\n⚠️ Permissão negada ao salvar em '{OUTPUT_FILE}' (provavelmente o arquivo está aberto no Excel).")
        print(f"  → Tentando salvar em uma cópia alternativa: '{fallback_file}'...")
        with pd.ExcelWriter(fallback_file, engine="openpyxl") as writer:
            df_consolidado.to_excel(writer, sheet_name="Consolidado Completo", index=False)
            if not df_sem_match.empty:
                df_sem_match.to_excel(writer, sheet_name="Sem Correspondência (Adyen)", index=False)
            df_metrics.to_excel(writer, sheet_name="Métricas Gerais", index=False)
        output_path_final = os.path.abspath(fallback_file)
        
    duration = datetime.now() - start_time
    print("=" * 70)
    print("  PIPELINE CONCLUÍDO COM SUCESSO!")
    print(f"  → Tempo total de execução: {duration.total_seconds():.2f} segundos")
    print(f"  → Planilha gerada: {output_path_final}")
    print(f"  → Taxa de cruzamento com Adyen: {match_adyen_pct:.2f}%")
    print(f"  → Taxa de cruzamento com ClearSale: {match_cs_pct:.2f}%")
    print(f"  → Taxa de enriquecimento com VTEX: {match_vtex_pct:.2f}%")
    print("=" * 70)


if __name__ == "__main__":
    try:
        consolidate_pipeline()
    except Exception as e:
        print(f"\n❌ ERRO FATAL no pipeline: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
