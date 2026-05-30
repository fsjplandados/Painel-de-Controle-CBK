"""
Extrator VTEX — Orders API
Baseado na resposta real confirmada (Status 200).

Limitação VTEX: máximo page=30 com per_page=100 (3.000 registros por query).
Solução: quebrar em janelas de 2 horas automaticamente.

Credenciais:
  Header X-VTEX-API-AppKey   = vtexappkey-sjdigital-NBIBYX
  Header X-VTEX-API-AppToken = <token>

Campos extras via fetch_details=True:
  - vtex_store:        nome/hostname da loja onde o pedido foi feito
  - vtex_installments: número de parcelas do pagamento
  - vtex_coupon:       código do cupom de desconto (se utilizado)
  - invoice_number:    número da nota fiscal (cupom fiscal)
  - invoice_key:       chave de acesso da nota fiscal
"""
import sys, os
# Garante que config/ e utils/ sejam encontrados a partir do próprio diretório do projeto
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import time
import requests
import urllib3
from datetime import date, datetime, timedelta, timezone
from typing import Iterator
import pandas as pd

from config.settings import VTEXConfig
from utils.logger import get_logger
from utils.retry import with_retry

logger = get_logger("vtex_extractor")

# VTEX limita a page * per_page <= 3000
VTEX_MAX_PAGE    = 30
VTEX_PER_PAGE    = 100
VTEX_MAX_RECORDS = VTEX_MAX_PAGE * VTEX_PER_PAGE  # 3.000


class VTEXExtractor:
    """
    Extrai pedidos da VTEX Orders API com paginação automática.
    Quebra automaticamente em janelas de 2h quando há mais de 3.000 pedidos.

    Uso:
        extractor = VTEXExtractor(config.vtex)
        extractor.session.verify = False  # redes corporativas com Fortinet
        df = extractor.extract(start_date=date(2026, 4, 20), end_date=date(2026, 4, 20))
    """

    BASE_URL = "https://{account}.vtexcommercestable.com.br"

    def __init__(self, config: VTEXConfig):
        self.config = config
        self.base_url = self.BASE_URL.format(account=config.account)
        self.session = requests.Session()
        self.session.headers.update({
            "X-VTEX-API-AppKey":   config.app_key,
            "X-VTEX-API-AppToken": config.app_token,
            "Content-Type":        "application/json",
        })

    # ── HTTP base ────────────────────────────────────────────

    @with_retry(max_attempts=3, backoff_seconds=2.0, exceptions=(requests.RequestException,))
    def _get(self, endpoint: str, params: dict = None) -> dict:
        url = f"{self.base_url}{endpoint}"
        response = self.session.get(url, params=params, timeout=self.config.timeout_seconds)
        if response.status_code == 404:
            # Levanta KeyError para desviar do decorator @with_retry e evitar retries lentos em erros 404 (Not Found)
            raise KeyError(f"Recurso não encontrado (404): {url}")
        response.raise_for_status()
        return response.json()

    # ── Ping ─────────────────────────────────────────────────

    def ping(self) -> bool:
        try:
            resp = self.session.get(f"{self.base_url}/api/oms/pvt/orders", timeout=10)
            ok = resp.status_code == 200
            if ok:
                total = resp.json().get("paging", {}).get("total", "?")
                logger.info(f"VTEX ping OK — {total:,} pedidos na conta")
            return ok
        except Exception as e:
            logger.error(f"VTEX ping erro: {e}")
            return False

    # ── Listagem de uma janela de tempo ──────────────────────

    def _fetch_window(self, start_dt: datetime, end_dt: datetime) -> list:
        """
        Busca todos os pedidos de uma janela datetime com paginação.
        Limitado a 30 páginas × 100 = 3.000 registros por janela.
        Se ultrapassar, loga aviso (janela deve ser menor).
        """
        start_str = start_dt.strftime("%Y-%m-%dT%H:%M:%S.000Z")
        end_str   = end_dt.strftime("%Y-%m-%dT%H:%M:%S.000Z")
        date_filter = f"creationDate:[{start_str} TO {end_str}]"

        # Página 1 — descobrir total
        first = self._get("/api/oms/pvt/orders", {
            "f_creationDate": date_filter,
            "per_page": VTEX_PER_PAGE,
            "page": 1,
            "orderBy": "creationDate,asc",
        })
        paging = first.get("paging", {})
        total  = paging.get("total", 0)
        pages  = min(paging.get("pages", 1), VTEX_MAX_PAGE)

        if total > VTEX_MAX_RECORDS:
            logger.warning(
                f"Janela {start_str[:16]} → {end_str[:16]}: {total} pedidos > {VTEX_MAX_RECORDS} limite. "
                "Alguns pedidos serão perdidos — reduza a janela de tempo."
            )

        rows = list(first.get("list", []))

        for page in range(2, pages + 1):
            result = self._get("/api/oms/pvt/orders", {
                "f_creationDate": date_filter,
                "per_page": VTEX_PER_PAGE,
                "page": page,
                "orderBy": "creationDate,asc",
            })
            rows.extend(result.get("list", []))
            time.sleep(0.25)

        return rows

    # ── Janelas de 2h automáticas ────────────────────────────

    def _iter_windows(self, start_date: date, end_date: date, window_hours: int = 2):
        """
        Divide o período em janelas de `window_hours` horas.
        Default 2h → máximo ~500 pedidos/janela para o volume observado (5.653/dia).
        """
        start_dt = datetime(start_date.year, start_date.month, start_date.day,
                            0, 0, 0, tzinfo=timezone.utc)
        end_dt   = datetime(end_date.year, end_date.month, end_date.day,
                            23, 59, 59, tzinfo=timezone.utc)

        window = timedelta(hours=window_hours)
        current = start_dt
        windows = []
        while current < end_dt:
            w_end = min(current + window - timedelta(seconds=1), end_dt)
            windows.append((current, w_end))
            current += window

        logger.info(
            f"VTEX: {start_date} → {end_date} dividido em "
            f"{len(windows)} janelas de {window_hours}h"
        )
        return windows

    # ── Extração principal ───────────────────────────────────

    def extract(
        self,
        start_date: date,
        end_date:   date,
        window_hours: int = 2,
        fetch_details: bool = False,
    ) -> pd.DataFrame:
        """
        Extrai todos os pedidos do período quebrando em janelas de tempo.

        Args:
            start_date:    data inicial (inclusive)
            end_date:      data final (inclusive)
            window_hours:  tamanho da janela em horas (padrão 2h)
                           diminuir se houver warning de >3.000 pedidos/janela
            fetch_details: busca detalhe de cada pedido (SLA, endereço, NF)
                           — muito mais lento, 1 req/pedido
        """
        logger.info(f"Iniciando extração VTEX: {start_date} → {end_date}")
        windows = self._iter_windows(start_date, end_date, window_hours)
        all_rows = []

        for i, (w_start, w_end) in enumerate(windows, 1):
            label = f"{w_start.strftime('%H:%M')}–{w_end.strftime('%H:%M')}"
            logger.info(f"VTEX: janela {i}/{len(windows)} ({label})...")
            try:
                orders = self._fetch_window(w_start, w_end)
                for order in orders:
                    row = self._flatten(order)
                    if fetch_details:
                        try:
                            detail = self.get_order_detail(order["orderId"])
                            row.update(self._flatten_detail(detail))
                            time.sleep(0.1)
                        except Exception as e:
                            logger.warning(f"Detalhe falhou {order['orderId']}: {e}")
                    all_rows.append(row)
                logger.info(f"  → {len(orders)} pedidos coletados")
            except Exception as e:
                logger.error(f"Janela {label} falhou: {e}")
                continue

        df = pd.DataFrame(all_rows)
        if not df.empty:
            df = self._normalize_types(df)
            # Remover duplicatas (janelas sobrepostas podem gerar)
            df = df.drop_duplicates(subset=["order_id"])

        logger.info(f"VTEX: {len(df)} pedidos extraídos no total")
        return df

    # ── Detalhe por order_id ─────────────────────────────────

    @with_retry(max_attempts=3, backoff_seconds=1.5, exceptions=(requests.RequestException,))
    def get_order_detail(self, order_id: str) -> dict:
        return self._get(f"/api/oms/pvt/orders/{order_id}")

    # ── Flatten campos reais ─────────────────────────────────

    def _flatten(self, o: dict) -> dict:
        return {
            "order_id":               o.get("orderId"),
            "sequence":               o.get("sequence"),
            "client_name":            o.get("clientName"),
            "hostname":               o.get("hostname"),
            "origin":                 o.get("origin"),
            "sales_channel":          o.get("salesChannel"),
            "status":                 o.get("status"),
            "status_description":     o.get("statusDescription"),
            "is_all_delivered":       o.get("isAllDelivered"),
            "is_any_delivered":       o.get("isAnyDelivered"),
            "workflow_error":         o.get("workflowInErrorState"),
            "total_value":            (o.get("totalValue") or 0) / 100,
            "total_items":            o.get("totalItems"),
            "payment_name":           o.get("paymentNames"),
            "creation_date":          o.get("creationDate"),
            "last_change":            o.get("lastChange"),
            "authorized_date":        o.get("authorizedDate"),
            "payment_approved_date":  o.get("paymentApprovedDate"),
            "ready_for_handling_date":o.get("readyForHandlingDate"),
            "shipping_estimated_date":    o.get("ShippingEstimatedDate"),
            "shipping_estimated_date_min":o.get("ShippingEstimatedDateMin"),
            "shipping_estimated_date_max":o.get("ShippingEstimatedDateMax"),
        }

    def _flatten_detail(self, d: dict) -> dict:
        extra = {}

        # ── Endereço ─────────────────────────────────────────
        shipping = d.get("shippingData") or {}
        address  = shipping.get("address") or {}
        extra["uf"]           = address.get("state")
        extra["cidade"]       = address.get("city")
        extra["cep"]          = address.get("postalCode")
        extra["address_type"] = address.get("addressType")

        # ── SLA / Logística ───────────────────────────────────
        log_info = shipping.get("logisticsInfo") or []
        if log_info:
            li = log_info[0]
            extra["sla_type"]          = li.get("selectedSla")
            extra["delivery_deadline"] = li.get("shippingEstimate")

        # ── Nota fiscal ───────────────────────────────────────
        packages = (d.get("packageAttachment") or {}).get("packages") or []
        if packages:
            extra["invoice_key"]    = packages[0].get("invoiceKey")
            extra["invoice_number"] = packages[0].get("invoiceNumber")

        # ── Loja (sellers[].name = nome da franquia/filial) ────
        # sellers[0].name contém o nome real da loja (ex: "BLUMENAU 1 - CNPJ - 1234")
        # Extraímos apenas a parte antes do primeiro " - " (CNPJ).
        store_name = None
        try:
            sellers = d.get("sellers") or []
            if sellers:
                raw_name = sellers[0].get("name", "")
                # Nome vem como "BLUMENAU 1 - 88212113... - 1234"
                # Pega tudo antes do primeiro " - "
                store_name = raw_name.split(" - ")[0].strip() if raw_name else None
        except (IndexError, AttributeError, KeyError):
            pass
        # Fallback: hostname (caso sellers esteja vazio)
        if not store_name:
            store_name = d.get("hostname") or None
        extra["vtex_store"] = store_name

        # ── Parcelamento + valor da parcela ───────────────────
        # paymentData.transactions[].payments[].installments
        # paymentData.transactions[].payments[].value (em centavos)
        installments = None
        installment_value = None
        try:
            payments = (
                (d.get("paymentData") or {})
                .get("transactions", [{}])[0]
                .get("payments", [{}])
            )
            if payments:
                installments = payments[0].get("installments")
                pay_value = payments[0].get("value")
                if pay_value is not None and installments:
                    # Valor em centavos → BRL, dividido pelo nº de parcelas
                    installment_value = round(pay_value / (100 * installments), 2)
        except (IndexError, AttributeError, KeyError):
            pass
        extra["vtex_installments"] = installments
        extra["vtex_installment_value"] = installment_value

        # ── Cupom de desconto + valor do desconto ─────────────
        # marketingData.coupon contém o código do cupom aplicado.
        marketing = d.get("marketingData") or {}
        coupon = marketing.get("coupon") or marketing.get("couponCode") or None
        extra["vtex_coupon"] = coupon

        # Valor total de descontos vem na lista totals, id="Discounts" (em centavos)
        discount_value = None
        try:
            totals = d.get("totals") or []
            for t in totals:
                if t.get("id") == "Discounts":
                    raw_disc = t.get("value", 0)
                    if raw_disc != 0:
                        # VTEX retorna negativo (ex: -500 = -R$5,00), convertemos para positivo em BRL
                        discount_value = round(abs(raw_disc) / 100, 2)
                    break
        except (IndexError, AttributeError, KeyError):
            pass
        extra["vtex_discount_value"] = discount_value

        # ── Tipo de Produto / Itens ───────────────────────────
        items = d.get("items") or []
        product_names = []
        categories_set = set()
        
        for item in items:
            name = item.get("name")
            if name:
                product_names.append(name)
            
            # Categorias podem vir em productCategories (dict: {"id": "Nome da Categoria"})
            prod_cats = item.get("productCategories")
            if isinstance(prod_cats, dict):
                for cat_name in prod_cats.values():
                    if cat_name:
                        categories_set.add(cat_name)
                        
        extra["vtex_products"] = " | ".join(product_names) if product_names else None
        extra["vtex_categories"] = " | ".join(sorted(categories_set)) if categories_set else None

        return extra

    def _normalize_types(self, df: pd.DataFrame) -> pd.DataFrame:
        ts_cols = [
            "creation_date", "last_change", "authorized_date",
            "payment_approved_date", "ready_for_handling_date",
            "shipping_estimated_date", "shipping_estimated_date_min",
            "shipping_estimated_date_max",
        ]
        for col in ts_cols:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], errors="coerce", utc=True)

        if "sequence" in df.columns:
            df["sequence"] = pd.to_numeric(df["sequence"], errors="coerce").astype("Int64")
        if "total_value" in df.columns:
            df["total_value"] = pd.to_numeric(df["total_value"], errors="coerce")
        return df

    # ── Fallback CSV ─────────────────────────────────────────

    @staticmethod
    def from_csv(filepath: str) -> pd.DataFrame:
        logger.info(f"VTEX: lendo CSV {filepath}")
        df = pd.read_csv(filepath, sep=None, engine="python")
        df.columns = [c.lstrip("\ufeff") for c in df.columns]
        rename = {
            "Order": "order_id", "Sequence": "sequence",
            "Status": "status_description", "Creation Date": "creation_date",
            "Last Change Date": "last_change", "Total Value": "total_value",
            "Origin": "origin", "SalesChannel": "sales_channel",
            "Seller Order Id": "seller_order_id", "SLA Type": "sla_type",
            "Estimate Delivery Date": "shipping_estimated_date",
            "Delivery Deadline": "delivery_deadline", "UF": "uf",
            "City": "cidade", "Address Type": "address_type",
            "Tracking Number": "tracking_number", "Delivered": "is_all_delivered",
            "Client Name": "client_name", "Payment System Name": "payment_name",
        }
        df = df.rename(columns={k: v for k, v in rename.items() if k in df.columns})
        for col in ["creation_date", "last_change", "shipping_estimated_date"]:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], errors="coerce", utc=True)
        if "sequence" in df.columns:
            df["sequence"] = pd.to_numeric(df["sequence"], errors="coerce").astype("Int64")
        agg_cols = {
            "sequence", "status_description", "creation_date", "last_change",
            "total_value", "sla_type", "shipping_estimated_date", "delivery_deadline",
            "uf", "cidade", "address_type", "tracking_number", "is_all_delivered",
            "client_name", "payment_name", "origin", "sales_channel",
        }
        agg = {c: "first" for c in agg_cols if c in df.columns}
        if "order_id" in df.columns:
            df = df.groupby("order_id", as_index=False).agg(agg)
        logger.info(f"VTEX CSV: {len(df)} pedidos únicos")
        return df
