"""Kingdee-ierp MCP Server."""

import datetime
import json
import logging
import os
import sys
import time
import uuid
from collections import defaultdict
from functools import lru_cache

import requests
from mcp.server.fastmcp import FastMCP

from ._export_report import export_excel


logging.basicConfig(level=logging.INFO, stream=sys.stderr, format="%(asctime)s [%(levelname)s] %(message)s")
logger = logging.getLogger("kingdee-mcp-ierp")

mcp = FastMCP("kingdee-ierp", dependencies=["requests", "openpyxl"])

SO_BILL_STATUS = {"A": "暂存", "B": "已提交", "C": "已审核", "D": "重新审核"}
SO_ORDER_STATUS = {"A": "未交货", "B": "部分交货", "C": "已交货", "K": "已结案"}

MO_BILL_STATUS = {
    "A": "暂存",
    "B": "已提交",
    "C": "已审核",
    "D": "重新审核",
    "E": "已下达",
    "F": "已开工",
    "G": "已完工",
    "H": "已关闭",
}
MO_BIZ_STATUS = {
    "A": "未开工",
    "B": "开工",
    "C": "完工",
    "D": "部分完工",
    "10": "未开工",
    "20": "开工",
    "30": "完工",
    "40": "部分完工",
}
MO_DONE_STATUS = {"G", "H"}
MO_ACTIVE_STATUS = {"E", "F"}

RISK_LABEL = {
    "completed": "已完工",
    "on_track": "正常",
    "at_risk": "风险",
    "overdue": "逾期",
    "unknown": "未知",
}

EXPORT_DIR = os.path.dirname(os.path.abspath(__file__))
TODAY = datetime.date.today()

NO_PROXY = {"http": None, "https": None}
RETRY_CODES = {500, 502, 503, 504}
MAX_RETRY = 3


def _get_required_env(name):
    value = (os.getenv(name) or "").strip()
    if not value:
        raise RuntimeError(f"缺少环境变量: {name}")
    return value


@lru_cache(maxsize=1)
def get_kingdee_config():
    base_url = _get_required_env("KINGDEE_BASE_URL").rstrip("/")
    return {
        "token_url": f"{base_url}/oauth2/getToken",
        "sale_order_url": f"{base_url}/v2/sm/sm_salorder/query",
        "mft_order_url": f"{base_url}/v2/pom/pom_mftorder/batchQueryNew",
        "mft_stock_url": f"{base_url}/v2/pom/pom_mftstock/query",
        "auth_payload": {
            "client_id": _get_required_env("KINGDEE_CLIENT_ID"),
            "client_secret": _get_required_env("KINGDEE_CLIENT_SECRET"),
            "username": _get_required_env("KINGDEE_USERNAME"),
            "accountId": _get_required_env("KINGDEE_ACCOUNT_ID"),
            "language": os.getenv("KINGDEE_LANGUAGE", "zh_CN"),
        },
    }


def safe_float(value, default=0.0):
    try:
        return float(value) if value is not None else default
    except (TypeError, ValueError):
        return default


def parse_date(value):
    if not value:
        return None
    try:
        return datetime.date.fromisoformat(str(value)[:10])
    except ValueError:
        return None


def get_mo_completion(billstatus, rptqty, mo_qty):
    is_done = billstatus in MO_DONE_STATUS
    rate = (safe_float(rptqty) / safe_float(mo_qty) * 100) if safe_float(mo_qty) > 0 else 0.0
    if billstatus in MO_DONE_STATUS:
        label = f"{'已完工' if billstatus == 'G' else '已关闭'}({billstatus})"
    elif billstatus in MO_ACTIVE_STATUS:
        label = f"{'已下达' if billstatus == 'E' else '已开工'}({billstatus}) {rate:.0f}%"
    else:
        label = f"{MO_BILL_STATUS.get(billstatus, billstatus)}({billstatus})"
    return label, rate, is_done


def get_delivery_risk(mo_planend_str, so_delivery_str, is_done, completime_str=None):
    so_delivery = parse_date(so_delivery_str)
    if is_done:
        actual_end = parse_date(completime_str) or TODAY
        if so_delivery and actual_end > so_delivery:
            return "overdue"
        return "completed"
    if not so_delivery:
        return "unknown"
    if TODAY > so_delivery:
        return "overdue"
    mo_planend = parse_date(mo_planend_str)
    if mo_planend and mo_planend > so_delivery:
        return "at_risk"
    return "on_track"


def _request(method, url, **kwargs):
    for attempt in range(1, MAX_RETRY + 1):
        try:
            response = requests.request(method, url, proxies=NO_PROXY, timeout=30, **kwargs)
            if response.status_code in RETRY_CODES and attempt < MAX_RETRY:
                time.sleep(2**attempt)
                continue
            response.raise_for_status()
            return response
        except (requests.ConnectionError, requests.Timeout):
            if attempt < MAX_RETRY:
                time.sleep(2**attempt)
            else:
                raise


def get_token():
    config = get_kingdee_config()
    payload = {
        **config["auth_payload"],
        "nonce": str(uuid.uuid4()),
        "timestamp": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }
    response = _request(
        "POST",
        config["token_url"],
        data=json.dumps(payload),
        headers={"Content-Type": "application/json"},
    )
    result = response.json()
    if not result.get("status"):
        raise Exception(f"Token 失败: {result.get('message')}")
    return result["data"]["access_token"]


def _post(url, token, body):
    response = _request(
        "POST",
        f"{url}?access_token={token}",
        data=json.dumps(body),
        headers={"Content-Type": "application/json", "access_token": token},
    )
    result = response.json()
    if not result.get("status"):
        raise Exception(f"API 错误 [{url}]: {result.get('errorCode')} {result.get('message')}")
    return result.get("data", {}) or {}


def _paginate(url, token, data_body, label, page_size=100):
    rows = []
    page_no = 1
    while True:
        logger.info("[%s] 第 %s 页...", label, page_no)
        data = _post(url, token, {**data_body, "pageNo": page_no, "pageSize": page_size})
        batch = data.get("rows", []) or []
        total = data.get("totalCount", 0)
        if not batch:
            break
        rows.extend(batch)
        if total > 0 and len(rows) >= total:
            break
        if total == -1 and len(batch) < page_size:
            break
        if total not in (-1, 0) and len(rows) >= total:
            break
        page_no += 1
    return rows


def fetch_sale_orders(token, start_date, end_date):
    return _paginate(
        get_kingdee_config()["sale_order_url"],
        token,
        {"data": {"start_bizdate": start_date, "end_bizdate": end_date}},
        "SO",
    )


def fetch_mft_orders(token, start_date, end_date):
    return _paginate(
        get_kingdee_config()["mft_order_url"],
        token,
        {"data": {"start_billdate": start_date, "end_billdate": end_date}},
        "MO",
    )


def fetch_mft_stock(token, start_date, end_date):
    return _paginate(
        get_kingdee_config()["mft_stock_url"],
        token,
        {"data": {"start_createtime": f"{start_date} 00:00:00", "end_createtime": f"{end_date} 23:59:59"}},
        "STOCK",
    )


def build_so_idx(so_rows):
    return {row.get("billno", ""): row for row in so_rows}


def build_so_entry_idx(so_rows):
    index = {}
    for row in so_rows:
        for entry in row.get("billentry", []) or []:
            index[(row.get("billno", ""), str(entry.get("seq", "")))] = entry
    return index


def build_picking_idx(stock_rows):
    index = defaultdict(
        lambda: {
            "demand": 0.0,
            "issued": 0.0,
            "unissued": 0.0,
            "feeding": 0.0,
            "rejected": 0.0,
            "stock_billnos": [],
        }
    )
    for stock in stock_rows:
        mo_no = stock.get("orderno", "") or ""
        if not mo_no:
            continue
        item = index[mo_no]
        bill_no = stock.get("billno", "")
        if bill_no and bill_no not in item["stock_billnos"]:
            item["stock_billnos"].append(bill_no)
        for entry in stock.get("stockentry", []) or []:
            item["demand"] += safe_float(entry.get("demandqty"))
            item["issued"] += safe_float(entry.get("actissueqty"))
            item["unissued"] += safe_float(entry.get("unissueqty"))
            item["feeding"] += safe_float(entry.get("feedingqty"))
            item["rejected"] += safe_float(entry.get("rejectedqty"))
    for item in index.values():
        item["rate"] = (item["issued"] / item["demand"] * 100) if item["demand"] > 0 else 0.0
    return index


def join_all(so_rows, mo_rows, picking_idx):
    so_idx = build_so_idx(so_rows)
    so_entry_idx = build_so_entry_idx(so_rows)
    relations = []

    for mo in mo_rows:
        mo_billno = mo.get("billno", "")
        mo_status = mo.get("billstatus", "")
        mo_billdate = (mo.get("billdate", "") or "")[:10]
        mo_org = mo.get("org_name", "")
        picking = picking_idx.get(mo_billno, {})

        for entry in mo.get("treeentryentity", []) or []:
            mo_qty = safe_float(entry.get("qty"))
            rptqty = safe_float(entry.get("rptqty"))
            qualified = safe_float(entry.get("qualifiedqty"))
            stockqty = safe_float(entry.get("stockqty"))
            completime = entry.get("completime", "") or ""
            planend = (entry.get("planendtime", "") or "")[:16]
            planbegin = (entry.get("planbegintime", "") or "")[:16]
            biz_raw = str(entry.get("bizstatus", "") or "")
            mo_bizstatus = MO_BIZ_STATUS.get(biz_raw, biz_raw)
            mo_completion, mo_rate, is_done = get_mo_completion(mo_status, rptqty, mo_qty)

            root_billno = entry.get("rootdemandbillno", "") or ""
            root_entity = entry.get("rootdemandentity_number", "") or ""
            root_seq = str(entry.get("rootdemandentryseq", "") or "")
            so_key = root_billno if root_entity == "sm_salorder" else ""
            seq_key = root_seq if root_entity == "sm_salorder" else ""

            so_header = so_idx.get(so_key)
            so_entry = so_entry_idx.get((so_key, seq_key))

            if so_header:
                so_status = so_header.get("billstatus", "")
                so_order_status = so_header.get("orderstatus", "")
                so_delivery = (so_entry.get("deliverydate", "") or "")[:10] if so_entry else ""
                so_mat_no = (so_entry.get("material_masterid_number", "") or "") if so_entry else ""
                so_qty_value = safe_float(so_entry.get("qty")) if so_entry else 0.0
                so_amount = safe_float(so_entry.get("amountandtax")) if so_entry else 0.0
                risk = get_delivery_risk(planend, so_delivery, is_done, completime)
                so_part = {
                    "so_billno": so_key,
                    "so_entry_seq": seq_key,
                    "so_bizdate": (so_header.get("bizdate", "") or "")[:10],
                    "so_billstatus": f"{SO_BILL_STATUS.get(so_status, so_status)}({so_status})",
                    "so_orderstatus": f"{SO_ORDER_STATUS.get(so_order_status, so_order_status)}({so_order_status})",
                    "so_customer_name": so_header.get("customer_name", "") or "",
                    "so_material_no": so_mat_no,
                    "so_qty": so_qty_value,
                    "so_amountandtax": so_amount,
                    "so_deliverydate": so_delivery,
                    "linked": True,
                }
            else:
                risk = get_delivery_risk(planend, "", is_done, completime)
                so_part = {
                    "so_billno": so_key or root_billno,
                    "so_entry_seq": seq_key,
                    "so_bizdate": "",
                    "so_billstatus": "(未找到SO)" if so_key else f"(非SO:{root_entity or '无'})",
                    "so_orderstatus": "",
                    "so_customer_name": "",
                    "so_material_no": "",
                    "so_qty": 0.0,
                    "so_amountandtax": 0.0,
                    "so_deliverydate": "",
                    "linked": False,
                }

            relations.append(
                {
                    "mo_billno": mo_billno,
                    "mo_raw_status": mo_status,
                    "mo_billstatus": f"{MO_BILL_STATUS.get(mo_status, mo_status)}({mo_status})",
                    "mo_bizstatus": mo_bizstatus,
                    "mo_billdate": mo_billdate,
                    "mo_org_name": mo_org,
                    "mo_entry_seq": entry.get("seq", ""),
                    "mo_material_no": entry.get("material_masterid_number", "") or "",
                    "mo_material_name": entry.get("material_masterid_name", "") or "",
                    "mo_qty": mo_qty,
                    "mo_planbegin": planbegin,
                    "mo_planend": planend,
                    "mo_completion": mo_completion,
                    "mo_completion_rate": round(mo_rate, 1),
                    "mo_is_done": is_done,
                    "mo_rptqty": rptqty,
                    "mo_qualifiedqty": qualified,
                    "mo_stockqty": stockqty,
                    "mo_completime": completime[:16] if completime else "",
                    "delivery_risk": risk,
                    "delivery_risk_label": RISK_LABEL[risk],
                    "pick_demand": picking.get("demand", 0.0),
                    "pick_issued": picking.get("issued", 0.0),
                    "pick_unissued": picking.get("unissued", 0.0),
                    "pick_feeding": picking.get("feeding", 0.0),
                    "pick_rejected": picking.get("rejected", 0.0),
                    "pick_rate": picking.get("rate", 0.0),
                    "pick_billnos": ", ".join(picking.get("stock_billnos", [])),
                    **so_part,
                }
            )

    return relations


def build_bottleneck(stock_rows, relations):
    mo_to_so = {}
    mo_to_customer = {}
    for row in relations:
        mo_no = row["mo_billno"]
        so_no = row.get("so_billno", "")
        customer = row.get("so_customer_name", "")
        if so_no:
            mo_to_so.setdefault(mo_no, set()).add(so_no)
        if customer:
            mo_to_customer.setdefault(mo_no, set()).add(customer)

    result = {}
    for stock in stock_rows or []:
        mo_no = stock.get("orderno", "") or ""
        for entry in stock.get("stockentry", []) or []:
            mat_no = entry.get("materialid_number", "") or ""
            if not mat_no:
                continue
            item = result.setdefault(
                mat_no,
                {
                    "mat_no": mat_no,
                    "mat_name": entry.get("materialid_name", "") or "",
                    "demand": 0.0,
                    "issued": 0.0,
                    "unissued": 0.0,
                    "mos": set(),
                    "sos": set(),
                    "customers": set(),
                },
            )
            item["demand"] += safe_float(entry.get("demandqty"))
            item["issued"] += safe_float(entry.get("actissueqty"))
            item["unissued"] += safe_float(entry.get("unissueqty"))
            if mo_no:
                item["mos"].add(mo_no)
                item["sos"].update(mo_to_so.get(mo_no, set()))
                item["customers"].update(mo_to_customer.get(mo_no, set()))

    output = []
    for item in result.values():
        if item["unissued"] <= 0:
            continue
        item["mo_count"] = len(item["mos"])
        item["so_count"] = len(item["sos"])
        item["rate"] = (item["issued"] / item["demand"] * 100) if item["demand"] > 0 else 0.0
        item["mos"] = ",".join(sorted(item["mos"]))
        item["sos"] = ",".join(sorted(item["sos"]))
        item["customers"] = ",".join(sorted(item["customers"]))
        output.append(item)
    output.sort(key=lambda item: (-item["unissued"], -item["demand"]))
    return output


@mcp.tool()
def query_sale_orders(start_date: str, end_date: str) -> str:
    """查询金蝶云苍穹销售订单列表。"""
    try:
        token = get_token()
        so_rows = fetch_sale_orders(token, start_date, end_date)
        if not so_rows:
            return json.dumps({"status": "success", "message": "无销售订单数据", "count": 0, "data": []}, ensure_ascii=False)

        summary = []
        for row in so_rows:
            entries = row.get("billentry", []) or []
            summary.append(
                {
                    "billno": row.get("billno", ""),
                    "bizdate": (row.get("bizdate", "") or "")[:10],
                    "billstatus": f"{SO_BILL_STATUS.get(row.get('billstatus', ''), row.get('billstatus', ''))}({row.get('billstatus', '')})",
                    "orderstatus": f"{SO_ORDER_STATUS.get(row.get('orderstatus', ''), row.get('orderstatus', ''))}({row.get('orderstatus', '')})",
                    "customer_name": row.get("customer_name", "") or "",
                    "entry_count": len(entries),
                    "total_qty": round(sum(safe_float(entry.get("qty")) for entry in entries), 2),
                    "total_amount": round(sum(safe_float(entry.get("amountandtax")) for entry in entries), 2),
                }
            )

        return json.dumps(
            {"status": "success", "count": len(summary), "date_range": f"{start_date} ~ {end_date}", "data": summary},
            ensure_ascii=False,
            indent=2,
        )
    except Exception as exc:
        return json.dumps({"status": "error", "message": str(exc)}, ensure_ascii=False)


@mcp.tool()
def query_mft_orders(start_date: str, end_date: str) -> str:
    """查询金蝶云苍穹生产工单列表。"""
    try:
        token = get_token()
        mo_rows = fetch_mft_orders(token, start_date, end_date)
        if not mo_rows:
            return json.dumps({"status": "success", "message": "无生产工单数据", "count": 0, "data": []}, ensure_ascii=False)

        summary = []
        for mo in mo_rows:
            entries = mo.get("treeentryentity", []) or []
            summary.append(
                {
                    "billno": mo.get("billno", ""),
                    "billdate": (mo.get("billdate", "") or "")[:10],
                    "billstatus": f"{MO_BILL_STATUS.get(mo.get('billstatus', ''), mo.get('billstatus', ''))}({mo.get('billstatus', '')})",
                    "org_name": mo.get("org_name", ""),
                    "entry_count": len(entries),
                    "total_qty": round(sum(safe_float(entry.get("qty")) for entry in entries), 2),
                    "total_rptqty": round(sum(safe_float(entry.get("rptqty")) for entry in entries), 2),
                    "is_done": mo.get("billstatus", "") in MO_DONE_STATUS,
                }
            )

        done_count = sum(1 for item in summary if item["is_done"])
        return json.dumps(
            {
                "status": "success",
                "count": len(summary),
                "done_count": done_count,
                "active_count": len(summary) - done_count,
                "date_range": f"{start_date} ~ {end_date}",
                "data": summary,
            },
            ensure_ascii=False,
            indent=2,
        )
    except Exception as exc:
        return json.dumps({"status": "error", "message": str(exc)}, ensure_ascii=False)


@mcp.tool()
def query_so_mo_relation(start_date: str, end_date: str, limit: int = 50) -> str:
    """查询销售订单与生产工单的关联关系，包含完工状态、交货风险和领料情况。"""
    try:
        token = get_token()
        so_rows = fetch_sale_orders(token, start_date, end_date)
        if not so_rows:
            return json.dumps({"status": "success", "message": "无销售订单数据"}, ensure_ascii=False)

        mo_rows = fetch_mft_orders(token, start_date, end_date)
        if not mo_rows:
            return json.dumps({"status": "success", "message": "有SO但无生产工单", "so_count": len(so_rows)}, ensure_ascii=False)

        stock_rows = fetch_mft_stock(token, start_date, end_date)
        relations = join_all(so_rows, mo_rows, build_picking_idx(stock_rows))

        linked = [row for row in relations if row["linked"]]
        done_count = sum(1 for row in relations if row["mo_is_done"])
        risk_count = defaultdict(int)
        for row in relations:
            risk_count[row["delivery_risk"]] += 1

        output = []
        for row in relations[:limit]:
            output.append(
                {
                    "mo_billno": row["mo_billno"],
                    "mo_billstatus": row["mo_billstatus"],
                    "mo_bizstatus": row["mo_bizstatus"],
                    "mo_material_name": row["mo_material_name"],
                    "mo_qty": row["mo_qty"],
                    "mo_completion": row["mo_completion"],
                    "mo_completion_rate": row["mo_completion_rate"],
                    "mo_is_done": row["mo_is_done"],
                    "mo_planend": row["mo_planend"],
                    "delivery_risk": row["delivery_risk"],
                    "delivery_risk_label": row["delivery_risk_label"],
                    "pick_rate": round(row["pick_rate"], 1),
                    "pick_demand": row["pick_demand"],
                    "pick_issued": row["pick_issued"],
                    "so_billno": row["so_billno"],
                    "so_customer_name": row["so_customer_name"],
                    "so_deliverydate": row["so_deliverydate"],
                    "so_qty": row["so_qty"],
                    "linked": row["linked"],
                }
            )

        return json.dumps(
            {
                "status": "success",
                "date_range": f"{start_date} ~ {end_date}",
                "summary": {
                    "total_relations": len(relations),
                    "linked_to_so": len(linked),
                    "unlinked": len(relations) - len(linked),
                    "done_count": done_count,
                    "active_count": len(relations) - done_count,
                    "risk_distribution": dict(risk_count),
                },
                "data": output,
            },
            ensure_ascii=False,
            indent=2,
        )
    except Exception as exc:
        return json.dumps({"status": "error", "message": str(exc)}, ensure_ascii=False)


@mcp.tool()
def query_delivery_risk(start_date: str, end_date: str) -> str:
    """查询存在交货风险的工单。"""
    try:
        token = get_token()
        so_rows = fetch_sale_orders(token, start_date, end_date)
        mo_rows = fetch_mft_orders(token, start_date, end_date)
        if not mo_rows:
            return json.dumps({"status": "success", "message": "无生产工单", "data": []}, ensure_ascii=False)

        stock_rows = fetch_mft_stock(token, start_date, end_date)
        relations = join_all(so_rows, mo_rows, build_picking_idx(stock_rows))
        urgent = [row for row in relations if row["delivery_risk"] in ("overdue", "at_risk")]
        urgent.sort(key=lambda row: (0 if row["delivery_risk"] == "overdue" else 1, row["pick_rate"]))

        output = []
        for row in urgent:
            output.append(
                {
                    "delivery_risk_label": row["delivery_risk_label"],
                    "mo_billno": row["mo_billno"],
                    "mo_billstatus": row["mo_billstatus"],
                    "mo_material_name": row["mo_material_name"],
                    "mo_qty": row["mo_qty"],
                    "mo_completion_rate": row["mo_completion_rate"],
                    "mo_rptqty": row["mo_rptqty"],
                    "mo_planend": row["mo_planend"],
                    "so_billno": row["so_billno"],
                    "so_customer_name": row["so_customer_name"],
                    "so_deliverydate": row["so_deliverydate"],
                    "pick_rate": round(row["pick_rate"], 1),
                    "pick_unissued": row["pick_unissued"],
                }
            )

        return json.dumps(
            {
                "status": "success",
                "date_range": f"{start_date} ~ {end_date}",
                "overdue_count": sum(1 for row in urgent if row["delivery_risk"] == "overdue"),
                "at_risk_count": sum(1 for row in urgent if row["delivery_risk"] == "at_risk"),
                "total_urgent": len(urgent),
                "data": output,
            },
            ensure_ascii=False,
            indent=2,
        )
    except Exception as exc:
        return json.dumps({"status": "error", "message": str(exc)}, ensure_ascii=False)


@mcp.tool()
def query_material_shortage(start_date: str, end_date: str, top_n: int = 20) -> str:
    """查询欠料/瓶颈物料汇总。"""
    try:
        token = get_token()
        so_rows = fetch_sale_orders(token, start_date, end_date)
        mo_rows = fetch_mft_orders(token, start_date, end_date)
        if not mo_rows:
            return json.dumps({"status": "success", "message": "无生产工单", "data": []}, ensure_ascii=False)

        stock_rows = fetch_mft_stock(token, start_date, end_date)
        relations = join_all(so_rows, mo_rows, build_picking_idx(stock_rows))
        bottleneck = build_bottleneck(stock_rows, relations)

        output = []
        for index, item in enumerate(bottleneck[:top_n], 1):
            output.append(
                {
                    "rank": index,
                    "mat_no": item["mat_no"],
                    "mat_name": item["mat_name"],
                    "demand": round(item["demand"], 2),
                    "issued": round(item["issued"], 2),
                    "unissued": round(item["unissued"], 2),
                    "rate": round(item["rate"], 1),
                    "mo_count": item["mo_count"],
                    "so_count": item["so_count"],
                    "mos": item["mos"],
                    "sos": item["sos"],
                    "customers": item["customers"],
                }
            )

        return json.dumps(
            {
                "status": "success",
                "date_range": f"{start_date} ~ {end_date}",
                "total_shortage_materials": len(bottleneck),
                "showing_top": min(top_n, len(bottleneck)),
                "data": output,
            },
            ensure_ascii=False,
            indent=2,
        )
    except Exception as exc:
        return json.dumps({"status": "error", "message": str(exc)}, ensure_ascii=False)


@mcp.tool()
def export_so_mo_report(start_date: str, end_date: str) -> str:
    """导出完整的 SO-MO 关联报告为 Excel 文件。"""
    try:
        token = get_token()
        so_rows = fetch_sale_orders(token, start_date, end_date)
        if not so_rows:
            return json.dumps({"status": "error", "message": "无销售订单数据"}, ensure_ascii=False)

        mo_rows = fetch_mft_orders(token, start_date, end_date)
        if not mo_rows:
            return json.dumps({"status": "error", "message": "无生产工单数据"}, ensure_ascii=False)

        stock_rows = fetch_mft_stock(token, start_date, end_date)
        relations = join_all(so_rows, mo_rows, build_picking_idx(stock_rows))
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        file_name = f"SO_MO完工领料_{start_date}_{end_date}_{timestamp}.xlsx"
        file_path = export_excel(relations, stock_rows, file_name, EXPORT_DIR)
        linked_count = sum(1 for row in relations if row["linked"])
        done_count = sum(1 for row in relations if row["mo_is_done"])
        return json.dumps(
            {
                "status": "success",
                "file_path": file_path,
                "file_name": file_name,
                "summary": {
                    "total_relations": len(relations),
                    "linked_to_so": linked_count,
                    "done_count": done_count,
                    "active_count": len(relations) - done_count,
                },
                "sheets": [
                    "SO-MO关联明细",
                    "按SO汇总",
                    "工单完工明细",
                    "逾期风险工单",
                    "欠料明细",
                    "瓶颈物料汇总",
                ],
            },
            ensure_ascii=False,
            indent=2,
        )
    except Exception as exc:
        return json.dumps({"status": "error", "message": str(exc)}, ensure_ascii=False)


def main():
    mcp.run(transport="stdio")


if __name__ == "__main__":
    main()
