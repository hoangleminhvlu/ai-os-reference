"""Máy chủ MCP mẫu cho phần mềm CRM. Chạy: uvicorn vendor_mcp_server:app --port 9100"""
import hmac
import json
import os
from fastapi import FastAPI, Header, HTTPException, Request
from fastapi.responses import JSONResponse, Response

app = FastAPI(title="Vendor CRM MCP Server")
SERVICE_TOKEN = os.environ["CRM_MCP_TOKEN"]   # bí mật nằm ngoài mã nguồn
PROTOCOL_VERSION = "2025-06-18"

# Thay bằng truy vấn API hoặc CSDL bản sao chỉ đọc của sản phẩm
CUSTOMERS = {
    "KH001": {"id": "KH001", "name": "Công ty An Phát", "segment": "SME", "open_tickets": 2},
}

TOOLS = [{
    "name": "get_customer_summary",
    "description": "Tra cứu tóm tắt một khách hàng theo mã (chỉ đọc).",
    "inputSchema": {
        "type": "object",
        "additionalProperties": False,
        "required": ["customer_id"],
        "properties": {"customer_id": {"type": "string", "maxLength": 32}},
    },
    "annotations": {"readOnlyHint": True, "openWorldHint": False},
}]


def text_result(text: str, is_error: bool = False) -> dict:
    return {"isError": is_error, "content": [{"type": "text", "text": text}]}


def call_tool(name: str, args: dict) -> dict:
    if name != "get_customer_summary":
        raise KeyError(name)
    customer = CUSTOMERS.get(str(args.get("customer_id", "")))
    if customer is None:   # lỗi nghiệp vụ: để mô hình đọc được và tự xử lý
        return text_result("Không tìm thấy khách hàng.", is_error=True)
    return text_result(json.dumps(customer, ensure_ascii=False))


def rpc_error(msg_id, code: int, message: str) -> JSONResponse:
    return JSONResponse({"jsonrpc": "2.0", "id": msg_id,
                         "error": {"code": code, "message": message}})


@app.post("/mcp")
async def mcp(request: Request, authorization: str = Header("")):
    if not hmac.compare_digest(authorization, f"Bearer {SERVICE_TOKEN}"):
        raise HTTPException(status_code=401, detail="Sai mã truy cập")
    message = await request.json()
    if "id" not in message:   # thông báo một chiều (vd. notifications/initialized)
        return Response(status_code=202)
    msg_id, method = message["id"], message.get("method")
    params = message.get("params") or {}
    if method == "initialize":
        result = {"protocolVersion": PROTOCOL_VERSION, "capabilities": {"tools": {}},
                  "serverInfo": {"name": "vendor-crm", "version": "1.0.0"}}
    elif method == "ping":
        result = {}
    elif method == "tools/list":
        result = {"tools": TOOLS}
    elif method == "tools/call":
        try:
            result = call_tool(params.get("name"), params.get("arguments") or {})
        except KeyError:
            return rpc_error(msg_id, -32602, "Công cụ không tồn tại")
    else:
        return rpc_error(msg_id, -32601, "Phương thức không hỗ trợ")
    return JSONResponse({"jsonrpc": "2.0", "id": msg_id, "result": result})
