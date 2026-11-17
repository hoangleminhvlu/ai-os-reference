# -*- coding: utf-8 -*-
"""
AI Gateway giả lập để chụp ảnh màn hình cho TL1, TL2 (dữ liệu mẫu, không có dữ liệu thật):
router THẬT của Định danh (OpenLDAP giả lập), Trao đổi trực tuyến (kho tin nhắn trong bộ nhớ), Trạng thái AI Platform;
dữ liệu mẫu cho Kho, mô hình, giám sát; /api/chat/rag-stream trả SSE đúng định dạng thật (meta, nguồn, token).
"""
import asyncio
import json
import os
import runpy
import sys
import tempfile
import time
from datetime import datetime, timedelta, timezone

UI_DIR = "/home/local/ai-platform/tests/identity/ui"
sys.path.insert(0, "/home/local/ai-platform")
sys.path.insert(0, UI_DIR)
os.chdir(os.path.dirname(os.path.abspath(__file__)))
ns = runpy.run_path(os.path.join(UI_DIR, "mock_gateway.py"), run_name="mock_gateway_lib")
app, M = ns["app"], ns["M"]
DROP = {"/api/{rest:path}", "/api/auth/me", "/api/chat/threads/all", "/api/chat/history"}
app.router.routes = [r for r in app.router.routes if getattr(r, "path", "") not in DROP]

from fastapi import Body, Header, HTTPException  # noqa: E402
from fastapi.responses import StreamingResponse  # noqa: E402
import config as Config  # noqa: E402

Config.REDIS_URL = ""
Config.COMM_ATTACHMENT_DIR = tempfile.mkdtemp(prefix="comm_att_")
Config.COMM_TURN_URLS = ["turn:turn.ai-tech.vn:3478"]

# Người dùng mẫu thêm vào OpenLDAP giả lập (tên mẫu, không phải người thật)
M.add(1101, ["AI-TECH:member"], "Trần Minh Khoa", iDAccountOwnerID="NV011", iDAccountStatus="555111", mailForwardingAddress=["khoa.tm@ai-tech.vn"])
M.add(1102, ["AI-TECH:active"], "Lê Thị Hồng Nhung", iDAccountOwnerID="GV012", iDAccountStatus="555222", mailForwardingAddress=["nhung.lth@ai-tech.vn"])
NAMES = {"1001": ("Nguyễn Văn An", "admin", "NV001"), "1007": ("Phạm Thu Dung", "partner", "DT001"),
         "1101": ("Trần Minh Khoa", "member", "NV011"), "1102": ("Lê Thị Hồng Nhung", "active", "GV012")}


def user_of(authorization):
    uid = (authorization or "").split(" ", 1)[-1]
    if uid not in NAMES:
        raise HTTPException(401, "Thiếu mã truy cập")
    name, role, owner = NAMES[uid]
    return {"user_id": uid, "user_name": name, "user_role": role, "user_org": Config.ORGNAME, "user_ip": "127.0.0.1", "owner_id": owner}


@app.get("/api/auth/me")
def me(authorization: str = Header(None)):
    u = user_of(authorization)
    return {**u, "org_id": "AI-TECH", "org_name": "AI-TECH Center - Vanlang University", "chat_history_enabled": True,
            "chat_history_roles": ["admin", "member", "active", "partner"], "store_info_editable": u["user_role"] == "admin",
            "identity_permissions": ["identity.users.read"] if u["user_role"] == "admin" else []}


@app.get("/api/org/info")
def org_info():
    return {"status": "SUCCESS", "org_id": "AI-TECH", "org_name": "AI-TECH Center - Vanlang University", "org_url": "https://ai-tech.vlu.edu.vn",
            "location": "FITLAB-03", "workspace_url": "https://aiwork.is-tech.vn", "gateway_url": "https://aigate.is-tech.vn",
            "network": "internet", "switch_enabled": True,
            "organizations": [
                {"org_id": "AI-TECH", "org_name": "AI-TECH Center - Vanlang University", "workspace_url": "https://aiwork.is-tech.vn", "current": True},
                {"org_id": "VLSC", "org_name": "Trường Cao đẳng Văn Lang Sài Gòn", "workspace_url": "https://vlsc.is-tech.vn", "current": False},
                {"org_id": "TOBOIT", "org_name": "Công ty TOBO IT", "workspace_url": "https://toboit.is-tech.vn", "current": False}]}


def cloud(gid, name, desc, perm="manage", opened=False, visibility="Public"):
    return {"storage_key": f"Projects/{gid}", "scope": "cloud", "display_name": f"Dự án: {name}", "description": desc,
            "collection_name": f"prj_ai_tech_{gid}", "permission": perm, "opened": opened,
            "cloud": {"name": name, "description": desc, "gid": gid, "branch": "Projects", "visibility": visibility,
                      "primary_type": "Project", "members": 6, "status": visibility}}


COLS = [
    {"storage_key": "Personal", "scope": "personal", "display_name": "Kho Cá nhân", "collection_name": "personal_knowledge",
     "permission": "manage", "file_count": 12, "chunk_count": 384, "storage_size": "2.41 MB", "opened": False,
     "description": "Tài liệu riêng của bạn, chỉ bạn khai thác."},
    {"storage_key": "Public", "scope": "org", "display_name": "Công khai", "collection_name": "public_knowledge",
     "permission": "manage", "file_count": 57, "chunk_count": 4210, "storage_size": "31.8 MB", "opened": False,
     "description": "Quy chế, quy trình, biểu mẫu dùng chung của Tổ chức."},
    {"storage_key": "Tech", "scope": "org", "display_name": "Kỹ thuật", "collection_name": "tech_knowledge",
     "permission": "manage", "file_count": 31, "chunk_count": 2210, "storage_size": "14.2 MB", "opened": False,
     "description": "Tài liệu kỹ thuật, hướng dẫn vận hành hệ thống."},
    {**cloud(1161000, "251_AIDMS_TC01", "Hệ thống quản lý tài liệu thông minh AIDMS"), "file_count": 23, "chunk_count": 1302, "storage_size": "9.6 MB"},
]
PUBLIC_CLOUDS = [
    cloud(1161002, "PUBLIC_DEMO", "Kho học liệu mở về Trí tuệ nhân tạo cho sinh viên", "read"),
    cloud(1161004, "252_LLM_LAB", "Phòng thí nghiệm mô hình ngôn ngữ lớn", "read"),
    cloud(1161005, "LAB0_TC01", "Phòng thí nghiệm AI số 0", "read"),
    cloud(1161007, "252_SMARTCAMPUS", "Đề tài Khuôn viên thông minh", "read"),
    cloud(1161009, "251_DIGITAL_LIB", "Thư viện số Khoa Công nghệ Thông tin", "read"),
]


@app.get("/api/rag/collections")
def collections(authorization: str = Header(None)):
    user_of(authorization)
    return COLS


@app.get("/api/rag/public-clouds")
def public_clouds(q: str = "", include_private: bool = False, authorization: str = Header(None)):
    user_of(authorization)
    items = [c for c in PUBLIC_CLOUDS if not q or q.lower() in json.dumps(c, ensure_ascii=False).lower()]
    return {"status": "SUCCESS", "clouds": items, "max_opened": 10, "can_include_private": True}


@app.get("/api/data/files")
def data_files(domain: str, subdir: str = "", authorization: str = Header(None)):
    user_of(authorization)
    f = lambda name, size, state, chunks, mod: {
        "path": f"{subdir}/{name}" if subdir else name, "name": name, "folder": subdir, "ext": os.path.splitext(name)[1], "size": 1,
        "size_display": size, "modified": mod, "chunk_count": chunks, "last_sync": "2026-11-10T08:05:00Z" if chunks else None,
        "sync_state": state, "rag_id": name if chunks else None}
    files = [f("Quy_che_thuc_tap_2026.pdf", "793.3 KB", "synced", 64, "2026-11-10T08:00:00Z"),
             f("Huong_dan_GVHD_thuc_tap.docx", "117.2 KB", "outdated", 12, "2026-11-15T08:00:00Z"),
             f("Bien_ban_hop_16-11.md", "4.1 KB", "not_synced", 0, "2026-11-16T08:00:00Z")]
    folders = [] if subdir else [{"name": "Bao_cao_2026", "path": "Bao_cao_2026", "rag_files": 7},
                                 {"name": "Bieu_mau", "path": "Bieu_mau", "rag_files": 0}]
    rag_only = [] if subdir else [{"id": "r9", "name": "Thong_bao_lich_bao_ve.pdf", "path": "Thong_bao_lich_bao_ve.pdf", "origin": "deleted",
                                   "chunk_count": 18, "size_display": "88 KB", "last_sync": "2026-10-01T08:00:00Z"}]
    return {"status": "SUCCESS", "subdir": subdir, "folders": folders, "files": files, "rag_only": rag_only}


MODELS = [
    {"name": "qwen3:32b", "runtime": "dgx", "runtime_name": "Máy chủ AI DGX Spark", "capabilities": ["completion", "tools"], "loaded": True,
     "parameter_size": "32.8B", "location": "onpremise"},
    {"name": "gemma4:latest", "runtime": "dgx", "runtime_name": "Máy chủ AI DGX Spark", "capabilities": ["completion", "tools"], "loaded": False,
     "parameter_size": "12B", "location": "onpremise"},
    {"name": "gpt-oss:120b-cloud", "runtime": "dgx", "runtime_name": "Máy chủ AI DGX Spark", "capabilities": ["completion", "tools"],
     "loaded": False, "parameter_size": "116.8B", "location": "onclouds"},
]


@app.get("/api/models")
def models(authorization: str = Header(None)):
    return {"status": "SUCCESS", "models": MODELS, "runtimes": [{"id": "dgx", "name": "Máy chủ AI DGX Spark", "location": "onpremise"}],
            "default_model": "qwen3:32b", "default_runtime": "dgx"}


SOURCES_MARK = "\n\n---\n📚 **Nguồn đã tra cứu trong Kho Tri thức:**\n"
ANSWER = ("Chào **An**! Theo Quy chế thực tập năm 2026, sinh viên thực hiện hai đợt thực tập [1]:\n\n"
          "| Đợt | Thời lượng | Yêu cầu chính |\n|---|---|---|\n"
          "| Thực tập cơ sở | 4 tuần | Báo cáo tuần, nhận xét của đơn vị [1] |\n"
          "| Thực tập tốt nghiệp | 8 tuần | Báo cáo cuối kỳ, bảo vệ trước hội đồng [1] |\n\n"
          "Giảng viên hướng dẫn chấm điểm quá trình theo phiếu đánh giá trong Hướng dẫn GVHD [2]; điểm cuối kỳ do hội đồng quyết định.")


@app.post("/api/chat/rag-stream")
async def rag_stream(body: dict = Body(...), authorization: str = Header(None)):
    user_of(authorization)
    model = body.get("aimodel") or "qwen3:32b"
    info = next((m for m in MODELS if m["name"] == model), MODELS[0])
    start = datetime.now(timezone.utc)
    meta = {"model": model, "runtime": "dgx", "runtime_name": info["runtime_name"], "location": info["location"],
            "started_at": start.isoformat(timespec="seconds")}

    def sse(obj):
        return f"data: {json.dumps(obj, ensure_ascii=False)}\n\n"

    async def gen():
        yield sse({"meta": {"event": "start", **meta}})
        yield sse({"choices": [{"delta": {"content": "📄 *Đã đọc riêng tài liệu: Quy_che_thuc_tap_2026.pdf.*\n\n🔍 *Đã tra thêm Kho Tri thức: Kỹ thuật (theo từ khóa).*\n\n"}}]})
        await asyncio.sleep(0.4)
        for i in range(0, len(ANSWER), 24):
            yield sse({"choices": [{"delta": {"content": ANSWER[i:i + 24]}}]})
            await asyncio.sleep(0.01)
        yield sse({"choices": [{"delta": {"content": SOURCES_MARK + "- [1] Quy_che_thuc_tap_2026.pdf (Kho Công khai)\n- [2] Huong_dan_GVHD_thuc_tap.docx (Kho Công khai)\n"}}]})
        end = start + timedelta(milliseconds=8240)
        yield sse({"meta": {"event": "end", **meta, "finished_at": end.isoformat(timespec="seconds"), "duration_ms": 8240,
                            "first_token_ms": 1320, "status": "ok", "sources": 2,
                            "usage": {"input_tokens": 2140, "output_tokens": 186, "total_tokens": 2326}}})
        yield "data: [DONE]\n\n"
    return StreamingResponse(gen(), media_type="text/event-stream")


# Lịch sử Làm việc
THREADS = [
    {"thread_id": "t1", "title": "Quy chế thực tập năm 2026 quy định những gì?", "domain": "Public", "created_at": "2026-11-16T02:00:00",
     "updated_at": "2026-11-16T02:30:00",
     "messages": [("user", "Quy chế thực tập năm 2026 quy định những gì?", None),
                  ("assistant", ANSWER + SOURCES_MARK + "- [1] Quy_che_thuc_tap_2026.pdf (Kho Công khai)\n- [2] Huong_dan_GVHD_thuc_tap.docx (Kho Công khai)\n",
                   {"model": "qwen3:32b", "runtime_name": "Máy chủ AI DGX Spark", "location": "onpremise", "started_at": "2026-11-16T02:30:01+00:00",
                    "finished_at": "2026-11-16T02:30:09+00:00", "duration_ms": 8240, "status": "ok", "sources": 2,
                    "usage": {"input_tokens": 2140, "output_tokens": 186, "total_tokens": 2326}})]},
    {"thread_id": "t2", "title": "Tóm tắt đề cương dự án AIDMS", "domain": "Projects/1161000", "created_at": "2026-11-14T08:00:00",
     "updated_at": "2026-11-14T09:30:00", "messages": [("user", "Tóm tắt đề cương dự án AIDMS", None), ("assistant", "**Đề cương** gồm 3 phần [1]...", None)]},
    {"thread_id": "t3", "title": "Cấu hình nginx cho AI Gateway", "domain": "Tech", "created_at": "2026-11-12T03:00:00",
     "updated_at": "2026-11-12T03:20:00", "messages": [("user", "Cấu hình nginx cho AI Gateway?", None), ("assistant", "Cần chuyển tiếp /api và /ws [1]...", None)]},
    {"thread_id": "t4", "title": "Kho Cá nhân có những tài liệu nào?", "domain": "Personal", "created_at": "2026-11-10T07:00:00",
     "updated_at": "2026-11-10T07:05:00", "messages": [("user", "Kho có những tài liệu nào?", None), ("assistant", "Kho có 12 tệp [0]...", None)]},
]


@app.get("/api/chat/threads/all")
def threads_all(q: str = "", domain: str = "", authorization: str = Header(None)):
    user_of(authorization)
    items = [{**{k: v for k, v in t.items() if k != "messages"}, "message_count": len(t["messages"])} for t in THREADS
             if (not q or q.lower() in t["title"].lower()) and (not domain or t["domain"] == domain)]
    return {"items": items, "chat_history_enabled": True}


@app.get("/api/chat/history")
def thread_history(thread_id: str, authorization: str = Header(None)):
    user_of(authorization)
    t = next((t for t in THREADS if t["thread_id"] == thread_id), None)
    return [{"role": r, "content": c, "timestamp": t["updated_at"], "metadata": m} for r, c, m in t["messages"]] if t else []


@app.get("/api/system/metrics")
def metrics(authorization: str = Header(None)):
    host = {"ok": True, "hostname": "dgx-spark-01", "uptime_seconds": 1234567, "cpu": {"percent": 37.5, "cores": 20},
            "memory": {"used_gb": 78.2, "total_gb": 128, "unified": True},
            "gpus": [{"index": 0, "name": "NVIDIA GB10", "load_percent": 62, "temperature_c": 58, "vram_used_gb": 41.3, "vram_total_gb": 128, "power_w": 96}],
            "disks": [{"mount": "/", "used_gb": 1320, "total_gb": 3700}], "errors": []}
    rt = {"id": "dgx", "name": "Máy chủ AI DGX Spark", "ollama_url": "http://10.0.0.5:11434", "location": "onpremise", "sources": ["metrics_agent"],
          "ollama": {"url": "http://10.0.0.5:11434", "ok": True, "version": "0.12.6"},
          "models": {"ok": True, "installed": [
              {"name": "qwen3:32b", "parameter_size": "32.8B", "quantization": "Q4_K_M", "size_gb": 20.2},
              {"name": "gemma4:latest", "parameter_size": "12B", "quantization": "Q4_K_M", "size_gb": 8.1},
              {"name": "gpt-oss:120b-cloud", "parameter_size": "116.8B", "quantization": "MXFP4", "size_gb": 0.0},
              {"name": "bge-m3:latest", "parameter_size": "567M", "quantization": "F16", "size_gb": 1.2}],
              "loaded": [{"name": "qwen3:32b", "vram_gb": 24.6}, {"name": "bge-m3:latest", "vram_gb": 1.4}]},
          "monitored": True, "host": host}
    return {"status": "SUCCESS", "timestamp": datetime.now().astimezone().isoformat(timespec="seconds"), "runtimes": [rt], "config_error": None}


@app.get("/api/system/logs")
def logs(lines: int = 100, authorization: str = Header(None)):
    now = datetime.now()
    rows = [("INFO", "[iAIP RAG] Đã ghi log cuộc chat cho user 1001"), ("INFO", "[RAG Directory Sync] Đã RAG 7 tệp trong thư mục Bao_cao_2026"),
            ("WARNING", "[Login Guard] Nhập sai PassCode 3 lần từ 10.0.12.25"), ("INFO", "[Comm] Trợ lý AI trả lời trong hội thoại 2"),
            ("INFO", "[AI Server] Máy chủ AI DGX Spark: 2 mô hình đang nạp")]
    return {"status": "SUCCESS", "lines": [{"level": lv, "text": f"{(now - timedelta(minutes=5 - i)).strftime('%Y-%m-%d %H:%M:%S')} {lv} {t}"}
                                            for i, (lv, t) in enumerate(rows)]}


# Trạng thái AI Platform: router thật, kết quả kiểm tra mẫu
from platform_status import router as PS  # noqa: E402


def _ok(detail):
    return lambda: {"detail": detail}


async def _emb():
    return {"detail": "Ollama 0.12.6 · tạo vector cho Kho Tri thức (RAG)"}


def _mcp():
    return {"detail": "1 MCP Server", "items": [{"id": "zalo", "name": "Zalo MCP", "target": "127.0.0.1:9110", "state": "ok",
                                                  "detail": "6 công cụ", "checked_at": "2026-11-17T01:00:00Z"}]}


for _id, _fn in {"postgres": _ok("PostgreSQL 18.6 · pgvector 0.8.1 · CSDL 1.8 GB"),
                 "ldap": _ok("Đã xác thực tài khoản quản trị · dc=idragon,dc=vn · TLS: StartTLS"),
                 "redis": _ok("Redis 7.4.1 · 12 kết nối"), "embedding": _emb,
                 "smtp": _ok("Máy chủ Email sẵn sàng (chỉ kiểm tra kết nối, không đăng nhập)"),
                 "storage": _ok("Còn trống 2,1 TB / 3,6 TB (58%)"), "mcp": _mcp}.items():
    n, r, _old, t = PS.CHECKS[_id]
    PS.CHECKS[_id] = (n, r, _fn, t)
PS.CHECKS["postgres"] = (*PS.CHECKS["postgres"][:3], lambda: "10.0.0.10:5432/aitech_db")
PS.CHECKS["ldap"] = (*PS.CHECKS["ldap"][:3], lambda: "10.0.0.11:389")
PS.CHECKS["redis"] = (*PS.CHECKS["redis"][:3], lambda: "10.0.0.10:6379")
PS.CHECKS["embedding"] = (*PS.CHECKS["embedding"][:3], lambda: "10.0.0.5:11434")
PS.CHECKS["smtp"] = (*PS.CHECKS["smtp"][:3], lambda: "mail.ai-tech.vn:587")
PS.CHECKS["storage"] = (*PS.CHECKS["storage"][:3], lambda: "/export/community/BlendedLAB/AI-TECH")
Config.TRUSTED_PROXIES_CONFIGURED = False
Config.JWT_SECRET_KEY = "x" * 64
Config.LDAP_SERVER = "127.0.0.1"
PS.platform_info = (lambda orig: (lambda: {**orig(), "hostname": "aiplatform-01", "pid": 24816, "os": "Linux 6.8.0-ubuntu", "python": "3.12.3",
                                            "uptime_seconds": 734400, "started_at": "2026-11-08T23:00:00+00:00",
                                            "workspace_url": "https://aiwork.is-tech.vn", "gateway_url": "https://aigate.is-tech.vn",
                                            "gateway_network": "internet", "location": "FITLAB-03"}))(PS.platform_info)
app.include_router(PS.create_platform_status_router(lambda authorization=Header(None): user_of(authorization)))

# Trao đổi trực tuyến: router thật, kho trong bộ nhớ
from comm_gateway import directory as D, store as S  # noqa: E402
from comm_gateway.hub import hub  # noqa: E402
from comm_gateway.router import create_comm_router  # noqa: E402


class Dir:
    def search(self, q, exclude_uid=""):
        return [{"user_id": u, "user_name": v[0], "title": "", "role": v[1]} for u, v in NAMES.items() if u != exclude_uid and q.lower() in v[0].lower()]

    def lookup(self, uids):
        return {u: {"user_name": NAMES[u][0], "role": NAMES[u][1], "ou": []} for u in uids if u in NAMES}

    def manageable_units(self, user):
        return None

    def list_units(self):
        return []

    def unit_members(self, unit):
        return None

    def cloud_readable(self, unit, roles):
        return True


D.directory = Dir()
S.store = st = S.MemoryCommStore()
N = lambda u: NAMES[u][0]
AI_GROUP = ("Chào **An**, **Dung**, **Khoa** và **Nhung**! Theo Quy chế thực tập năm 2026, thực tập tốt nghiệp kéo dài **8 tuần** [1]; "
            "báo cáo cuối kỳ nộp chậm nhất 1 tuần sau khi kết thúc đợt thực tập [1].")
META = {"model": "qwen3:32b", "runtime": "dgx", "runtime_name": "Máy chủ AI DGX Spark", "location": "onpremise", "domains": ["Public"],
        "started_at": (datetime.now(timezone.utc) - timedelta(seconds=40)).isoformat(timespec="seconds"),
        "finished_at": (datetime.now(timezone.utc) - timedelta(seconds=37)).isoformat(timespec="seconds"), "duration_ms": 3410, "first_token_ms": 980,
        "status": "ok", "sources": 1, "usage": {"input_tokens": 1650, "output_tokens": 64, "total_tokens": 1714}}
g, _ = st.create_conversation("group", "Nhóm Thực tập 2026", "1001", [("1001", N("1001"), "owner"), ("1007", N("1007"), "member"),
                              ("1101", N("1101"), "member"), ("1102", N("1102"), "member")], metadata={"domain": "Public"})
st.insert_message(g["id"], "1102", N("1102"), "user", "text", "Chào cả nhóm, tuần sau các em nộp báo cáo giữa kỳ nhé!")
st.insert_message(g["id"], "1101", N("1101"), "user", "text", "Dạ cô, báo cáo theo mẫu nào ạ?")
q, _ = st.insert_message(g["id"], "1001", N("1001"), "user", "text", "@ai thời gian thực tập tốt nghiệp và hạn nộp báo cáo cuối kỳ là khi nào?")
st.insert_message(g["id"], None, "Trợ lý AI", "assistant", "text",
                  AI_GROUP + SOURCES_MARK + "- [1] Quy_che_thuc_tap_2026.pdf (Kho Công khai)\n", q["id"], [], None, META)
d, _ = st.create_conversation("direct", "", "1001", [("1001", N("1001"), "member"), ("1007", N("1007"), "member")], direct_key="1001:1007")
st.insert_message(d["id"], "1007", N("1007"), "user", "text", "Chiều nay anh An có rảnh trao đổi về hợp đồng AIDMS không?")
st.insert_message(d["id"], "1001", N("1001"), "user", "text", "Được chị, 15 giờ mình gọi video nhé.")
a, _ = st.create_conversation("assistant", "Trợ lý AI", "1001", [("1001", N("1001"), "owner")], metadata={"domain": "Public"})


async def fake_assistant(request):
    request.answer_meta = {"model": request.aimodel, "runtime": "dgx", "runtime_name": "Máy chủ AI DGX Spark", "location": "onpremise",
                           "started_at": datetime.now(timezone.utc).isoformat(timespec="seconds")}

    def gen():
        text = AI_GROUP + SOURCES_MARK + "- [1] Quy_che_thuc_tap_2026.pdf (Kho Công khai)\n"
        for i in range(0, len(text), 30):
            time.sleep(0.02)
            yield text[i:i + 30]
        request.answer_meta.update(finished_at=datetime.now(timezone.utc).isoformat(timespec="seconds"), duration_ms=3410, status="ok",
                                   sources=1, usage={"input_tokens": 1650, "output_tokens": 64, "total_tokens": 1714})
    return gen()


hub.assistant = fake_assistant
app.include_router(create_comm_router(lambda authorization=Header(None): user_of(authorization), lambda *x: None,
                                      ["http://localhost:5175", "http://127.0.0.1:5175"]))


@app.get("/api/{rest:path}")
def stub(rest: str):
    return []


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="127.0.0.1", port=8098, log_level="warning")
