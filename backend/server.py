import os
import time
import json
import shutil
import secrets
import tempfile
import threading
import zipfile
from datetime import datetime
from typing import List, Optional
from fastapi import FastAPI, UploadFile, File, Form, HTTPException, BackgroundTasks, Header, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
import huliot_core

app = FastAPI(title="VIEW SUMMARY STUDIO - BOQ Intelligence Engine")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
    expose_headers=["X-Summary-Verified", "X-Summary-Warnings", "X-Agent-Log"],
)


@app.middleware("http")
async def disable_static_cache(request, call_next):
    response = await call_next(request)
    if not request.url.path.startswith("/api/"):
        response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
    return response

WORKSPACE_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
RAW_BOQ_DIR = os.path.abspath(os.path.join(WORKSPACE_DIR, "data", "raw_boq"))
KNOWLEDGE_PATH = os.path.abspath(os.path.join(WORKSPACE_DIR, "data", "trained_formats", "knowledge.json"))
WEB_BUILD_DIR = os.path.abspath(os.path.join(WORKSPACE_DIR, "huliot_flutter_app", "build", "web"))
ADMIN_PASSWORD = os.environ.get("ADMIN_PASSWORD")
MAX_ADMIN_ATTEMPTS = 5
ADMIN_LOCKOUT_SECONDS = 15 * 60  # 15 minutes
_ADMIN_LOCK = threading.Lock()
_ADMIN_ATTEMPTS = {}  # {ip: {"count": int, "locked_until": float}}


def _get_client_ip(request: Optional[Request]) -> str:
    if request is None:
        return "local"
    xff = request.headers.get("x-forwarded-for")
    if xff:
        return xff.split(",")[0].strip()
    return request.client.host if request.client else "local"


def _verify_admin(password: Optional[str], request: Optional[Request] = None):
    if not ADMIN_PASSWORD:
        raise HTTPException(
            status_code=503,
            detail="ADMIN_PASSWORD environment variable is not configured on the server.",
        )
    ip = _get_client_ip(request)
    now = time.time()
    with _ADMIN_LOCK:
        rec = _ADMIN_ATTEMPTS.get(ip, {"count": 0, "locked_until": 0.0})
        if rec["locked_until"] > now:
            rem_sec = int(rec["locked_until"] - now)
            rem_min = max(1, (rem_sec + 59) // 60)
            raise HTTPException(
                status_code=429,
                detail=f"5 बार गलत Password डालने के कारण 15 मिनट के लिए Lock कर दिया गया है ({rem_min} मिनट बाकी)।",
            )

        is_valid = bool(password) and secrets.compare_digest(password, ADMIN_PASSWORD)
        if not is_valid:
            rec["count"] = rec.get("count", 0) + 1
            if rec["count"] >= MAX_ADMIN_ATTEMPTS:
                rec["locked_until"] = now + ADMIN_LOCKOUT_SECONDS
                rec["count"] = 0
                _ADMIN_ATTEMPTS[ip] = rec
                raise HTTPException(
                    status_code=429,
                    detail="5 बार गलत Password डालने के कारण 15 मिनट के लिए Lock कर दिया गया है।",
                )
            left = MAX_ADMIN_ATTEMPTS - rec["count"]
            _ADMIN_ATTEMPTS[ip] = rec
            raise HTTPException(
                status_code=401,
                detail=f"गलत Admin Password (बाकी प्रयास: {left})",
            )

        # Valid password -> reset attempts
        _ADMIN_ATTEMPTS.pop(ip, None)


@app.get("/health")
def health_check():
    """Health check endpoint reporting app live status and total trained formats."""
    db = huliot_core.load_knowledge_db(KNOWLEDGE_PATH)
    total_formats = len(db.get("learned_projects", []))
    return {
        "status": "ok",
        "message": "VIEW SUMMARY STUDIO App Live hai",
        "total_formats": total_formats,
        "timestamp": datetime.now().isoformat(timespec="seconds"),
    }


@app.get("/api/status")
def get_status():
    knowledge_count = 0
    learned_titles = []
    if os.path.exists(KNOWLEDGE_PATH):
        try:
            import json
            with open(KNOWLEDGE_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
                knowledge_count = len(data.get("learned_projects", []))
                learned_titles = [p.get("project_title", "Unknown") for p in data.get("learned_projects", [])]
        except Exception:
            pass

    return {
        "status": "online",
        "engine": "Huliot MEP Estimator & Summary Generation Engine",
        "learned_projects_count": knowledge_count,
        "learned_projects": learned_titles,
        "workspace": WORKSPACE_DIR
    }


@app.post("/api/analyze")
async def analyze_file(file: UploadFile = File(...)):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        shutil.copyfileobj(file.file, tmp)
        tmp_path = tmp.name

    try:
        res = huliot_core.analyze_workbook(tmp_path, knowledge_db_path=KNOWLEDGE_PATH, original_filename=file.filename)
        return res
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
    finally:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)


@app.post("/api/train")
async def train_file(file: UploadFile = File(...)):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        shutil.copyfileobj(file.file, tmp)
        tmp_path = tmp.name

    try:
        res = huliot_core.train_from_workbook(tmp_path, knowledge_db_path=KNOWLEDGE_PATH, original_filename=file.filename)
        return res
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
    finally:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)


def _safe_remove(path: str):
    try:
        if path and os.path.exists(path):
            os.remove(path)
    except Exception:
        pass


@app.post("/api/generate")
async def generate_file(
    background_tasks: BackgroundTasks,
    file: UploadFile = File(...),
    project_name: str = Form(None),
    num_toilets: int = Form(None),
    num_kitchens: int = Form(None),
    num_floors: int = Form(None),
    num_towers: int = Form(None),
    discount_p: float = Form(0.55),
    discount_f: float = Form(0.55),
    discount_m: float = Form(0.55),
    discount_l: float = Form(0.45),
    discount_c: float = Form(0.45),
    include_rainwater: bool = Form(True)
):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_in:
        shutil.copyfileobj(file.file, tmp_in)
        in_path = tmp_in.name

    out_path = in_path.replace(".xlsx", "_summary_generated.xlsx")

    try:
        discounts = {
            "P": discount_p,
            "F": discount_f,
            "M": discount_m,
            "L": discount_l,
            "C": discount_c
        }
        huliot_core.generate_summary_sheet(
            in_path,
            out_path,
            project_name=project_name,
            num_toilets=num_toilets,
            num_kitchens=num_kitchens,
            num_floors=num_floors,
            num_towers=num_towers,
            discounts=discounts,
            original_filename=file.filename,
            include_rainwater=include_rainwater
        )

        # Step 4: Verify generated formulas before serving download
        verification = huliot_core.verify_generated_summary(out_path)

        # Step 2 (#2): Auto-Learn on Generate (persist user-confirmed discounts & quantities)
        try:
            huliot_core.learn_from_generation(
                out_path,
                knowledge_db_path=KNOWLEDGE_PATH,
                original_filename=file.filename
            )
        except Exception:
            pass

        orig_filename = file.filename or "BOQ_Summary.xlsx"
        if not orig_filename.lower().endswith(".xlsx"):
            orig_filename += ".xlsx"
        gen_filename = orig_filename.replace(".xlsx", "_with_Summary.xlsx")

        background_tasks.add_task(_safe_remove, out_path)
        headers = {
            "X-Summary-Verified": "true" if verification.get("ok") else "false",
            "X-Summary-Warnings": json.dumps(verification.get("warnings", []), ensure_ascii=True),
        }

        return FileResponse(
            out_path,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename=gen_filename,
            headers=headers
        )
    except Exception as e:
        _safe_remove(out_path)
        raise HTTPException(status_code=500, detail=str(e))
    finally:
        _safe_remove(in_path)


@app.post("/api/generate-batch")
async def generate_batch(
    background_tasks: BackgroundTasks,
    files: List[UploadFile] = File(...),
    include_rainwater: bool = Form(True)
):
    """Generates Summary Sheets for multiple uploaded .xlsx files and returns a verified ZIP archive."""
    if not files:
        raise HTTPException(status_code=400, detail="No files uploaded")

    tmp_dir = tempfile.mkdtemp(prefix="boq_batch_")
    zip_fd, zip_path = tempfile.mkstemp(suffix="_Summary_Batch.zip")
    os.close(zip_fd)
    report_lines = ["=== VIEW SUMMARY STUDIO - BATCH VERIFICATION REPORT ===", ""]

    try:
        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
            for up in files:
                orig_name = up.filename or "BOQ.xlsx"
                if not orig_name.lower().endswith(".xlsx"):
                    continue
                in_p = os.path.join(tmp_dir, f"in_{orig_name}")
                out_name = orig_name.replace(".xlsx", "_with_Summary.xlsx")
                out_p = os.path.join(tmp_dir, out_name)
                with open(in_p, "wb") as f_out:
                    shutil.copyfileobj(up.file, f_out)

                try:
                    huliot_core.generate_summary_sheet(
                        in_p,
                        out_p,
                        original_filename=orig_name,
                        include_rainwater=include_rainwater
                    )
                    ver = huliot_core.verify_generated_summary(out_p)
                    try:
                        huliot_core.learn_from_generation(out_p, knowledge_db_path=KNOWLEDGE_PATH, original_filename=orig_name)
                    except Exception:
                        pass
                    zf.write(out_p, arcname=out_name)
                    status_str = "VERIFIED OK" if ver.get("ok") else f"WARNINGS: {ver.get('warnings')}"
                    report_lines.append(f"[OK] {out_name} -> {status_str}")
                except Exception as file_err:
                    report_lines.append(f"[ERROR] {orig_name} -> {file_err}")

            zf.writestr("Summary_Verification_Report.txt", "\n".join(report_lines))

        background_tasks.add_task(_safe_remove, zip_path)
        background_tasks.add_task(shutil.rmtree, tmp_dir, True)
        return FileResponse(
            zip_path,
            media_type="application/zip",
            filename="BOQ_Summaries_Batch.zip"
        )
    except Exception as e:
        _safe_remove(zip_path)
        shutil.rmtree(tmp_dir, ignore_errors=True)
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/api/auto-agent")
async def auto_agent_run(
    background_tasks: BackgroundTasks,
    file: UploadFile = File(...),
    project_name: str = Form(None),
    num_toilets: int = Form(None),
    num_kitchens: int = Form(None),
    discount_p: float = Form(None),
    discount_f: float = Form(None),
    discount_m: float = Form(None),
    discount_l: float = Form(None),
    discount_c: float = Form(None),
    include_rainwater: bool = Form(True)
):
    """
    Auto-Agent Mode:
      1. Inspects workbook & runs compute_format_confidence via analyze_workbook.
      2. If confidence >= 70%: skips training and directly generates Summary.
         If confidence < 70%: trains format first, then generates Summary.
      3. Returns generated .xlsx plus structured X-Agent-Log header.
    """
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_in:
        shutil.copyfileobj(file.file, tmp_in)
        in_path = tmp_in.name

    out_path = in_path.replace(".xlsx", "_agent_summary.xlsx")

    try:
        orig_name = file.filename or "BOQ.xlsx"
        t_start = time.perf_counter()

        # Step 1: Confidence Check
        t1 = time.perf_counter()
        meta = huliot_core.analyze_workbook(
            in_path,
            knowledge_db_path=KNOWLEDGE_PATH,
            original_filename=orig_name
        )
        t1_s = round(time.perf_counter() - t1, 2)
        conf = int(meta.get("format_confidence", 0))
        fmt_label = meta.get("detected_format_label", "Standard Format")
        auto_rainwater = bool(meta.get("has_rainwater"))

        saw_msg = f"Format: {fmt_label}, Confidence: {conf}% (Toilets: {meta.get('estimated_toilets')}, Kitchens: {meta.get('estimated_kitchens')}, Rainwater: {'Yes' if auto_rainwater else 'No'})"
        did_train = False

        # Step 2: Decide Train / No Train
        t2 = time.perf_counter()
        if conf >= 70:
            decided_msg = f"Confidence {conf}% (>= 70%) था, इसलिए Train नहीं किया — सीधे Summary बनाने का निर्णय लिया।"
            summary_line = f"Confidence {conf}% था, इसलिए सीधे Summary बनाया"
            step2_label = "No Train (Known Format)"
        else:
            decided_msg = f"Confidence {conf}% (< 70%) था, इसलिए पहले Format Train करने और फिर Summary बनाने का निर्णय लिया।"
            huliot_core.train_from_workbook(
                in_path,
                knowledge_db_path=KNOWLEDGE_PATH,
                original_filename=orig_name
            )
            did_train = True
            summary_line = f"Confidence {conf}% था, इसलिए पहले Train किया, फिर Summary बनाया"
            step2_label = "Trained New Format"
        t2_s = round(time.perf_counter() - t2, 2)

        discounts = None
        if all(v is not None for v in [discount_p, discount_f, discount_m, discount_l, discount_c]):
            discounts = {
                "P": discount_p,
                "F": discount_f,
                "M": discount_m,
                "L": discount_l,
                "C": discount_c,
            }

        # Step 3: Summary Generate
        t3 = time.perf_counter()
        huliot_core.generate_summary_sheet(
            in_path,
            out_path,
            project_name=project_name or None,
            num_toilets=num_toilets,
            num_kitchens=num_kitchens,
            discounts=discounts,
            original_filename=orig_name,
            include_rainwater=auto_rainwater
        )
        t3_s = round(time.perf_counter() - t3, 2)

        # Step 4: Verify
        t4 = time.perf_counter()
        verification = huliot_core.verify_generated_summary(out_path)
        t4_s = round(time.perf_counter() - t4, 2)
        total_s = round(time.perf_counter() - t_start, 2)

        if did_train:
            did_msg = f"पहले Format को Knowledge Base में Train किया, फिर {os.path.basename(orig_name)} की Summary Sheet बनाई और Verify की ({verification.get('checked_cells', 0)} cells OK)।"
        else:
            did_msg = f"सीधे {os.path.basename(orig_name)} की Summary Sheet बनाई और Verify की ({verification.get('checked_cells', 0)} cells OK)।"

        ver_ok = bool(verification.get("ok"))
        ver_cells = verification.get("checked_cells", 0)
        steps = [
            {
                "step": "Step 1: Confidence Check",
                "time_s": t1_s,
                "status": "done",
                "detail": f"{fmt_label} • Confidence {conf}%"
            },
            {
                "step": "Step 2: Decide Train / No Train",
                "time_s": t2_s,
                "status": "trained" if did_train else "skipped",
                "detail": step2_label
            },
            {
                "step": "Step 3: Summary Generate",
                "time_s": t3_s,
                "status": "done",
                "detail": f"Live Excel Summary generated{' (+ Rainwater Auto-Included)' if auto_rainwater else ''}"
            },
            {
                "step": "Step 4: Verify",
                "time_s": t4_s,
                "status": "done" if ver_ok else "warning",
                "detail": f"{ver_cells} cells checked • {'0 errors' if ver_ok else 'warnings found'}"
            },
        ]

        agent_log = {
            "confidence": conf,
            "format_label": fmt_label,
            "trained": did_train,
            "saw": saw_msg,
            "decided": decided_msg,
            "did": did_msg,
            "summary_line": summary_line,
            "verified": ver_ok,
            "warnings": verification.get("warnings", []),
            "total_time_s": total_s,
            "steps": steps,
        }

        if not orig_name.lower().endswith(".xlsx"):
            orig_name += ".xlsx"
        gen_filename = orig_name.replace(".xlsx", "_with_Summary.xlsx")

        background_tasks.add_task(_safe_remove, out_path)
        headers = {
            "X-Summary-Verified": "true" if verification.get("ok") else "false",
            "X-Summary-Warnings": json.dumps(verification.get("warnings", []), ensure_ascii=True),
            "X-Agent-Log": json.dumps(agent_log, ensure_ascii=True),
        }

        return FileResponse(
            out_path,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename=gen_filename,
            headers=headers
        )
    except Exception as e:
        _safe_remove(out_path)
        raise HTTPException(status_code=500, detail=str(e))
    finally:
        _safe_remove(in_path)


@app.get("/api/scan-folder")
def scan_workspace_folder():
    """Scans data/raw_boq folder for existing training files and trains knowledge base."""
    if not os.path.exists(RAW_BOQ_DIR):
        return {"status": "success", "scanned_files": 0, "trained_files": []}
    files = [f for f in os.listdir(RAW_BOQ_DIR) if f.endswith(".xlsx") and not f.startswith("~")]
    trained = []
    for f in files:
        full_path = os.path.join(RAW_BOQ_DIR, f)
        res = huliot_core.train_from_workbook(full_path, knowledge_db_path=KNOWLEDGE_PATH)
        if res.get("status") == "success":
            trained.append(f)

    return {"status": "success", "scanned_files": len(files), "trained_files": trained}


@app.post("/api/admin/login")
async def admin_login(request: Request, password: str = Form(...)):
    """Validates Admin password with 5-attempt / 15-minute lockout protection."""
    _verify_admin(password, request)
    return {"status": "ok", "role": "admin"}


@app.get("/api/admin/dashboard")
def admin_dashboard(request: Request, x_admin_password: Optional[str] = Header(default=None)):
    """Returns Admin Dashboard metrics from knowledge.json."""
    _verify_admin(x_admin_password, request)
    db = huliot_core.load_knowledge_db(KNOWLEDGE_PATH)
    projects = db.get("learned_projects", [])
    today_str = datetime.now().strftime("%Y-%m-%d")

    trained_today = sum(1 for p in projects if str(p.get("trained_at", "")).startswith(today_str))
    dated = [p for p in projects if p.get("trained_at")]
    dated_sorted = sorted(dated, key=lambda x: x.get("trained_at", ""), reverse=True)
    last_trained_at = dated_sorted[0].get("trained_at") if dated_sorted else (
        "Legacy (No timestamp)" if projects else "Never"
    )

    # Last 10 trained files (newest first if timestamped, else tail of list reversed)
    undated = [p for p in reversed(projects) if not p.get("trained_at")]
    recent_pool = dated_sorted + undated
    all_formats = [
        {
            "filename": p.get("filename", ""),
            "project_title": p.get("project_title", ""),
            "training_type": p.get("training_type", "full_summary"),
            "archetype": p.get("archetype", "standard"),
            "trained_at": p.get("trained_at", "Legacy"),
        }
        for p in recent_pool
    ]

    return {
        "total_formats": len(projects),
        "trained_today": trained_today,
        "last_trained_at": last_trained_at,
        "recent_10": all_formats[:10],
        "all_formats": all_formats,
    }


@app.get("/api/admin/download-knowledge")
def admin_download_knowledge(request: Request, x_admin_password: Optional[str] = Header(default=None)):
    """Allows Admin to download a backup of knowledge.json."""
    _verify_admin(x_admin_password, request)
    if not os.path.exists(KNOWLEDGE_PATH):
        raise HTTPException(status_code=404, detail="knowledge.json not found")
    return FileResponse(
        KNOWLEDGE_PATH,
        media_type="application/json",
        filename=f"knowledge_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json",
    )


# Mount Flutter Web if built
if os.path.exists(WEB_BUILD_DIR):
    app.mount("/", StaticFiles(directory=WEB_BUILD_DIR, html=True), name="static_flutter")


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="127.0.0.1", port=8000)
