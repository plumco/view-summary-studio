import os
import sys

# Ensure backend directory is on sys.path when invoked from any working directory
BACKEND_DIR = os.path.dirname(os.path.abspath(__file__))
if BACKEND_DIR not in sys.path:
    sys.path.insert(0, BACKEND_DIR)

from mcp.server.mcpserver import MCPServer
import huliot_core

mcp = MCPServer("VIEW SUMMARY STUDIO")


@mcp.tool()
def train_format(file_path: str) -> dict:
    """
    Takes an Excel (.xlsx) BOQ file path, extracts its Summary or Raw BOQ layout format,
    and saves the learned format into data/trained_formats/knowledge.json.
    """
    if not file_path or not os.path.exists(file_path):
        return {"status": "error", "message": f"File not found: {file_path}"}
    if not file_path.lower().endswith(".xlsx"):
        return {"status": "error", "message": "Only .xlsx Excel files are supported."}

    return huliot_core.train_from_workbook(
        file_path,
        knowledge_db_path=huliot_core.DEFAULT_KNOWLEDGE_PATH,
        original_filename=os.path.basename(file_path),
    )


@mcp.tool()
def generate_summary(file_path: str, output_path: str = "", include_rainwater: bool = True) -> dict:
    """
    Takes an Excel (.xlsx) BOQ file path, generates a dynamic live-formula Summary Sheet,
    verifies formulas for zero errors, auto-learns the format, and returns the output Excel path.
    """
    if not file_path or not os.path.exists(file_path):
        return {"status": "error", "message": f"File not found: {file_path}"}
    if not file_path.lower().endswith(".xlsx"):
        return {"status": "error", "message": "Only .xlsx Excel files are supported."}

    out_path = output_path.strip() if output_path else ""
    if not out_path:
        base, ext = os.path.splitext(file_path)
        out_path = f"{base}_with_Summary{ext}"

    huliot_core.generate_summary_sheet(
        file_path,
        out_path,
        original_filename=os.path.basename(file_path),
        include_rainwater=include_rainwater,
    )
    verification = huliot_core.verify_generated_summary(out_path)
    try:
        huliot_core.learn_from_generation(
            out_path,
            knowledge_db_path=huliot_core.DEFAULT_KNOWLEDGE_PATH,
            original_filename=os.path.basename(file_path),
        )
    except Exception:
        pass

    return {
        "status": "success",
        "output_path": os.path.abspath(out_path),
        "verification": verification,
    }


if __name__ == "__main__":
    mcp.run()
