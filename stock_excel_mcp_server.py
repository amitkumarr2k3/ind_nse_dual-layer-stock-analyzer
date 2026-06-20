from pathlib import Path

from mcp.server.fastmcp import FastMCP

from Stock_Agent import (
    answer_stock_question_from_report,
    build_stock_report_snapshot,
    list_report_stocks,
    load_config,
)

APP_ROOT = Path(__file__).resolve().parent
DEFAULT_CONFIG = APP_ROOT / "config.json"
mcp = FastMCP("stock-analysis-report")


def _load_runtime_config(config_file=None):
    config_path = Path(config_file) if config_file else DEFAULT_CONFIG
    return load_config(str(config_path))


@mcp.tool()
def list_available_stocks(excel_file: str = "", config_file: str = "", limit: int = 50) -> dict:
    """List stocks available in the generated Excel Analysis sheet."""
    config = _load_runtime_config(config_file or None)
    workbook = excel_file or config.get("export_file", "AI_STOCK_ANALYSIS.xlsx")
    stocks = list_report_stocks(workbook, limit=limit)
    return {
        "excel_file": workbook,
        "count": len(stocks),
        "stocks": stocks,
    }


@mcp.tool()
def get_stock_snapshot(stock: str, excel_file: str = "", config_file: str = "") -> dict:
    """Return the exact Excel-backed data snapshot for one stock."""
    config = _load_runtime_config(config_file or None)
    snapshot = build_stock_report_snapshot(stock, config, excel_file or None)
    return {
        "stock": snapshot["stock"],
        "excel_file": snapshot["excel_file"],
        "report_fields": snapshot["evidence"],
    }


@mcp.tool()
def answer_stock_question(stock: str, question: str, excel_file: str = "", config_file: str = "") -> dict:
    """Answer a stock question using only the generated Excel report and its stored AI recommendation."""
    config = _load_runtime_config(config_file or None)
    return answer_stock_question_from_report(stock, question, config, excel_file or None)


if __name__ == "__main__":
    mcp.run()
