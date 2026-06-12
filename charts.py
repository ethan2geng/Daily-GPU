#!/usr/bin/env python3
"""
Chart generation, history tracking, and email for GPU price trends.
Reads history.csv → builds GPU_Price_Trends.xlsx → emails HTML table + attachment.
"""

import io
import json
import smtplib
import tempfile
from datetime import datetime, date
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
import numpy as np
import pandas as pd
import xlsxwriter
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

SCRIPT_DIR  = Path(__file__).parent
SAVE_DIR    = SCRIPT_DIR / "output"
SAVE_DIR.mkdir(parents=True, exist_ok=True)
HISTORY_CSV = SCRIPT_DIR / "history.csv"
CHARTS_FILE = SAVE_DIR / "GPU_Price_Trends.xlsx"

START_DATE = date(2026, 6, 9)   # day 0

GPU_MODELS = [
    "NVIDIA B300 SXM6",
    "NVIDIA B200 NVL",
    "NVIDIA B200 SXM",
    "NVIDIA H200 NVL",
    "NVIDIA H200 SXM",
    "NVIDIA GH200 Grace Hopper",
    "NVIDIA H100 SXM5",
    "NVIDIA H100 NVL",
    "AMD Instinct MI325X",
    "AMD Instinct MI300X",
    "NVIDIA L40S",
    "Intel Gaudi 2",
]

RETAIL_PROVIDERS = {
    "Vast.ai", "RunPod", "TensorDock", "Salad", "JarvisLabs", "LeaderGPU",
}

ENT_COLOR    = "#1F4E79"
RETAIL_COLOR = "#C55A11"
GRID_COLOR   = "#CCCCCC"
AXIS_COLOR   = "#000000"


def classify(provider: str) -> str:
    return "retail" if provider in RETAIL_PROVIDERS else "enterprise"


def _days_since_start(date_str: str) -> int:
    d = date.fromisoformat(date_str)
    return (d - START_DATE).days


# ---------------------------------------------------------------------------
# History
# ---------------------------------------------------------------------------

def append_history(offers: list[dict]) -> pd.DataFrame:
    today = datetime.now().strftime("%Y-%m-%d")
    df = pd.DataFrame(offers)
    df["tier"] = df["provider"].apply(classify)

    rows = []
    for model in GPU_MODELS:
        sub = df[df["gpu_model"] == model]
        ent = sub[sub["tier"] == "enterprise"]["price"]
        ret = sub[sub["tier"] == "retail"]["price"]
        rows.append({
            "date":           today,
            "gpu_model":      model,
            "enterprise_min": round(ent.min(), 4) if not ent.empty else None,
            "retail_min":     round(ret.min(), 4) if not ret.empty else None,
        })

    new_df = pd.DataFrame(rows)
    if HISTORY_CSV.exists():
        hist = pd.read_csv(HISTORY_CSV)
        hist = hist[hist["date"] != today]
        hist = pd.concat([hist, new_df], ignore_index=True)
    else:
        hist = new_df

    hist.to_csv(HISTORY_CSV, index=False)
    print(f"History updated → {HISTORY_CSV}  ({len(hist)} total rows)", flush=True)
    return hist


def backfill_from_excels():
    files = sorted(SAVE_DIR.glob("gpu_prices_????-??-??.xlsx"))
    if not files:
        print("No existing Excel files to backfill from.")
        return

    all_rows = []
    for f in files:
        date_str = f.stem.replace("gpu_prices_", "")
        try:
            pivot = pd.read_excel(f, sheet_name="Prices", index_col=0)
        except Exception as e:
            print(f"  Skipping {f.name}: {e}")
            continue

        for model in GPU_MODELS:
            if model not in pivot.index:
                all_rows.append({"date": date_str, "gpu_model": model,
                                 "enterprise_min": None, "retail_min": None})
                continue
            row = pivot.loc[model]
            ent_prices, ret_prices = [], []
            for provider, price in row.items():
                if pd.isna(price):
                    continue
                (ret_prices if classify(provider) == "retail" else ent_prices).append(price)
            all_rows.append({
                "date":           date_str,
                "gpu_model":      model,
                "enterprise_min": round(min(ent_prices), 4) if ent_prices else None,
                "retail_min":     round(min(ret_prices), 4) if ret_prices else None,
            })
        print(f"  Backfilled {f.name}")

    hist = pd.DataFrame(all_rows)
    hist.to_csv(HISTORY_CSV, index=False)
    print(f"Backfill done → {HISTORY_CSV}  ({len(hist)} rows)")
    return hist


# ---------------------------------------------------------------------------
# Chart image generation (matplotlib)
# ---------------------------------------------------------------------------

def _make_chart_image(model: str, sub: pd.DataFrame, out_path: Path):
    """Save a PNG chart for one model to out_path."""
    sub = sub.sort_values("date").reset_index(drop=True)
    sub["day"] = sub["date"].apply(_days_since_start)

    fig, ax = plt.subplots(figsize=(9, 5))
    fig.patch.set_facecolor("white")
    ax.set_facecolor("white")

    has_any = False

    for col, label, color in [
        ("enterprise_min", "Enterprise Min", ENT_COLOR),
        ("retail_min",     "Retail Min",     RETAIL_COLOR),
    ]:
        mask = sub[col].notna()
        if mask.any():
            xs = sub.loc[mask, "day"].values
            ys = sub.loc[mask, col].values
            ax.plot(xs, ys,
                    color=color, linewidth=2,
                    marker="o", markersize=6, markerfacecolor=color,
                    markeredgecolor="white", markeredgewidth=1.2,
                    label=label, zorder=3)
            has_any = True

    if not has_any:
        ax.text(0.5, 0.5, "No data yet", transform=ax.transAxes,
                ha="center", va="center", fontsize=13, color="#999999")
        ax.axis("off")
        plt.tight_layout()
        fig.savefig(str(out_path), format="png", dpi=130, bbox_inches="tight")
        plt.close(fig)
        return

    # ── Axes: only bottom + left, both black ─────────────────────────────────
    for spine in ax.spines.values():
        spine.set_visible(False)
    ax.spines["bottom"].set_visible(True)
    ax.spines["left"].set_visible(True)
    ax.spines["bottom"].set_color(AXIS_COLOR)
    ax.spines["left"].set_color(AXIS_COLOR)
    ax.spines["bottom"].set_linewidth(1.5)
    ax.spines["left"].set_linewidth(1.5)
    ax.tick_params(colors=AXIS_COLOR, length=4, width=1)

    # ── Light gray grid lines behind the data ────────────────────────────────
    ax.yaxis.grid(True, color=GRID_COLOR, linewidth=0.8, linestyle="-", zorder=0)
    ax.xaxis.grid(True, color=GRID_COLOR, linewidth=0.8, linestyle="-", zorder=0)
    ax.set_axisbelow(True)

    # ── X-axis: integer days ──────────────────────────────────────────────────
    all_days = sub["day"].values
    ax.set_xlim(all_days.min() - 0.3, all_days.max() + 0.3)
    ax.xaxis.set_major_locator(ticker.MaxNLocator(integer=True))
    ax.set_xlabel("Days since June 9", fontsize=11, color=AXIS_COLOR, labelpad=6)
    ax.tick_params(axis="x", labelsize=10)

    # ── Y-axis: sensible $ increments ────────────────────────────────────────
    all_vals = []
    for col in ("enterprise_min", "retail_min"):
        all_vals.extend(sub[col].dropna().tolist())

    if all_vals:
        lo, hi = min(all_vals), max(all_vals)
        span = hi - lo if hi > lo else 1.0
        raw_step = span / 5
        magnitude = 10 ** np.floor(np.log10(raw_step))
        for nice in [0.25, 0.5, 1, 2, 2.5, 5, 10]:
            step = nice * magnitude
            if span / step <= 8:
                break
        ax.yaxis.set_major_locator(ticker.MultipleLocator(step))
        padding = step * 0.6
        ax.set_ylim(max(0, lo - padding), hi + padding)

    ax.yaxis.set_major_formatter(ticker.FuncFormatter(lambda v, _: f"${v:.2f}"))
    ax.set_ylabel("Price ($/GPU/hr)", fontsize=11, color=AXIS_COLOR, labelpad=6)
    ax.tick_params(axis="y", labelsize=10)

    ax.set_title(model, fontsize=12, fontweight="bold", color="#111111", pad=10)
    ax.legend(frameon=False, fontsize=10, loc="upper left")

    plt.tight_layout(pad=1.5)
    fig.savefig(str(out_path), format="png", dpi=130, bbox_inches="tight")
    plt.close(fig)


# ---------------------------------------------------------------------------
# Excel workbook
# ---------------------------------------------------------------------------

def generate_charts_excel(history_df: pd.DataFrame) -> Path:
    SAVE_DIR.mkdir(exist_ok=True)
    # Delete old file first so xlsxwriter starts clean
    if CHARTS_FILE.exists():
        CHARTS_FILE.unlink()

    wb  = xlsxwriter.Workbook(str(CHARTS_FILE))

    hdr_fmt  = wb.add_format({"bold": True, "font_color": "#FFFFFF",
                               "bg_color": "#1F4E79", "align": "center",
                               "valign": "vcenter", "border": 1,
                               "border_color": "#CCCCCC", "font_size": 10})
    price_fmt = wb.add_format({"num_format": '"$"#,##0.00', "align": "center",
                                "border": 1, "border_color": "#CCCCCC", "font_size": 10})
    na_fmt   = wb.add_format({"font_color": "#999999", "italic": True,
                               "align": "center", "border": 1,
                               "border_color": "#CCCCCC", "font_size": 10})
    date_fmt = wb.add_format({"align": "center", "border": 1,
                               "border_color": "#CCCCCC", "font_size": 10})

    chart_tmp_dir = SCRIPT_DIR / "chart_tmp"
    chart_tmp_dir.mkdir(exist_ok=True)
    tmp_paths: list[Path] = []

    for idx, model in enumerate(GPU_MODELS):
        sub = history_df[history_df["gpu_model"] == model].copy()
        sub = sub.sort_values("date").reset_index(drop=True)

        short = (model
                 .replace("NVIDIA ", "")
                 .replace("AMD ", "")
                 .replace("Intel ", ""))[:31]
        ws = wb.add_worksheet(short)

        # ── Chart PNG ────────────────────────────────────────────────────
        safe_name = model.replace(" ", "_").replace("/", "-")
        tmp = chart_tmp_dir / f"{safe_name}.png"
        tmp_paths.append(tmp)
        _make_chart_image(model, sub, tmp)
        print(f"  [{idx+1:02d}] {short}: {tmp.stat().st_size} bytes", flush=True)

        ws.insert_image("E1", str(tmp), {"width": 620, "height": 360})

        # ── Data table (columns A–C) ──────────────────────────────────────
        ws.set_column("A:A", 14)
        ws.set_column("B:B", 18)
        ws.set_column("C:C", 16)
        ws.set_row(0, 18)

        ws.write(0, 0, "Date",           hdr_fmt)
        ws.write(0, 1, "Enterprise Min", hdr_fmt)
        ws.write(0, 2, "Retail Min",     hdr_fmt)

        for i, r in sub.iterrows():
            row = i + 1
            ent = float(r["enterprise_min"]) if pd.notna(r.get("enterprise_min")) else None
            ret = float(r["retail_min"])     if pd.notna(r.get("retail_min"))     else None

            ws.write(row, 0, r["date"], date_fmt)
            ws.write(row, 1, ent if ent is not None else "N/A",
                     price_fmt if ent is not None else na_fmt)
            ws.write(row, 2, ret if ret is not None else "N/A",
                     price_fmt if ret is not None else na_fmt)
            ws.set_row(row, 16)

    wb.close()

    for tmp in tmp_paths:
        tmp.unlink(missing_ok=True)
    try:
        chart_tmp_dir.rmdir()
    except Exception:
        pass

    print(f"Charts saved → {CHARTS_FILE}", flush=True)
    return CHARTS_FILE


# ---------------------------------------------------------------------------
# Email
# ---------------------------------------------------------------------------

def _fmt(val) -> str:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return "N/A"
    return f"${val:.2f}"


def _build_html_table(history_df: pd.DataFrame, today: str) -> str:
    latest = (history_df.sort_values("date")
              .groupby("gpu_model").last().reset_index()
              .set_index("gpu_model").reindex(GPU_MODELS).reset_index())

    rows_html = ""
    for i, row in latest.iterrows():
        bg  = "#F7F9FC" if i % 2 == 0 else "#FFFFFF"
        ent = _fmt(row.get("enterprise_min"))
        ret = _fmt(row.get("retail_min"))
        ec  = ENT_COLOR    if ent != "N/A" else "#AAAAAA"
        rc  = RETAIL_COLOR if ret != "N/A" else "#AAAAAA"
        rows_html += f"""
        <tr style="background:{bg};">
          <td style="padding:9px 16px;border:1px solid #DDD;">{row['gpu_model']}</td>
          <td style="padding:9px 16px;border:1px solid #DDD;text-align:center;color:{ec};font-weight:600;">{ent}</td>
          <td style="padding:9px 16px;border:1px solid #DDD;text-align:center;color:{rc};font-weight:600;">{ret}</td>
        </tr>"""

    return f"""<!DOCTYPE html><html><body style="font-family:Arial,sans-serif;color:#222;padding:20px;">
  <p style="font-size:15px;margin-bottom:4px;">GPU price summary for <strong>{today}</strong>.</p>
  <p style="font-size:12px;color:#888;margin-top:0;margin-bottom:14px;">
    Enterprise = cloud/DC providers &nbsp;·&nbsp; Retail = marketplace platforms (Vast.ai, RunPod, etc.)
  </p>
  <table style="border-collapse:collapse;font-size:14px;min-width:480px;">
    <thead>
      <tr style="background:#1F4E79;color:#FFF;">
        <th style="padding:10px 16px;border:1px solid #1A3F66;text-align:left;">GPU Model</th>
        <th style="padding:10px 16px;border:1px solid #1A3F66;text-align:center;">Enterprise Min</th>
        <th style="padding:10px 16px;border:1px solid #1A3F66;text-align:center;">Retail Min</th>
      </tr>
    </thead>
    <tbody>{rows_html}
    </tbody>
  </table>
  <p style="font-size:12px;color:#999;margin-top:14px;">GPU_Price_Trends.xlsx and today's full price table attached.</p>
</body></html>"""


def send_email(charts_path: Path, daily_path: Path | None = None,
               history_df: pd.DataFrame | None = None):
    import os
    user  = os.getenv("SMTP_USER")
    pw    = os.getenv("SMTP_PASSWORD")
    to    = os.getenv("TO_EMAIL")

    if not user or not pw or not to:
        config_file = SCRIPT_DIR / "email_config.json"
        if config_file.exists():
            cfg = json.loads(config_file.read_text())
            user = cfg.get("smtp_user")
            pw = cfg.get("smtp_password")
            to = cfg.get("to_email")
        else:
            print("Email config not set (SMTP_USER, SMTP_PASSWORD, TO_EMAIL env vars or email_config.json) — skipping email.", flush=True)
            return

    if not user or not pw or not to:
        print("Email config incomplete — skipping email.", flush=True)
        return
    today = datetime.now().strftime("%Y-%m-%d")

    msg = MIMEMultipart("mixed")
    msg["From"]    = user
    msg["To"]      = to
    msg["Subject"] = f"GPU Price Update — {today}"

    html = _build_html_table(history_df, today) if history_df is not None else \
           f"<p>GPU price update for {today}. See attachments.</p>"
    msg.attach(MIMEMultipart("alternative"))
    msg.get_payload()[-1].attach(MIMEText(html, "html"))

    for path in [charts_path, daily_path]:
        if path and path.exists():
            with open(path, "rb") as f:
                part = MIMEBase("application", "octet-stream")
                part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header("Content-Disposition", f'attachment; filename="{path.name}"')
            msg.attach(part)

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as srv:
        srv.login(user, pw)
        srv.sendmail(user, to, msg.as_string())
    print(f"Email sent → {to}", flush=True)


# ---------------------------------------------------------------------------
# Standalone
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    if not HISTORY_CSV.exists():
        print("No history.csv — backfilling from existing Excel files…")
        backfill_from_excels()
    hist = pd.read_csv(HISTORY_CSV)
    generate_charts_excel(hist)
    print("Done.")
