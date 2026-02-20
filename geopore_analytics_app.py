# geopore_analytics_app_v1_1_0.py
# GeoPore Analytics — MICP (Multi‑sample Dashboard) v1.1.0
#
# Features requested:
# - UI scaled to ~75% (viewport initial-scale + optional CSS zoom fallback)
# - Dark theme (dbc.themes.CYBORG) + custom palette:
#   App background:   #1E1E2F
#   Panel background: #27293D
#   Accent:           #00D2FF
#   Text:             #E0E0E0
#   Curves:           #FF007C, #00F2C3, #FD9A00
# - Improved layout and button styling
# - Multi-file upload (multiple Excel files for same well)
# - Dashboard header shows the currently selected document name
# - Real PSD: dV/dlog(r) (normalized)
# - Winland crossplot (k vs r35) with rock-type classification
# - Thomeer fit with plot + export parameters per sample
# - PDF report with plots (Pc‑Sw, PSD, Thomeer, SHF) + metrics tables
#
# Run:
#   pip install dash dash-ag-grid dash-bootstrap-components pandas numpy scipy openpyxl xlrd reportlab matplotlib
#   python geopore_analytics_app_v1_1_0.py
#
# Notes:
# - This is an MVP designed to be robust with Micromeritics/AutoPore style Excel exports.
# - The UI is dark by default; if you prefer a different bootstrap theme, switch CYBORG -> DARKLY.

from __future__ import annotations

import base64
import io
import json
import math
import os
import re
import datetime as _dt
import hashlib
import time
from dataclasses import dataclass
from typing import Any, Dict, List, Optional, Tuple

import numpy as np
import pandas as pd

from dash import Dash, Input, Output, State, dcc, html, no_update, callback_context
from dash.exceptions import PreventUpdate
import dash_ag_grid as dag
import dash_bootstrap_components as dbc

# Optional scientific stack
try:
    from scipy.optimize import curve_fit
except Exception:  # pragma: no cover
    curve_fit = None  # type: ignore

# PDF/report
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors as rl_colors

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt


VERSION = '1.8.25.6'
# Base dashboard (plot/grid) height in px — requested 1.75× increase over prior 640px
DASHBOARD_HEIGHT_PX = int(640 * 1.75)  # = 1120px

# ---------------------------
# Local workspace (disk persistence)
# ---------------------------
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(SCRIPT_DIR, "geopore_config.json")
DEFAULT_WORKSPACE_DIR = os.path.join(SCRIPT_DIR, "geopore_workspace")


def _safe_read_json(path: str, default: Any):
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return default


def _safe_write_json(path: str, data: Any) -> None:
    try:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2)
    except Exception:
        # Best-effort; do not crash app if disk is read-only, etc.
        pass


def get_workspace_dir() -> str:
    cfg = _safe_read_json(CONFIG_PATH, {})
    ws = None
    if isinstance(cfg, dict):
        ws = cfg.get("workspace_dir")
    if not ws:
        ws = DEFAULT_WORKSPACE_DIR
    return ws


def set_workspace_dir(path: str) -> str:
    # Persist workspace path to config file (best-effort)
    cfg = _safe_read_json(CONFIG_PATH, {})
    if not isinstance(cfg, dict):
        cfg = {}
    cfg["workspace_dir"] = path
    _safe_write_json(CONFIG_PATH, cfg)
    return path


def ensure_workspace(path: str) -> Dict[str, str]:
    """Create the workspace folders if missing and return useful subpaths."""
    ws = path or DEFAULT_WORKSPACE_DIR
    projects = os.path.join(ws, "projects")
    uploads = os.path.join(ws, "uploads")
    os.makedirs(projects, exist_ok=True)
    os.makedirs(uploads, exist_ok=True)
    return {"workspace": ws, "projects": projects, "uploads": uploads}


def list_project_files(projects_dir: str) -> List[Dict[str, str]]:
    """Return dropdown options for saved project JSON files."""
    opts: List[Dict[str, str]] = []
    try:
        if not projects_dir or not os.path.isdir(projects_dir):
            return []
        files = [f for f in os.listdir(projects_dir) if f.lower().endswith(".json")]

        def _sort_key(fn: str):
            full = os.path.join(projects_dir, fn)
            mtime = os.path.getmtime(full) if os.path.exists(full) else 0
            # autosave first
            is_auto = 0 if fn.lower().startswith("autosave") else 1
            return (is_auto, -mtime)

        files.sort(key=_sort_key)
        for fn in files:
            full = os.path.join(projects_dir, fn)
            label = fn
            try:
                mtime = os.path.getmtime(full)
                ts = _dt.datetime.fromtimestamp(mtime).strftime("%Y-%m-%d %H:%M")
                label = f"{fn}  ({ts})"
            except Exception:
                pass
            opts.append({"label": label, "value": full})
    except Exception:
        return []
    return opts


# Initialize workspace
WORKSPACE_DIR = get_workspace_dir()
_WS_PATHS = ensure_workspace(WORKSPACE_DIR)
PROJECTS_DIR = _WS_PATHS["projects"]
UPLOADS_DIR = _WS_PATHS["uploads"]
AUTOSAVE_PATH = os.path.join(PROJECTS_DIR, "autosave_project.json")


# ---------------------------
# Theme / palette
# ---------------------------
COLORS = {
    "app_bg": "#1E1E2F",
    "panel_bg": "#27293D",
    "accent": "#00D2FF",
    "text": "#E0E0E0",
    "curve1": "#FF007C",
    "curve2": "#00F2C3",
    "curve3": "#FD9A00",
}

# ---------------------------
# Plot background themes (Graph area)
# ---------------------------
PLOT_THEMES: Dict[str, Dict[str, str]] = {
    "dark": {
        "template": "plotly_dark",
        "paper_bgcolor": COLORS["panel_bg"],
        "plot_bgcolor": COLORS["app_bg"],
        "font_color": COLORS["text"],
        "grid_color": "rgba(255,255,255,0.08)",
        "zeroline_color": "rgba(255,255,255,0.18)",
    },
    "light": {
        "template": "plotly_white",
        "paper_bgcolor": "#F3F5F9",
        "plot_bgcolor": "#FFFFFF",
        "font_color": "#111111",
        "grid_color": "rgba(0,0,0,0.10)",
        "zeroline_color": "rgba(0,0,0,0.20)",
    },
    "white": {
        "template": "plotly_white",
        "paper_bgcolor": "#FFFFFF",
        "plot_bgcolor": "#FFFFFF",
        "font_color": "#111111",
        "grid_color": "rgba(0,0,0,0.08)",
        "zeroline_color": "rgba(0,0,0,0.18)",
    },
}


def apply_plot_theme(fig: go.Figure, theme: str = "dark") -> go.Figure:
    """Apply a plot background theme to the main Plotly figure."""
    theme_key = (theme or "dark").strip().lower()
    t = PLOT_THEMES.get(theme_key, PLOT_THEMES["dark"])
    fig.update_layout(
        template=t["template"],
        paper_bgcolor=t["paper_bgcolor"],
        plot_bgcolor=t["plot_bgcolor"],
        font=dict(color=t["font_color"]),
        legend=dict(font=dict(color=t["font_color"])),
        hoverlabel=dict(font=dict(color=t["font_color"])),
        hovermode="closest",
    )
    fig.update_xaxes(
        gridcolor=t["grid_color"],
        zerolinecolor=t["zeroline_color"],
        color=t["font_color"],
    )
    fig.update_yaxes(
        gridcolor=t["grid_color"],
        zerolinecolor=t["zeroline_color"],
        color=t["font_color"],
    )
    # 3D scenes (for PNM network, etc.)
    try:
        fig.update_scenes(
            bgcolor=t["plot_bgcolor"],
            xaxis=dict(
                backgroundcolor=t["plot_bgcolor"],
                gridcolor=t["grid_color"],
                zerolinecolor=t["zeroline_color"],
                color=t["font_color"],
            ),
            yaxis=dict(
                backgroundcolor=t["plot_bgcolor"],
                gridcolor=t["grid_color"],
                zerolinecolor=t["zeroline_color"],
                color=t["font_color"],
            ),
            zaxis=dict(
                backgroundcolor=t["plot_bgcolor"],
                gridcolor=t["grid_color"],
                zerolinecolor=t["zeroline_color"],
                color=t["font_color"],
            ),
        )
    except Exception:
        pass
    return fig


# -----------------------------------------------------------------------------
# Robust axis auto-scaling
def _empty_fig(message: str = "No data.", height: int = 260):
    """Return a themed empty Plotly figure with a centered message."""
    try:
        import plotly.graph_objects as go  # local import (keeps startup light)
    except Exception:
        # Fallback: avoid crashing if Plotly is not available for some reason
        return {}

    fig = go.Figure()
    fig.add_annotation(
        text=message,
        x=0.5,
        y=0.5,
        xref="paper",
        yref="paper",
        showarrow=False,
        font=dict(size=12, color=COLORS.get("text", "white")),
    )

    fig.update_xaxes(visible=False, showgrid=False, zeroline=False)
    fig.update_yaxes(visible=False, showgrid=False, zeroline=False)

    fig.update_layout(
        height=height,
        margin=dict(l=20, r=20, t=25, b=20),
    )
    try:
        fig = apply_plot_theme(fig)
    except Exception:
        pass
    return fig

# -----------------------------------------------------------------------------
def _collect_trace_axis_values(fig: go.Figure, axis: str = "x") -> np.ndarray:
    """Collect numeric axis values from all traces in a figure.

    Notes
    -----
    - Ignores non-numeric values.
    - Returns a 1D float array (may be empty).
    """
    vals: list[np.ndarray] = []
    for tr in getattr(fig, "data", []) or []:
        v = getattr(tr, axis, None)
        if v is None:
            continue

        # Try fast vector conversion first
        arr = None
        try:
            arr = np.asarray(v, dtype=float)
        except Exception:
            arr = None

        if arr is None:
            # Fallback element-wise conversion (handles mixed lists/tuples)
            tmp = []
            try:
                it = list(v) if isinstance(v, (list, tuple, np.ndarray, pd.Series)) else [v]
            except Exception:
                it = [v]
            for item in it:
                try:
                    f = float(item)
                    if np.isfinite(f):
                        tmp.append(f)
                except Exception:
                    continue
            arr = np.asarray(tmp, dtype=float)

        if arr.size:
            arr = arr.astype(float, copy=False).ravel()
            arr = arr[np.isfinite(arr)]
            if arr.size:
                vals.append(arr)

    if not vals:
        return np.asarray([], dtype=float)
    return np.concatenate(vals)


def _robust_range_linear(
    x: np.ndarray,
    q_low: float = 0.5,
    q_high: float = 99.5,
    pad_frac: float = 0.05,
) -> list[float] | None:
    x = np.asarray(x, dtype=float)
    x = x[np.isfinite(x)]
    if x.size < 2:
        return None

    lo = float(np.nanpercentile(x, q_low))
    hi = float(np.nanpercentile(x, q_high))
    if not (np.isfinite(lo) and np.isfinite(hi)):
        return None

    if hi < lo:
        lo, hi = hi, lo

    span = hi - lo
    if span <= 0:
        pad = max(1.0, abs(lo) * 0.1)
    else:
        pad = span * pad_frac

    return [lo - pad, hi + pad]


def _robust_range_log10(
    x: np.ndarray,
    q_low: float = 0.5,
    q_high: float = 99.5,
    pad_decades: float = 0.15,
    hard_min: float = 1e-6,
    hard_max: float = 1e8,
) -> list[float] | None:
    x = np.asarray(x, dtype=float)
    x = x[np.isfinite(x) & (x > 0)]
    if x.size < 2:
        return None

    lo = float(np.nanpercentile(x, q_low))
    hi = float(np.nanpercentile(x, q_high))
    if not (np.isfinite(lo) and np.isfinite(hi)):
        return None

    lo = max(lo, hard_min)
    hi = min(hi, hard_max)

    if hi <= lo:
        hi = lo * 10.0

    lo_log = math.log10(lo) - pad_decades
    hi_log = math.log10(hi) + pad_decades
    return [lo_log, hi_log]


def apply_robust_autoscale_x(
    fig: go.Figure,
    xlog: bool,
    *,
    q_low: float = 0.5,
    q_high: float = 99.5,
    pad_frac: float = 0.05,
    pad_decades: float = 0.15,
    hard_min: float = 1e-6,
    hard_max: float = 1e8,
) -> go.Figure:
    """Apply a robust x-axis range to avoid extreme/outlier-driven scaling.

    This prevents situations where a single outlier (or a misplaced vertical line)
    makes the curve appear 'angosta' by stretching the axis to unrealistic values.
    """
    try:
        # Skip 3D scenes / figures without xaxis (e.g., PNM 3D)
        if not hasattr(fig, "layout") or not getattr(fig.layout, "xaxis", None):
            return fig

        # Skip multi-x-axis layouts (e.g., petro logs with xaxis2) to avoid
        # unintentionally rescaling the secondary axis.
        layout_keys = set(fig.layout.to_plotly_json().keys())
        if any(k.startswith("xaxis") and k != "xaxis" for k in layout_keys):
            return fig

        xvals = _collect_trace_axis_values(fig, axis="x")
        if xvals.size < 2:
            return fig

        if xlog:
            rng = _robust_range_log10(
                xvals,
                q_low=q_low,
                q_high=q_high,
                pad_decades=pad_decades,
                hard_min=hard_min,
                hard_max=hard_max,
            )
            if rng is None:
                return fig
            fig.update_xaxes(type="log", range=rng, autorange=False)
        else:
            rng = _robust_range_linear(xvals, q_low=q_low, q_high=q_high, pad_frac=pad_frac)
            if rng is None:
                return fig
            fig.update_xaxes(type="linear", range=rng, autorange=False)
    except Exception:
        # Never fail the callback due to autoscale
        return fig

    return fig



def apply_smart_legend(
    fig: go.Figure,
    *,
    threshold_items: int = 6,
    max_items_per_row: int = 6,
    base_top: int = 60,
    row_height: int = 20,
) -> go.Figure:
    """Improve legend layout when many traces exist (multi-sample plots).

    Plotly legends can easily overlap titles / each other when you plot many
    samples. This helper forces a horizontal legend and increases the top margin
    so labels wrap cleanly (instead of stacking on top of each other).
    """

    try:
        if not getattr(fig, "data", None):
            return fig

        # Count legend-visible items
        names: List[str] = []
        for tr in fig.data:
            try:
                show = tr.showlegend if tr.showlegend is not None else True
            except Exception:
                show = True
            if not show:
                continue
            nm = getattr(tr, "name", None)
            if nm is None:
                continue
            nm = str(nm).strip()
            if not nm:
                continue
            names.append(nm)
        n = len(names)

        # Respect an explicit "legend below" placement (common in Multi‑Sample plots).
        # Several plotting functions intentionally place legends below the plot (y < 0)
        # to keep the title area clean.
        leg = getattr(fig.layout, "legend", None)
        legy = getattr(leg, "y", None) if leg is not None else None
        try:
            if legy is not None and float(legy) < 0:
                return fig
        except Exception:
            pass

        # Apply when we have many legend items OR when the figure already uses a top legend
        # (common in multi-sample plots where titles/legends can overlap).
        legend_at_top = False
        try:
            if legy is not None and float(legy) >= 0.95:
                legend_at_top = True
        except Exception:
            pass

        if (n <= threshold_items) and (not legend_at_top):
            return fig

        maxlen = max((len(nm) for nm in names), default=0)

        # Estimate a reasonable legend item width (px) so long sample names
        # don't overrun into the next row.
        itemwidth = int(max(110, min(240, maxlen * 7)))

        # Rough rows estimate -> expand the top margin to prevent overlap.
        # Use itemwidth and assume ~900 px of usable plot width.
        per_row_guess = max(3, int(900 / max(1, itemwidth)))
        per_row = max(1, min(int(max_items_per_row), per_row_guess))
        rows = int(math.ceil(n / per_row))
        # Reserve an extra band for the title above the legend rows.
        title_band = 34
        top_needed = int(base_top + title_band + rows * row_height)

        # Keep existing margins if they're already larger
        m = getattr(fig.layout, "margin", None)
        t_existing = (getattr(m, "t", None) or 0) if m is not None else 0

        fig.update_layout(
            legend=dict(
                orientation="h",
                x=0.0,
                xanchor="left",
                y=1.02,
                yanchor="bottom",
                itemwidth=itemwidth,
                font=dict(size=10),
                bgcolor="rgba(0,0,0,0)",
                title_text="",
            ),
            margin=dict(t=max(int(t_existing), top_needed)),
        )

        # Ensure title sits clearly above the legend band (avoids overlap).
        try:
            if getattr(fig.layout, "title", None) is not None:
                fig.update_layout(title=dict(y=1.18, yanchor="top", x=0.01, xanchor="left"))
        except Exception:
            pass

    except Exception:
        return fig

    return fig


AG_THEME = "ag-theme-alpine-dark"

# Plotly graph config: allow dragging the closure (conformance) vertical line
GRAPH_CONFIG: Dict[str, Any] = {
    "displaylogo": False,
    "scrollZoom": True,
    # Enable editing but restrict to shape movement (draggable vertical line)
    "editable": True,
    "edits": {
        "shapePosition": True,
        "annotationPosition": False,
        "annotationText": False,
        "axisTitleText": False,
        "legendPosition": False,
        "legendText": False,
        "titleText": False,
        "colorbarPosition": False,
        "colorbarTitleText": False,
    },
}

# Basic Plotly config for non-editable plots (XRD / auxiliary logs)
# NOTE: The main Pc plot uses GRAPH_CONFIG because it needs editable shapes
# (the draggable closure/conformance vertical line). For the other plots we
# keep editing disabled to avoid accidental modifications.
PLOTLY_CONFIG: Dict[str, Any] = {
    "displaylogo": False,
    "scrollZoom": True,
    "editable": False,
    "responsive": True,
}


# ---------------------------
# Data model
# ---------------------------
REQUIRED_COLS = ["Pressure", "CumVol", "IncVol"]
FLAG_COLS = ["Flag_Pressure_Down", "Flag_Cum_Down", "Flag_Inc_Neg_Fail"]

DERIVED_COLS = [
    "HgSat", "Sw",
    "r_um", "d_um",
    "dlogr", "dVdlogr", "dVdlogr_norm",
    "Pc_res_pa", "Height_m",
]

DEFAULT_PARAMS: Dict[str, Any] = {
    # Hg-air (MICP) parameters (typical)
    "sigma_hg_air_npm": 0.480,  # N/m
    "theta_hg_air_deg": 130.0,  # degrees

    # Reservoir scaling (oil/water by default, adjust as needed)
    "sigma_res_npm": 0.030,     # N/m
    "theta_res_deg": 30.0,      # degrees
    "rho_w_kgm3": 1000.0,
    "rho_hc_kgm3": 800.0,
    "fwl_depth_m": None,  # optional, used for Sw vs depth (TVD) plots

    # Swanson permeability: k = a * (Sp)^b
    "swanson_a": 374.25,
    "swanson_b": 1.9712,
    

    # Winland macro-normalized guardrails (bimodal)
    "winland_macro_frac_min": 0.15,
    "winland_macro_frac_max": 0.85,
    "winland_r35_macro_max_um": 50.0,

    # Bimodal Thomeer QC: minimum separation between Pd1 and Pd2 (log10 decades).
    # If separation is below this, the bimodal split becomes degenerate and may look unstable.
    "bimodal_pd_sep_min_log10": 0.8,
    # PNM (fast) defaults
    "pnm_coordination_z": 4.0,
    "pnm_tau_eff": None,  # if None, uses tortuosity/tortuosity_factor from meta or 2.5
    "pnm_constriction": 0.5,
    "pnm_psd_bins": 40,
    "pnm_k0_md_per_um2": 50.0,
    "pnm_calibrate_to_core": True,

    # PNM 3D visualization (synthetic network)
    "pnm3d_nodes": 350,
    "pnm3d_seed": "",  # if empty, deterministic from sample_id
    "pnm3d_bimodal_min_nodes": 160,
    # Plotly marker size is in px. Increase to make pore spheres larger.
    "pnm3d_sphere_scale": 2.5,


    # Rock typing thresholds (r35 in µm)
    "rt_bins_um": [0.01, 0.1, 1.0, 10.0],  # => RT5..RT1
    "rt_labels": ["RT5 (Nano)", "RT4 (Micro)", "RT3 (Meso)", "RT2 (Macro)", "RT1 (Mega)"],

    # KMeans-ish clustering settings
    "cluster_k": 3,

    # Overrides (optional)
    "phi_override_pct": None,    # e.g. 12.5
    "k_override_md": None,       # e.g. 0.12
    
    # Conformance correction (optional). Heuristic knee detection.
    "conf_method": "auto",  # auto/pre_pth/legacy
    # Use as an *assist* and always review visually.
    "conf_inc_frac": 0.20,          # fraction of peak Incremental used to detect entry (0.2 = 20%)
    "conf_min_consecutive": 3,      # consecutive points above threshold
    "conf_curv_ymin_frac": 0.01,    # curvature fallback search window (fraction of y_max)
    "conf_curv_ymax_frac": 0.60,
    # New robust conformance cut (saturation-based) — prevents selecting a late knee in high-perm/bimodal rocks
    "conf_max_sat_frac": 0.01,      # safety cap: never classify >1% of Vmax as conformance
    "conf_jump_sat_frac": 0.005,    # first ΔHgSat >= 0.5% => pore entry; cutoff is point before the jump



    # Threshold pressure (entry pressure) from curve
    # Computed as pressure where CumVol reaches (pth_frac_vmax * Vmax)
    "pth_frac_vmax": 0.02,


    # Threshold pressure (Pth) detection (acceleration / ΔS)  
    "pth_method": "ds_accel",
    "pth_ds_jump_pct": 1.0,
    "pth_ds_validate_pct": 0.10,

    # Backbone / Fractal proxy (pressure at effective saturation fraction)
    "backbone_sat_eff_frac": 0.45,

    # Plateau / fully-filled check (end-slope of intrusion curve)
    "plateau_tail_n": 5,
    "plateau_slope_pass": 1.0e-4,
    "plateau_slope_warn": 1.0e-3,
    # Petrophysical QC limits (broad physical plausibility checks)
        "petro_qc_limits": {
        # --------------------------
        # Petrophysical QAQC limits
        # --------------------------
        # Split into:
        #   - HARD: physically impossible / reject
        #   - WARN: suspicious / review manually
        #
        # NOTE: These defaults are based on your latest QC criteria (Hard vs Warn).

        # Porosity (%)
        "porosity_hard_min_pct": 0.0,
        "porosity_hard_max_pct": 100.0,
        "porosity_warn_high_pct": 45.0,

        # Bulk density at ~0.58 psia (g/mL)
        "bulk_density_hard_min_g_ml": 0.5,
        "bulk_density_hard_max_g_ml": 5.0,
        "bulk_density_warn_low_g_ml": 1.8,
        "bulk_density_warn_high_g_ml": 3.0,

        # Apparent/Skeletal density (g/mL) (proxy for grain density)
        "skeletal_density_hard_min_g_ml": 1.0,
        "skeletal_density_hard_max_g_ml": 6.0,
        "skeletal_density_warn_low_g_ml": 2.50,
        "skeletal_density_warn_high_g_ml": 2.95,

        # Permeability (mD)
        "permeability_hard_min_md": 0.0,
        "permeability_warn_high_md": 10000.0,  # >10D suggests fractures / broken sample

        # Stem volume used (%)
        "stem_used_hard_max_pct": 100.0,
        "stem_used_warn_high_pct": 85.0,

        # Threshold pressure (psia)
        "threshold_pressure_hard_min_psia": 0.0,
        "threshold_pressure_warn_low_psia": 0.5,  # <0.5 psia may indicate conforming/roughness fill

        # Max Hg saturation (% of pore volume)
        "max_shg_sat_hard_min_pct": 10.0,
        "max_shg_sat_warn_min_pct": 90.0,

        # Tortuosity (dimensionless)
        "tortuosity_hard_min": 1.0,
        "tortuosity_warn_high": 10.0,

        # Tortuosity factor (Archie 'a' proxy)
        "tortuosity_factor_hard_min": 0.1,
        "tortuosity_factor_hard_max": 5.0,
        "tortuosity_factor_warn_low": 0.6,
        "tortuosity_factor_warn_high": 2.0,

        # Formation factor (F = Ro/Rw) must be >= 1
        "formation_factor_hard_min": 1.0,

        # Extrusion efficiency (%). Typical Hg trapped ~30–50%, so <10% or >80% is suspicious.
        "extrusion_eff_hard_min_pct": 0.0,
        "extrusion_eff_hard_max_pct": 100.0,
        "extrusion_eff_warn_low_pct": 10.0,
        "extrusion_eff_warn_high_pct": 80.0,

        # Logic checks / derived (densities)
        "grain_bulk_warn_diff_g_ml": 0.1,  # warn if (rho_grain - rho_bulk) < 0.1 g/cc

        # Logical (model) checks
        "k_discrep_ratio_warn_low": 0.2,   # k_swanson / k_air
        "k_discrep_ratio_warn_high": 5.0,

        "m_cementation_warn_low": 1.1,
        "m_cementation_warn_high": 3.0,

        # Optional (if you have independent grain density from RCA)
        "rho_grain_diff_warn_g_ml": 0.05,
    },



    # -------------------------------------------------
    # Bi-modal pore-system hinting (auto-detection)
    # -------------------------------------------------
    # Used to suggest trying "Fit Thomeer (Bimodal)" when:
    #   - Unimodal Thomeer R² is low, and/or
    #   - PSD proxy (dV/dlog(r)) shows two distinct peaks.
    "bimodal_hint_r2_threshold": 0.93,      # if unimodal R² < threshold => suggest bimodal
    "bimodal_peak_min_frac": 0.35,          # peak must be >= (min_frac * max_peak) to count
    "bimodal_peak_min_sep_log10r": 0.40,    # minimum separation between peaks in log10(r_um)
    "bimodal_peak_smooth_window": 7,        # rolling median window (points) for PSD smoothing
}

# ---------------------------
# Parsing helpers (robust for AutoPore/Micromeritics exports)
# ---------------------------
def _norm(s: Any) -> str:
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return ""
    t = str(s).strip().lower()
    t = t.replace("á", "a").replace("é", "e").replace("í", "i").replace("ó", "o").replace("ú", "u")
    t = re.sub(r"\s+", " ", t)
    return t

def _coerce_numeric(series: pd.Series) -> pd.Series:
    s = series.astype(str)
    s = s.str.replace("\u00a0", " ", regex=False)
    s = s.str.replace(" ", "", regex=False)
    s = s.str.replace(",", ".", regex=False)
    return pd.to_numeric(s, errors="coerce")

def _is_pressure_header(x: Any) -> bool:
    x = _norm(x)
    return ("pressure" in x) and ("psia" in x or "psi" in x)

def _is_cum_intr_header(x: Any) -> bool:
    x = _norm(x)
    if ("cumulative" in x or x.startswith("cum")) and ("intrusion" in x or "intruded" in x):
        return True
    return ("cumulative" in x) and ("vol" in x)

def _is_inc_intr_header(x: Any) -> bool:
    x = _norm(x)
    if ("incremental" in x or x.startswith("inc")) and ("intrusion" in x or "intruded" in x):
        return True
    return ("incremental" in x) and ("vol" in x)

def _ensure_schema(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    for c in REQUIRED_COLS:
        if c not in df.columns:
            df[c] = pd.NA
    for c in FLAG_COLS:
        if c not in df.columns:
            df[c] = ""
    # Derived cols (kept, but not strictly required)
    for c in DERIVED_COLS:
        if c not in df.columns:
            df[c] = pd.NA

    # Coerce base numerics
    df["Pressure"] = _coerce_numeric(df["Pressure"])
    df["CumVol"] = _coerce_numeric(df["CumVol"])
    df["IncVol"] = _coerce_numeric(df["IncVol"])

    for c in FLAG_COLS:
        df[c] = df[c].fillna("").astype(str)

    return df

def _clean_imported_data(df: pd.DataFrame) -> pd.DataFrame:
    df = _ensure_schema(df)
    mask_all_nan = df["Pressure"].isna() & df["CumVol"].isna() & df["IncVol"].isna()
    df = df.loc[~mask_all_nan].copy()
    df = df.loc[df["Pressure"].notna()].reset_index(drop=True)
    return df

def _read_excel_bytes(decoded: bytes, ext: str) -> Tuple[pd.DataFrame, str]:
    """Read Excel-like bytes robustly.

    Notes:
      * .xlsx: uses openpyxl when available.
      * .xls: tries xlrd, then falls back to html-table or delimited text.

    Returns:
      (df_raw, fmt_tag)
    """
    ext = (ext or "").lower().strip()
    bio = io.BytesIO(decoded)

    # --- XLSX ---
    if ext == ".xlsx":
        try:
            df_raw = pd.read_excel(bio, sheet_name=0, engine="openpyxl", header=None)
            return df_raw, "excel.xlsx-raw"
        except Exception:
            # fallback to pandas default engine
            bio.seek(0)
            df_raw = pd.read_excel(bio, sheet_name=0, header=None)
            return df_raw, "excel.xlsx-raw"

    # --- XLS (or unknown Excel) ---
    # 1) Try xlrd (classic .xls)
    try:
        df_raw = pd.read_excel(io.BytesIO(decoded), sheet_name=0, engine="xlrd", header=None)
        return df_raw, f"excel{ext or '.xls'}-raw"
    except Exception:
        pass

    # 2) Try pandas default (sometimes works if engine installed)
    try:
        df_raw = pd.read_excel(io.BytesIO(decoded), sheet_name=0, header=None)
        return df_raw, f"excel{ext or '.xls'}-raw"
    except Exception:
        pass

    # 3) Some Micromeritics exports saved as .xls are actually HTML tables
    text = decoded.decode("utf-8", errors="ignore")
    if "<table" in text.lower():
        try:
            tables = pd.read_html(text)
            if tables:
                return tables[0], "html-table"
        except Exception:
            pass

    # 4) Fallback: try delimited text
    try:
        df = pd.read_csv(io.StringIO(text), sep="	")
        return df, "xls-as-tsv"
    except Exception:
        pass

    df = pd.read_csv(io.StringIO(text), delim_whitespace=True, on_bad_lines="skip")
    return df, "xls-as-whitespace"

def _read_uploaded(contents: str, filename: str) -> Tuple[pd.DataFrame, str]:
    # Data URL: "data:...;base64,<payload>"
    _, content_string = contents.split(",", 1)
    decoded = base64.b64decode(content_string)

    ext = os.path.splitext((filename or "").lower())[1]
    if ext in [".xls", ".xlsx"]:
        return _read_excel_bytes(decoded, ext)

    # CSV/TXT
    text = decoded.decode("utf-8", errors="replace")
    try:
        df = pd.read_csv(io.StringIO(text), sep=None, engine="python", on_bad_lines="skip")
        return df, "csv/auto-sep"
    except Exception:
        df = pd.read_csv(io.StringIO(text), delim_whitespace=True, on_bad_lines="skip")
        return df, "txt/whitespace"

def _score_numeric_block(tmp: pd.DataFrame, header_row: int, cols: Tuple[int, int, int], nlook: int = 25) -> float:
    p_idx, c_idx, i_idx = cols
    block = tmp.iloc[header_row + 1: header_row + 1 + nlook, [p_idx, c_idx, i_idx]].copy()
    pn = _coerce_numeric(block.iloc[:, 0]).notna().mean()
    cn = _coerce_numeric(block.iloc[:, 1]).notna().mean()
    inn = _coerce_numeric(block.iloc[:, 2]).notna().mean()

    # Bonus if pressure in cum table is identical (or very close) to pressure in inc table (if present)
    # If there is another pressure column immediately left of i_idx, compare quickly
    bonus = 0.0
    if i_idx - 1 >= 0:
        maybe_p2 = tmp.iloc[header_row + 1: header_row + 1 + nlook, i_idx - 1]
        p2 = _coerce_numeric(maybe_p2)
        p1 = _coerce_numeric(block.iloc[:, 0])
        if p1.notna().any() and p2.notna().any():
            # relative match on overlapping non-nans
            m = p1.notna() & p2.notna()
            if m.any():
                rel = (np.abs(p1[m] - p2[m]) / (np.abs(p1[m]) + 1e-12)).median()
                if np.isfinite(rel) and rel < 1e-6:
                    bonus = 0.15

    return (min(pn, cn, inn) * 0.65 + (pn + cn + inn) / 3.0 * 0.35) + bonus

def _find_best_threecol_table(tmp: pd.DataFrame, max_scan: int = 2500) -> Tuple[Optional[pd.DataFrame], str]:
    max_scan = min(len(tmp), max_scan)
    best_score = -1.0
    best = None  # (row, p_idx, c_idx, i_idx)

    for r in range(max_scan):
        row_vals = tmp.iloc[r].tolist()
        cum_idxs = [k for k, v in enumerate(row_vals) if _is_cum_intr_header(v)]
        inc_idxs = [k for k, v in enumerate(row_vals) if _is_inc_intr_header(v)]
        pres_idxs = [k for k, v in enumerate(row_vals) if _is_pressure_header(v)]
        if not cum_idxs or not inc_idxs or not pres_idxs:
            continue

        for c_idx in cum_idxs:
            near_p = [p for p in pres_idxs if 0 <= (c_idx - p) <= 2]
            if not near_p:
                near_p = sorted(pres_idxs, key=lambda p: abs(p - c_idx))
            p_idx = near_p[0]

            for i_idx in inc_idxs:
                cols = (p_idx, c_idx, i_idx)
                sc = _score_numeric_block(tmp, r, cols, nlook=25)
                if sc > best_score:
                    best_score = sc
                    best = (r, p_idx, c_idx, i_idx)

    if best is None:
        return None, "no_match_3cols"

    r, p_idx, c_idx, i_idx = best
    data = tmp.iloc[r + 1:, [p_idx, c_idx, i_idx]].copy()
    data.columns = ["Pressure", "CumVol", "IncVol"]
    data = _clean_imported_data(data)
    detail = f"best_3col_table: row={r}, cols=(P{p_idx},C{c_idx},I{i_idx}), score={best_score:.3f}"
    return data, detail

def _find_two_col_table(tmp: pd.DataFrame, kind: str, max_scan: int = 2500) -> Tuple[Optional[pd.DataFrame], str]:
    max_scan = min(len(tmp), max_scan)
    best_score = -1.0
    best = None  # (row, p_idx, v_idx)

    is_val = _is_cum_intr_header if kind == "cum" else _is_inc_intr_header

    for r in range(max_scan):
        row_vals = tmp.iloc[r].tolist()
        val_idxs = [k for k, v in enumerate(row_vals) if is_val(v)]
        pres_idxs = [k for k, v in enumerate(row_vals) if _is_pressure_header(v)]
        if not val_idxs or not pres_idxs:
            continue

        for v_idx in val_idxs:
            near_p = [p for p in pres_idxs if 0 <= (v_idx - p) <= 2]
            if not near_p:
                near_p = sorted(pres_idxs, key=lambda p: abs(p - v_idx))
            p_idx = near_p[0]

            block = tmp.iloc[r + 1: r + 1 + 25, [p_idx, v_idx]].copy()
            pn = _coerce_numeric(block.iloc[:, 0]).notna().mean()
            vn = _coerce_numeric(block.iloc[:, 1]).notna().mean()
            sc = min(pn, vn) * 0.7 + (pn + vn) / 2.0 * 0.3

            if sc > best_score:
                best_score = sc
                best = (r, p_idx, v_idx)

    if best is None:
        return None, f"no_match_{kind}_2cols"

    r, p_idx, v_idx = best
    data = tmp.iloc[r + 1:, [p_idx, v_idx]].copy()
    if kind == "cum":
        data.columns = ["Pressure", "CumVol"]
    else:
        data.columns = ["Pressure", "IncVol"]

    data["Pressure"] = _coerce_numeric(data["Pressure"])
    data = data.loc[data["Pressure"].notna()].reset_index(drop=True)

    detail = f"best_{kind}_2col_table: row={r}, cols=(P{p_idx},V{v_idx}), score={best_score:.3f}"
    return data, detail

def parse_to_three_cols(df_in: pd.DataFrame, fmt: str) -> Tuple[pd.DataFrame, str]:
    if df_in is None or df_in.empty:
        return _ensure_schema(pd.DataFrame({"Pressure": [], "CumVol": [], "IncVol": []})), "sin_data"

    # Excel raw: integer columns
    if df_in.columns.dtype.kind in ("i", "u"):
        tmp = df_in.copy()
        tmp.columns = range(tmp.shape[1])

        df3, detail3 = _find_best_threecol_table(tmp)
        if df3 is not None and not df3.empty:
            return _ensure_schema(df3), detail3

        # fallback: separate tables
        df_cum, d_cum = _find_two_col_table(tmp, "cum")
        df_inc, d_inc = _find_two_col_table(tmp, "inc")

        if df_cum is not None and df_inc is not None and not df_cum.empty:
            dfm = pd.merge(df_cum, df_inc, on="Pressure", how="outer")
            dfm = dfm.sort_values("Pressure", ascending=True).reset_index(drop=True)
            dfm = _ensure_schema(dfm)
            return dfm, f"merged_2tables: {d_cum} | {d_inc}"

        return _ensure_schema(pd.DataFrame({"Pressure": [], "CumVol": [], "IncVol": []})), f"excel_no_match: {detail3} | {d_cum} | {d_inc}"

    # CSV/TXT: header match or fallback to first 3 columns
    cols = list(df_in.columns)
    idx_p = idx_c = idx_i = None
    for j, c in enumerate(cols):
        if idx_p is None and _is_pressure_header(c):
            idx_p = j
        if idx_c is None and _is_cum_intr_header(c):
            idx_c = j
        if idx_i is None and _is_inc_intr_header(c):
            idx_i = j

    if idx_p is not None and idx_c is not None and idx_i is not None:
        out = pd.DataFrame({
            "Pressure": df_in.iloc[:, idx_p],
            "CumVol": df_in.iloc[:, idx_c],
            "IncVol": df_in.iloc[:, idx_i],
        })
        return _ensure_schema(_clean_imported_data(out)), "csv_headers_match"

    if df_in.shape[1] >= 3:
        out = pd.DataFrame({
            "Pressure": df_in.iloc[:, 0],
            "CumVol": df_in.iloc[:, 1],
            "IncVol": df_in.iloc[:, 2],
        })
        out = _ensure_schema(out)
        out = _clean_imported_data(out)
        return out, "fallback_first3cols"

    return _ensure_schema(pd.DataFrame({"Pressure": [], "CumVol": [], "IncVol": []})), "no_match"


# -----------------------------------------------------------------------------
# External N4 logs (LogN4A / LogN4_Vert) import utilities
# -----------------------------------------------------------------------------

def _logn4_coerce_float(x):
    """Best-effort numeric coercion for LogN4 worksheets."""
    if x is None or (isinstance(x, float) and np.isnan(x)):
        return np.nan
    if isinstance(x, (int, float, np.number)):
        try:
            return float(x)
        except Exception:
            return np.nan

    s = str(x).strip()
    if s == "" or s.lower() in {"nan", "none", "null"}:
        return np.nan

    # Common Excel artifacts / errors
    if "#ref" in s.lower() or "#n/a" in s.lower():
        return np.nan

    # Remove comparison signs and thousand separators
    s = s.replace(",", "")
    s = re.sub(r"[<>]", "", s)

    # Keep only the first numeric-like token if extra text exists
    m = re.search(r"[-+]?\d*\.?\d+(?:[eE][-+]?\d+)?", s)
    if not m:
        return np.nan
    try:
        return float(m.group(0))
    except Exception:
        return np.nan


def _logn4_make_unique(names):
    """Make column names unique while preserving order."""
    seen = {}
    out = []
    for n in names:
        if n not in seen:
            seen[n] = 1
            out.append(n)
        else:
            seen[n] += 1
            out.append(f"{n} ({seen[n]})")
    return out


def _parse_logn4_sheet(df_raw):
    """
    Parse a LogN4A / LogN4_Vert worksheet that typically contains a header row with:
    Depth, PorAmb, PorOB, PermAmb, PermOB, ...
    Returns a tidy DataFrame with a 'Depth' column plus any available logs.
    """
    if df_raw is None or df_raw.empty:
        return pd.DataFrame()

    # Find header row containing "Depth"
    header_row = None
    max_scan = min(len(df_raw), 80)
    for i in range(max_scan):
        row_vals = df_raw.iloc[i].values.tolist()
        for v in row_vals:
            if isinstance(v, str) and v.strip().lower() == "depth":
                header_row = i
                break
        if header_row is not None:
            break

    if header_row is None:
        # Fallback: any cell containing 'depth'
        for i in range(max_scan):
            row_vals = df_raw.iloc[i].values.tolist()
            if any(isinstance(v, str) and "depth" in v.strip().lower() for v in row_vals):
                header_row = i
                break

    if header_row is None:
        return pd.DataFrame()

    main = df_raw.iloc[header_row]
    prev = df_raw.iloc[header_row - 1] if header_row > 0 else None

    # Build column names (optionally prefixing with the group label above)
    colnames = []
    for j, v in enumerate(main.values.tolist()):
        if v is None or (isinstance(v, float) and np.isnan(v)) or str(v).strip() == "":
            colnames.append(None)
            continue
        name = re.sub(r"\s+", " ", str(v).strip())
        if prev is not None:
            g = prev.iloc[j]
            if g is not None and not (isinstance(g, float) and np.isnan(g)):
                gname = re.sub(r"\s+", " ", str(g).strip())
                if (
                    gname
                    and gname.lower() not in {"nan", "none"}
                    and gname != name
                    and gname not in name
                    and (" " in name)
                ):
                    # e.g., "Saturation" + "Oil" -> "Saturation Oil"
                    name = f"{gname} {name}"
        colnames.append(name)

    df = df_raw.iloc[header_row + 1 :].copy()
    df.columns = colnames
    df = df.loc[:, [c for c in df.columns if c is not None]]
    df = df.dropna(how="all")

    # Identify depth column
    depth_candidates = [c for c in df.columns if isinstance(c, str) and "depth" in c.lower()]
    depth_col = depth_candidates[0] if depth_candidates else (df.columns[0] if len(df.columns) else None)
    if depth_col is None:
        return pd.DataFrame()

    # Rename depth to canonical name
    df = df.rename(columns={depth_col: "Depth"})

    # Coerce numeric values
    df["Depth"] = pd.to_numeric(df["Depth"], errors="coerce")
    df = df.dropna(subset=["Depth"])

    for c in df.columns:
        if c == "Depth":
            continue
        df[c] = df[c].apply(_logn4_coerce_float)

    # Only drop all-null log columns if we actually have any rows.
    if len(df) > 0:
        drop_cols = [c for c in df.columns if c != "Depth" and df[c].notna().sum() == 0]
        if drop_cols:
            df = df.drop(columns=drop_cols)

    df.columns = _logn4_make_unique(list(df.columns))
    df = df.sort_values("Depth").reset_index(drop=True)
    return df


def _read_logn4_from_excel_bytes(decoded, filename=""):
    """Parse external core-log Excel workbooks into {sheet_name: DataFrame}.

    The parser is intentionally permissive: any sheet that contains a usable **Depth** column and at
    least one additional log column will be imported.

    This supports:
    - Classic LogN4 sheets (LogN4A / LogN4_Vert)
    - Plug/property tables (e.g., Helium porosity, Gas/Air permeability, Klinkenberg)

    The returned data is later merged into the store-logn4 structure by `_merge_logn4_store`.
    """
    out = {}
    if not decoded:
        return out

    bio = io.BytesIO(decoded)

    def _try_parse(raw_df):
        df = _parse_logn4_sheet(raw_df)
        if df is None or df.empty:
            return None
        if "Depth" not in df.columns:
            return None
        if len(df.columns) <= 1:
            return None
        return df

    # Primary: use ExcelFile (faster + stable for multiple sheets)
    try:
        xl = pd.ExcelFile(bio, engine="openpyxl")
        for sn in xl.sheet_names:
            try:
                raw = pd.read_excel(xl, sheet_name=sn, header=None)
            except Exception:
                continue
            df = _try_parse(raw)
            if df is not None:
                out[sn] = df
        return out
    except Exception:
        pass

    # Fallback: best-effort direct reads
    try:
        bio.seek(0)
    except Exception:
        pass
    try:
        try:
            bio.seek(0)
        except Exception:
            pass
        xl2 = pd.ExcelFile(bio)
        sheet_names = xl2.sheet_names
    except Exception:
        sheet_names = []

    if not sheet_names:
        sheet_names = ["LogN4A", "LogN4_Vert"]

    for sn in sheet_names:
        try:
            try:
                bio.seek(0)
            except Exception:
                pass
            raw = pd.read_excel(bio, sheet_name=sn, header=None)
        except Exception:
            continue
        df = _try_parse(raw)
        if df is not None:
            out[sn] = df

    return out


def _merge_logn4_store(store, new_sheets, filename=""):
    """
    Merge parsed sheets into store-logn4 structure.
    Store format:
        {
          "files": [{"name": ..., "ts": ...}, ...],
          "sheets": {
             "LogN4A": {"columns": [...], "records": [...]},
             "LogN4_Vert": {...}
          }
        }
    """
    store = store or {}
    store.setdefault("files", [])
    store.setdefault("sheets", {})

    now_ts = int(time.time() * 1000)
    if filename:
        store["files"].append({"name": filename, "ts": now_ts})

    for sn, df in (new_sheets or {}).items():
        if df is None:
            continue

        existing = store["sheets"].get(sn)
        if existing and isinstance(existing, dict) and existing.get("records"):
            try:
                df_old = pd.DataFrame(existing.get("records", []))
                if "Depth" in df_old.columns and "Depth" in df.columns and (len(df_old) > 0 or len(df) > 0):
                    df_old = df_old.set_index("Depth")
                    df_new = df.set_index("Depth")
                    # Prefer new values where present, otherwise keep old.
                    df_merged = df_new.combine_first(df_old).reset_index()
                else:
                    # If Depth is missing, just concatenate (best effort)
                    df_merged = pd.concat([df_old, df], ignore_index=True)
                df = df_merged
            except Exception:
                pass

        store["sheets"][sn] = {
            "columns": list(df.columns),
            "records": df.to_dict("records"),
        }

    store["ts"] = now_ts
    return store


def _summarize_logn4_store(store):
    if not store or not isinstance(store, dict):
        return "No external core log data loaded."

    sheets = store.get("sheets", {}) or {}
    if not sheets:
        return "No external core log data loaded."

    parts = []
    for sn, payload in sheets.items():
        cols = (payload or {}).get("columns", []) or []
        nrows = len((payload or {}).get("records", []) or [])
        nlogs = max(0, len([c for c in cols if c != "Depth"]))
        parts.append(f"{sn}: {nrows} rows, {nlogs} logs")

    files = store.get("files", []) or []
    file_part = f"{len(files)} file(s)" if files else "0 file(s)"
    return f"Loaded external logs from {file_part}. " + " | ".join(parts)



# ---------------------------
# Core Validation helpers
# ---------------------------

def _canon_key(x) -> str:
    try:
        return re.sub(r"[^a-z0-9]", "", str(x).lower())
    except Exception:
        return ""


def _extract_depth_from_sample_id(sample_id: str):
    """Best-effort depth extraction from sample id/filename.

    We pick the last numeric token >= 100 (to avoid N2V3 style tokens).
    Returns float depth or None.
    """
    if not sample_id:
        return None
    nums = re.findall(r"\d+(?:\.\d+)?", str(sample_id))
    if not nums:
        return None
    vals = []
    for n in nums:
        try:
            vals.append(float(n))
        except Exception:
            pass
    if not vals:
        return None
    big = [v for v in vals if v >= 100]
    return big[-1] if big else vals[-1]


def _pick_record_value(record: dict, hints) -> float | None:
    """Pick a numeric value from a record dict using fuzzy key matching."""
    if not isinstance(record, dict) or not record:
        return None
    hints = [str(h).lower() for h in (hints or [])]
    for k, v in record.items():
        ck = _canon_key(k)
        if not ck:
            continue
        for h in hints:
            if h and h in ck:
                try:
                    val = _safe_float(v)
                except Exception:
                    try:
                        val = float(v)
                    except Exception:
                        val = None
                if val is None:
                    continue
                if isinstance(val, float) and (math.isnan(val) or math.isinf(val)):
                    continue
                return float(val)
    return None


def _find_nearest_logn4_record(logn4_store: dict, sheet_name: str, target_depth: float | None):
    """Return (record_dict, depth_used, depth_delta)."""
    try:
        sheets = (logn4_store or {}).get("sheets", {})
        recs = (sheets.get(sheet_name) or {}).get("records", [])
    except Exception:
        recs = []
    if not recs:
        return None, None, None

    df = pd.DataFrame(recs)
    if df.empty:
        return None, None, None

    # Find a depth-like column
    depth_col = None
    for c in df.columns:
        cc = _canon_key(c)
        if cc in ("depth", "dept", "md", "tvd", "tvdss") or ("depth" in cc) or ("dept" in cc):
            depth_col = c
            break
    if depth_col is None:
        return None, None, None

    depth_num = pd.to_numeric(df[depth_col], errors="coerce")
    df2 = df.copy()
    df2["__depth__"] = depth_num
    df2 = df2.dropna(subset=["__depth__"])
    if df2.empty:
        return None, None, None

    if target_depth is None:
        row = df2.iloc[0]
        return row.to_dict(), float(row["__depth__"]), None

    diffs = (df2["__depth__"] - float(target_depth)).abs()
    if diffs.isna().all():
        return None, None, None
    idx = diffs.idxmin()
    row = df2.loc[idx]
    return row.to_dict(), float(row["__depth__"]), float(diffs.loc[idx])


def _parse_external_perm_row(row: dict) -> dict:
    """Extract permeability-related values from an arbitrary external-log row dict."""
    out: dict = {}
    if not isinstance(row, dict):
        return out

    for k, v in row.items():
        lk = str(k).strip().lower()
        val = _num(v)
        if val is None:
            continue

        # Ambient / Overburden permeability (LogN4 style)
        if "permamb" in lk:
            out["permamb_md"] = val
            continue
        if "permob" in lk:
            out["permob_md"] = val
            continue

        # Klinkenberg-corrected permeability
        if "klink" in lk:
            out["k_klinkenberg_md"] = val
            continue

        # Gas / air / helium permeability (various naming conventions)
        if ("permeability" in lk and ("air" in lk or "gas" in lk or "helium" in lk)) or lk in {"kair", "k_air", "k-air"}:
            out["k_air_md"] = val
            continue
        if "helium" in lk and "perm" in lk:
            out["k_helium_md"] = val
            continue

    # Many datasets report only gas/air permeability + Klinkenberg; treat gas/air as "helium" if helium is missing.
    if "k_helium_md" not in out and "k_air_md" in out:
        out["k_helium_md"] = out["k_air_md"]

    return out


def _best_external_perm_at_depth(logn4_store: dict, depth_ft: float | None) -> dict:
    """Best-effort lookup of external permeability values near a given depth.

    Searches across all imported sheets and returns a dict with (some of):
      - permamb_md, permob_md
      - k_air_md, k_helium_md
      - k_klinkenberg_md

    Selection strategy:
      1) Prefer sheets that provide more of these values.
      2) Break ties using smallest |Depth - depth_ft|.
    """
    sheets = (logn4_store or {}).get("sheets", {}) or {}
    depth_ft = _num(depth_ft)
    if depth_ft is None or not sheets:
        return {}

    best_vals: dict = {}
    best_score = -1
    best_dist = None

    for sheet_name, sh in sheets.items():
        cols = sh.get("columns") or []
        cols_low = [str(c).lower() for c in cols]
        if not any(("perm" in c or "klink" in c or "permeability" in c) for c in cols_low):
            continue

        rec, actual_depth, _delta = _find_nearest_logn4_record(logn4_store, sheet_name, depth_ft)
        if not rec:
            continue

        vals = _parse_external_perm_row(rec)
        if not vals:
            continue

        score = sum(1 for k in ("k_klinkenberg_md", "k_air_md", "k_helium_md", "permamb_md", "permob_md") if vals.get(k) is not None)
        dist = None
        ad = _num(actual_depth)
        if ad is not None:
            dist = abs(ad - depth_ft)

        if score > best_score or (score == best_score and dist is not None and (best_dist is None or dist < best_dist)):
            best_vals = vals
            best_score = score
            best_dist = dist

    return best_vals



def _safe_div(a, b):
    try:
        a = float(a)
        b = float(b)
        if b == 0:
            return None
        if math.isnan(a) or math.isnan(b) or math.isinf(a) or math.isinf(b):
            return None
        return a / b
    except Exception:
        return None


def _fmt_num(x, ndigits: int = 4):
    if x is None:
        return "n/a"
    try:
        x = float(x)
        if math.isnan(x) or math.isinf(x):
            return "n/a"
        return f"{x:.{ndigits}g}"
    except Exception:
        return "n/a"


def _bimodal_supported_flag(pd_ratio=None, stress_ratio=None, clay_pct=None):
    """Heuristic flag.

    Score components:
      - Pd2/Pd1 >= 3 -> +2, >=2 -> +1
      - PermOB/PermAmb <= 0.7 -> +1
      - Clay% >= 10 (optional) -> +1

    Supported if score >= 2.
    """
    score = 0
    reasons = []

    if pd_ratio is not None:
        try:
            r = float(pd_ratio)
            if not (math.isnan(r) or math.isinf(r)):
                if r >= 3:
                    score += 2
                    reasons.append("Pd2/Pd1 ≥ 3")
                elif r >= 2:
                    score += 1
                    reasons.append("Pd2/Pd1 ≥ 2")
        except Exception:
            pass

    if stress_ratio is not None:
        try:
            sr = float(stress_ratio)
            if not (math.isnan(sr) or math.isinf(sr)):
                if sr <= 0.7:
                    score += 1
                    reasons.append("PermOB/PermAmb ≤ 0.7")
        except Exception:
            pass

    if clay_pct is not None:
        try:
            cp = float(clay_pct)
            if not (math.isnan(cp) or math.isinf(cp)):
                if cp >= 10:
                    score += 1
                    reasons.append("Clay ≥ 10%")
        except Exception:
            pass

    return (score >= 2), score, reasons

# ---------------------------
# Metadata extraction (Micromeritics-like report)
# ---------------------------
def _as_str(x: Any) -> str:
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return ""
    return str(x)

def _try_float(x: Any) -> Optional[float]:
    try:
        s = str(x).strip()
        s = s.replace("\u00a0", " ").replace(" ", "").replace(",", ".")
        return float(s)
    except Exception:
        return None

def _find_row_col(df: pd.DataFrame, pattern: str, max_scan_rows: int = 500) -> Optional[Tuple[int, int, str]]:
    pat = re.compile(pattern, re.IGNORECASE)
    nrows = min(df.shape[0], max_scan_rows)
    for r in range(nrows):
        for c in range(df.shape[1]):
            v = df.iat[r, c]
            if isinstance(v, str) and pat.search(v):
                return r, c, v
    return None

def extract_metadata(df_raw: pd.DataFrame) -> Dict[str, Any]:
    """
    Best-effort extraction for:
      - sample_id
      - porosity_pct
      - permeability_md
      - threshold_pressure_psia
    """
    meta: Dict[str, Any] = {}

    # Sample ID
    hit = _find_row_col(df_raw, r"sample\s*id")
    if hit:
        r, c, _ = hit
        # prefer right cell
        if c + 1 < df_raw.shape[1]:
            sample_id = _as_str(df_raw.iat[r, c + 1]).strip()
            if sample_id:
                meta["sample_id"] = sample_id
        # fallback: parse from same cell
        if "sample_id" not in meta:
            m = re.search(r"sample\s*id\s*:\s*(.*)", _as_str(df_raw.iat[r, c]), re.IGNORECASE)
            if m:
                meta["sample_id"] = m.group(1).strip()

    # Porosity
    hit = _find_row_col(df_raw, r"\bporosity\b")
    if hit:
        r, c, txt = hit
        # Most Micromeritics exports: "Porosity =" in col 0 and value in col 1
        if c + 1 < df_raw.shape[1]:
            v = _try_float(df_raw.iat[r, c + 1])
            if v is not None:
                meta["porosity_pct"] = v
        if "porosity_pct" not in meta:
            m = re.search(r"porosity\s*=\s*([0-9\.,]+)", txt, re.IGNORECASE)
            if m:
                meta["porosity_pct"] = float(m.group(1).replace(",", "."))

    # Permeability
    hit = _find_row_col(df_raw, r"\bpermeability\s*=")
    if hit:
        r, c, txt = hit
        if c + 1 < df_raw.shape[1]:
            v = _try_float(df_raw.iat[r, c + 1])
            if v is not None:
                meta["permeability_md"] = v
        if "permeability_md" not in meta:
            m = re.search(r"permeability\s*=\s*([0-9\.,]+)", txt, re.IGNORECASE)
            if m:
                meta["permeability_md"] = float(m.group(1).replace(",", "."))

    # Threshold pressure (optional)
    hit = _find_row_col(df_raw, r"threshold\s+pressure")
    if hit:
        r, c, txt = hit
        if c + 1 < df_raw.shape[1]:
            v = _try_float(df_raw.iat[r, c + 1])
            if v is not None:
                meta["threshold_pressure_psia"] = v


    # Bulk density (g/mL)
    hit = _find_row_col(df_raw, r"\bbulk\s*density\b")
    if hit:
        r, c, txt = hit
        if c + 1 < df_raw.shape[1]:
            v = _try_float(df_raw.iat[r, c + 1])
            if v is not None:
                meta["bulk_density_g_ml"] = v
        if "bulk_density_g_ml" not in meta:
            m = re.search(r"bulk\s*density.*?=\s*([0-9\.,]+)", txt, re.IGNORECASE)
            if m:
                meta["bulk_density_g_ml"] = float(m.group(1).replace(",", "."))

    # Apparent / skeletal density (g/mL)
    hit = _find_row_col(df_raw, r"apparent.*density")
    if hit:
        r, c, txt = hit
        if c + 1 < df_raw.shape[1]:
            v = _try_float(df_raw.iat[r, c + 1])
            if v is not None:
                meta["skeletal_density_g_ml"] = v
        if "skeletal_density_g_ml" not in meta:
            m = re.search(r"apparent.*density.*?=\s*([0-9\.,]+)", txt, re.IGNORECASE)
            if m:
                meta["skeletal_density_g_ml"] = float(m.group(1).replace(",", "."))

    # Alternate label: Skeletal density (g/mL)
    if "skeletal_density_g_ml" not in meta:
        hit = _find_row_col(df_raw, r"\bskeletal\s*density\b")
        if hit:
            r, c, txt = hit
            if c + 1 < df_raw.shape[1]:
                v = _try_float(df_raw.iat[r, c + 1])
                if v is not None:
                    meta["skeletal_density_g_ml"] = v
            if "skeletal_density_g_ml" not in meta:
                m = re.search(r"skeletal\s*density.*?=\s*([0-9\.,]+)", txt, re.IGNORECASE)
                if m:
                    meta["skeletal_density_g_ml"] = float(m.group(1).replace(",", "."))

    # Stem volume (mL)
    hit = _find_row_col(df_raw, r"stem\s*volume\s*:")
    if hit:
        r, c, txt = hit
        if c + 1 < df_raw.shape[1]:
            v = _try_float(df_raw.iat[r, c + 1])
            if v is not None:
                meta["stem_volume_ml"] = v
        if "stem_volume_ml" not in meta:
            m = re.search(r"stem\s*volume\s*:\s*([0-9\.,]+)", txt, re.IGNORECASE)
            if m:
                meta["stem_volume_ml"] = float(m.group(1).replace(",", "."))

    # Stem volume used (%)
    hit = _find_row_col(df_raw, r"stem\s*volume\s*used")
    if hit:
        r, c, txt = hit
        if c + 1 < df_raw.shape[1]:
            v = _try_float(df_raw.iat[r, c + 1])
            if v is not None:
                meta["stem_volume_used_pct"] = v
        if "stem_volume_used_pct" not in meta:
            m = re.search(r"stem\s*volume\s*used.*?([0-9\.,]+)\s*%?", txt, re.IGNORECASE)
            if m:
                meta["stem_volume_used_pct"] = float(m.group(1).replace(",", "."))

    # Tortuosity factor
    hit = _find_row_col(df_raw, r"tortuosity\s*factor")
    if hit:
        r, c, txt = hit
        if c + 1 < df_raw.shape[1]:
            v = _try_float(df_raw.iat[r, c + 1])
            if v is not None:
                meta["tortuosity_factor"] = v
        if "tortuosity_factor" not in meta:
            m = re.search(r"tortuosity\s*factor.*?=\s*([0-9\.,]+)", txt, re.IGNORECASE)
            if m:
                meta["tortuosity_factor"] = float(m.group(1).replace(",", "."))

    # Tortuosity
    hit = _find_row_col(df_raw, r"\btortuosity\b(?!\s*factor)")
    if hit:
        r, c, txt = hit
        if c + 1 < df_raw.shape[1]:
            v = _try_float(df_raw.iat[r, c + 1])
            if v is not None:
                meta["tortuosity"] = v
        if "tortuosity" not in meta:
            m = re.search(r"\btortuosity\b.*?=\s*([0-9\.,]+)", txt, re.IGNORECASE)
            if m:
                meta["tortuosity"] = float(m.group(1).replace(",", "."))

    # Formation factor
    # NOTE: Some MICP reports provide **Conductivity formation factor** (= 1/F). If label contains
    # 'conductivity', we invert it to get the standard Formation Factor F = Ro/Rw (must be >= 1).
    hit = _find_row_col(df_raw, r"formation\s*factor")
    if hit:
        r, c, txt = hit
        v = None
        if c + 1 < df_raw.shape[1]:
            v = _try_float(df_raw.iat[r, c + 1])
        if v is None:
            m = re.search(r"formation\s*factor.*?=\s*([0-9\.,]+)", str(txt), re.IGNORECASE)
            if m:
                v = float(m.group(1).replace(",", "."))
        if v is not None:
            if "conductivity" in str(txt).lower() and float(v) != 0:
                meta["formation_factor"] = 1.0 / float(v)
                meta["formation_factor_raw"] = float(v)
            else:
                meta["formation_factor"] = float(v)



    # Extrusion efficiency (if present)
    hit = _find_row_col(df_raw, r"extrusion\s*efficiency")
    if hit:
        r, c, txt = hit
        v = None
        if c + 1 < df_raw.shape[1]:
            v = _try_float(df_raw.iat[r, c + 1])
        if v is None:
            m = re.search(r"extrusion\s*efficiency.*?=\s*([0-9\.,]+)", str(txt), re.IGNORECASE)
            if m:
                v = float(m.group(1).replace(",", "."))
        if v is not None:
            meta["extrusion_efficiency_pct"] = float(v)

    # Well name heuristic
    sid = meta.get("sample_id")
    if sid and isinstance(sid, str):
        meta["well_guess"] = sid.split()[0].strip()

    # Derive WELL / CORE / DEPTH from Sample ID (best-effort, used by PetroQC table)
    sid = meta.get("sample_id")
    if sid and isinstance(sid, str):
        sid_norm = re.sub(r"[_]+", " ", sid).strip()
        parts = sid_norm.split()

        if parts:
            meta.setdefault("well_guess", parts[0].strip())

        # Core and depth are typically encoded like: WELL  CORE-DEPTH
        # Example: "KAN-1EXP ST-N2V3-2918.93" -> core="ST-N2V3", depth=2918.93
        if len(parts) >= 2:
            core_tok = parts[1].strip()

            m = re.search(r"(\d+(?:\.\d+)?)\s*$", core_tok)
            if m:
                # Depth at the end of the 2nd token
                try:
                    depth_val = float(m.group(1))
                    meta.setdefault("depth_m", depth_val)
                except Exception:
                    pass

                core_guess = core_tok[: m.start()].rstrip("-_ ")
                if core_guess:
                    meta.setdefault("core_guess", core_guess)
            else:
                # Depth at the end of the full string
                m2 = re.search(r"(\d+(?:\.\d+)?)\s*$", sid_norm)
                if m2:
                    try:
                        meta.setdefault("depth_m", float(m2.group(1)))
                    except Exception:
                        pass

                if len(parts) > 2:
                    meta.setdefault("core_guess", " ".join(parts[1:-1]).strip())
                else:
                    meta.setdefault("core_guess", core_tok)

    return meta

def extract_meta_from_raw(df_raw: pd.DataFrame, filename: str | None = None) -> Dict[str, Any]:
    """Backward-compatible wrapper.

    Older versions referenced `extract_meta_from_raw(...)`. The canonical extractor is
    `extract_metadata(df_raw)`. This wrapper also adds filename-derived hints when needed.
    """
    meta = extract_metadata(df_raw) or {}
    if filename:
        fn = str(filename)
        meta.setdefault("source_filename", os.path.basename(fn))
        # Use filename stem as sample_id if missing
        if not meta.get("sample_id"):
            meta["sample_id"] = os.path.splitext(os.path.basename(fn))[0]

    # Best-effort well/depth guess from sample_id (common pattern: "WELL ...-DEPTH")
    sid = str(meta.get("sample_id") or "")
    if sid:
        meta.setdefault("well_guess", sid.split()[0])
        m_depth = re.search(r"(\d{3,5}(?:\.\d+)?)\s*$", sid)
        if m_depth and "depth_m" not in meta:
            try:
                meta["depth_m"] = float(m_depth.group(1))
            except Exception:
                pass
    return meta



def petrophysical_qaqc(
    meta: Dict[str, Any],
    params: Dict[str, Any],
    df: Optional[pd.DataFrame] = None,
    res: Optional[Dict[str, Any]] = None,
) -> Dict[str, Any]:
    """
    Petrophysical QA/QC (sample-level) to catch physically impossible values (HARD -> Reject)
    and suspicious values (WARN -> Manual review).

    This is intentionally separate from curve-level QAQC flags, because a sample can have a "nice"
    intrusion curve but still be unusable due to metadata / petrophysical inconsistencies.

    Returns a dict with:
      - petro_qc_done: bool
      - petro_qc_grade: PASS / WARN / FAIL
      - petro_qc_recommendation: ACCEPT / REVIEW / DISCARD
      - petro_qc_issues: list[dict] with (level, code, message, value)
      - derived metrics used by PetroQC table (max_shg_sat_pct, conformance_vol_pct, etc.)
    """

    meta = meta or {}
    params = params or DEFAULT_PARAMS
    res = res or {}

    limits = (params.get("petro_qc_limits") or DEFAULT_PARAMS.get("petro_qc_limits") or {}).copy()

    issues: List[Dict[str, Any]] = []

    def _add(level: str, code: str, message: str, value: Any = None) -> None:
        item = {"level": level, "code": code, "message": message}
        if value is not None:
            item["value"] = value
        issues.append(item)

    def _num(v: Any) -> Optional[float]:
        return _try_float(v)

    # --------------------------
    # Grab raw inputs / overrides
    # --------------------------
    phi_override = _num(params.get("phi_override_pct"))
    k_override = _num(params.get("k_override_md"))

    porosity_pct = phi_override if phi_override is not None else _num(meta.get("porosity_pct"))
    k_air_md = k_override if k_override is not None else _num(meta.get("permeability_md"))

    bulk_density_g_ml = _num(meta.get("bulk_density_g_ml"))
    skeletal_density_g_ml = _num(meta.get("skeletal_density_g_ml"))

    stem_used_pct = _num(meta.get("stem_volume_used_pct"))
    threshold_psia = _num(meta.get("threshold_pressure_psia"))

    tortuosity = _num(meta.get("tortuosity"))
    tort_factor = _num(meta.get("tortuosity_factor"))

    formation_factor = _num(meta.get("formation_factor"))

    extrusion_eff_pct = _num(meta.get("extrusion_efficiency_pct"))

    # --------------------------
    # Derived quantities
    # --------------------------
    dfx: Optional[pd.DataFrame] = None
    cumvol_max = None
    if df is not None:
        try:
            dfx = _ensure_schema(df if isinstance(df, pd.DataFrame) else pd.DataFrame(df))

            # CumVol is expected to be in mL/g for the imported intrusion table.
            # Keep backward-compatibility if a legacy column name is found.
            if len(dfx) > 0:
                if "CumVol" in dfx.columns:
                    cv = pd.to_numeric(dfx["CumVol"], errors="coerce").to_numpy(dtype=float)
                    if np.isfinite(cv).any():
                        cumvol_max = float(np.nanmax(cv))
                elif "CumVol_ml_g" in dfx.columns:
                    cv = pd.to_numeric(dfx["CumVol_ml_g"], errors="coerce").to_numpy(dtype=float)
                    if np.isfinite(cv).any():
                        cumvol_max = float(np.nanmax(cv))
        except Exception:
            dfx = None
            cumvol_max = None

    # Threshold pressure:
    # - keep the value parsed from the Excel report (if present),
    # - but also compute a robust "curve-based" threshold to avoid early conformance.
    threshold_report = threshold_psia
    threshold_curve = None
    threshold_method = "report" if (threshold_report is not None) else ""
    threshold_detail = ""
    if dfx is not None:
        try:
            pth, info = compute_threshold_pressure_psia(dfx, params)
            if pth is not None and np.isfinite(pth):
                threshold_curve = float(pth)
                threshold_method = str(info.get("method", "curve"))
                threshold_detail = str(info.get("detail", ""))
        except Exception:
            pass

    # Use curve-based threshold when available, otherwise fallback to report value.
    if threshold_curve is not None and np.isfinite(threshold_curve):
        threshold_psia = float(threshold_curve)
    else:
        threshold_psia = threshold_report


    # Backbone / Fractal proxy pressure (after any conformance correction)
    backbone_psia = None
    backbone_info: Dict[str, Any] = {}
    if dfx is not None:
        try:
            backbone_psia, backbone_info = compute_backbone_pressure_psia(dfx, params)
        except Exception:
            backbone_psia, backbone_info = None, {"method": "error"}


    # Max Hg saturation (% of pore volume). Needs porosity + bulk density + cumvol.
    max_shg_sat_pct = None
    if (porosity_pct is not None) and (bulk_density_g_ml is not None) and (cumvol_max is not None):
        try:
            phi_frac = float(porosity_pct) / 100.0
            if phi_frac > 0 and bulk_density_g_ml > 0:
                pore_vol_ml_g = phi_frac / float(bulk_density_g_ml)  # (mL pore) / g
                if pore_vol_ml_g > 0:
                    max_shg_sat_pct = 100.0 * float(cumvol_max) / pore_vol_ml_g
        except Exception:
            max_shg_sat_pct = None

    # Grain - bulk density difference (g/mL)
    grain_density_diff = None
    if (skeletal_density_g_ml is not None) and (bulk_density_g_ml is not None):
        grain_density_diff = float(skeletal_density_g_ml) - float(bulk_density_g_ml)

    # Conformance volume percent (if knee detected or already present in results)
    conformance_vol_pct = None
    conf_vol_ml_g = _num(res.get("conf_vol_ml_g"))
    conf_pknee_psia = _num(res.get("conf_pknee_psia"))

    if (conf_vol_ml_g is None or conf_pknee_psia is None) and dfx is not None:
        try:
            pknee, vconf, _method = detect_conformance_knee(dfx, params)
            if conf_pknee_psia is None and pknee is not None:
                conf_pknee_psia = float(pknee)
            if conf_vol_ml_g is None and vconf is not None:
                conf_vol_ml_g = float(vconf)
        except Exception:
            pass

    if conf_vol_ml_g is not None and cumvol_max is not None and cumvol_max > 0:
        conformance_vol_pct = 100.0 * float(conf_vol_ml_g) / float(cumvol_max)

    # Swanson permeability discrepancy ratio (k_swanson/k_air)
    k_swanson_md = _num(res.get("k_swanson_md"))
    k_ratio = None
    if k_swanson_md is None and dfx is not None:
        try:
            ks = compute_swanson_k(dfx, params, meta)
            if ks and ks.get("k_swanson_md") is not None:
                k_swanson_md = float(ks["k_swanson_md"])
        except Exception:
            pass
    if (k_swanson_md is not None) and (k_air_md is not None) and k_air_md > 0:
        k_ratio = float(k_swanson_md) / float(k_air_md)

    # Archie cementation exponent m (if possible)
    m_cementation = None
    if (
        formation_factor is not None
        and formation_factor > 0
        and porosity_pct is not None
        and porosity_pct > 0
        and porosity_pct < 100
    ):
        try:
            a_archie = tort_factor if (tort_factor is not None and tort_factor > 0) else 1.0
            phi_frac = float(porosity_pct) / 100.0
            if 0 < phi_frac < 1 and a_archie > 0:
                # F = a / phi^m -> m = ln(a/F) / ln(phi)
                m_cementation = math.log(a_archie / float(formation_factor)) / math.log(phi_frac)
        except Exception:
            m_cementation = None

    # --------------------------
    # Rule checks (Hard vs Warn)
    # --------------------------
    # Porosity
    if porosity_pct is None:
        _add("WARN", "MISSING_POROSITY", "Porosity not found in metadata.")
    else:
        if porosity_pct < limits.get("porosity_hard_min_pct", 0.0) or porosity_pct > limits.get(
            "porosity_hard_max_pct", 100.0
        ):
            _add(
                "FAIL",
                "POROSITY_HARD",
                "Porosity is outside physical bounds (0–100%).",
                porosity_pct,
            )
        elif porosity_pct > limits.get("porosity_warn_high_pct", 45.0):
            _add(
                "WARN",
                "POROSITY_WARN",
                "Porosity is unusually high for consolidated rocks.",
                porosity_pct,
            )

    # Bulk density
    if bulk_density_g_ml is None:
        _add("WARN", "MISSING_BULK_DENSITY", "Bulk density not found in metadata.")
    else:
        if bulk_density_g_ml < limits.get("bulk_density_hard_min_g_ml", 0.5) or bulk_density_g_ml > limits.get(
            "bulk_density_hard_max_g_ml", 5.0
        ):
            _add(
                "FAIL",
                "BULK_DENSITY_HARD",
                "Bulk density is outside physical bounds.",
                bulk_density_g_ml,
            )
        elif bulk_density_g_ml < limits.get("bulk_density_warn_low_g_ml", 1.8) or bulk_density_g_ml > limits.get(
            "bulk_density_warn_high_g_ml", 3.0
        ):
            _add(
                "WARN",
                "BULK_DENSITY_WARN",
                "Bulk density is suspicious (too low/high).",
                bulk_density_g_ml,
            )

    # Skeletal density (grain density proxy)
    if skeletal_density_g_ml is None:
        _add("WARN", "MISSING_SKELETAL_DENSITY", "Apparent/Skeletal density not found in metadata.")
    else:
        if skeletal_density_g_ml < limits.get("skeletal_density_hard_min_g_ml", 1.0) or skeletal_density_g_ml > limits.get(
            "skeletal_density_hard_max_g_ml", 6.0
        ):
            _add(
                "FAIL",
                "SKELETAL_DENSITY_HARD",
                "Skeletal density is outside physical bounds.",
                skeletal_density_g_ml,
            )
        elif skeletal_density_g_ml < limits.get("skeletal_density_warn_low_g_ml", 2.5) or skeletal_density_g_ml > limits.get(
            "skeletal_density_warn_high_g_ml", 2.95
        ):
            _add(
                "WARN",
                "SKELETAL_DENSITY_WARN",
                "Skeletal density suggests unusual mineralogy or possible measurement issues.",
                skeletal_density_g_ml,
            )

    # Permeability
    if k_air_md is None:
        _add("WARN", "MISSING_PERMEABILITY", "Permeability not found in metadata.")
    else:
        if k_air_md < limits.get("permeability_hard_min_md", 0.0):
            _add("FAIL", "PERMEABILITY_HARD", "Permeability cannot be negative.", k_air_md)
        elif k_air_md > limits.get("permeability_warn_high_md", 10000.0):
            _add(
                "WARN",
                "PERMEABILITY_WARN",
                "Very high permeability; may indicate fractures or a broken sample.",
                k_air_md,
            )

    # Stem volume used
    if stem_used_pct is None:
        _add("WARN", "MISSING_STEM_USED", "Stem volume used (%) not found in metadata.")
    else:
        if stem_used_pct < 0:
            _add("FAIL", "STEM_USED_HARD", "Stem volume used cannot be negative.", stem_used_pct)
        elif stem_used_pct > limits.get("stem_used_hard_max_pct", 100.0):
            _add(
                "FAIL",
                "STEM_USED_HARD",
                "Stem volume used exceeds 100% (likely leak / wrong penetrometer).",
                stem_used_pct,
            )
        elif stem_used_pct > limits.get("stem_used_warn_high_pct", 85.0):
            _add(
                "WARN",
                "STEM_USED_WARN",
                "Stem volume used is high; risk of running out of mercury at high pressure.",
                stem_used_pct,
            )

    # Max Hg saturation
    if max_shg_sat_pct is None:
        _add(
            "WARN",
            "MISSING_MAX_SAT",
            "Max Hg saturation could not be computed (needs porosity, bulk density, intrusion volume).",
        )
    else:
        if max_shg_sat_pct < limits.get("max_shg_sat_hard_min_pct", 10.0):
            _add(
                "FAIL",
                "MAX_SAT_HARD",
                "Max Hg saturation is too low; data likely unusable for Sw_irreducible/SHF.",
                max_shg_sat_pct,
            )
        elif max_shg_sat_pct < limits.get("max_shg_sat_warn_min_pct", 90.0):
            _add(
                "WARN",
                "MAX_SAT_WARN",
                "Max Hg saturation is below recommended level (<90%).",
                max_shg_sat_pct,
            )


    # Plateau / fully-filled check (end-slope of intrusion curve)
    plateau_slope_ml_g_psia = None
    plateau_fill_flag: Optional[str] = None
    if dfx is not None:
        plateau_slope_ml_g_psia, _pl_info = compute_plateau_tail_slope(dfx, params)

    if plateau_slope_ml_g_psia is not None and np.isfinite(plateau_slope_ml_g_psia):
        pass_th = float(params.get("plateau_slope_pass", 1.0e-4))
        warn_th = float(params.get("plateau_slope_warn", 1.0e-3))
        if plateau_slope_ml_g_psia < pass_th:
            plateau_fill_flag = "PASS"
        elif plateau_slope_ml_g_psia > warn_th:
            plateau_fill_flag = "WARN"
            _add(
                "WARN",
                "PLATEAU_WARN",
                "End-curve not at plateau (Hg still intruding at max pressure).",
                plateau_slope_ml_g_psia,
            )
        else:
            plateau_fill_flag = "INFO"


    # Threshold pressure
    if threshold_psia is None:
        _add("WARN", "MISSING_THRESHOLD_P", "Threshold pressure not found in metadata.")
    else:
        if threshold_psia < limits.get("threshold_pressure_hard_min_psia", 0.0):
            _add("FAIL", "THRESHOLD_P_HARD", "Threshold pressure cannot be negative.", threshold_psia)
        elif threshold_psia < limits.get("threshold_pressure_warn_low_psia", 0.5):
            _add(
                "WARN",
                "THRESHOLD_P_WARN",
                "Very low threshold pressure; may be conforming fill (roughness/fractures), not true pores.",
                threshold_psia,
            )

    # Tortuosity
    if tortuosity is None:
        _add("WARN", "MISSING_TORTUOSITY", "Tortuosity not found in metadata.")
    else:
        if tortuosity < limits.get("tortuosity_hard_min", 1.0):
            _add("FAIL", "TORTUOSITY_HARD", "Tortuosity cannot be < 1.", tortuosity)
        elif tortuosity > limits.get("tortuosity_warn_high", 10.0):
            _add("WARN", "TORTUOSITY_WARN", "Tortuosity is very high.", tortuosity)

    # Tortuosity factor (Archie a proxy)
    if tort_factor is None:
        _add("WARN", "MISSING_TORT_FACTOR", "Tortuosity factor (Archie a) not found in metadata.")
    else:
        if tort_factor < limits.get("tortuosity_factor_hard_min", 0.1) or tort_factor > limits.get(
            "tortuosity_factor_hard_max", 5.0
        ):
            _add("FAIL", "TORT_FACTOR_HARD", "Tortuosity factor is outside physical bounds.", tort_factor)
        elif tort_factor < limits.get("tortuosity_factor_warn_low", 0.6) or tort_factor > limits.get(
            "tortuosity_factor_warn_high", 2.0
        ):
            _add("WARN", "TORT_FACTOR_WARN", "Tortuosity factor is unusual for Archie a.", tort_factor)

    # Formation factor
    if formation_factor is not None:
        if formation_factor < limits.get("formation_factor_hard_min", 1.0):
            _add("FAIL", "FORMATION_FACTOR_HARD", "Formation factor must be >= 1.", formation_factor)

    # Extrusion efficiency
    if extrusion_eff_pct is not None:
        if extrusion_eff_pct < limits.get("extrusion_eff_hard_min_pct", 0.0) or extrusion_eff_pct > limits.get(
            "extrusion_eff_hard_max_pct", 100.0
        ):
            _add(
                "FAIL",
                "EXTRUSION_EFF_HARD",
                "Extrusion efficiency is outside 0–100%.",
                extrusion_eff_pct,
            )
        elif extrusion_eff_pct < limits.get("extrusion_eff_warn_low_pct", 10.0) or extrusion_eff_pct > limits.get(
            "extrusion_eff_warn_high_pct", 80.0
        ):
            _add(
                "WARN",
                "EXTRUSION_EFF_WARN",
                "Extrusion efficiency is suspicious (<10% or >80%).",
                extrusion_eff_pct,
            )

    # Logic check: Bulk density should be < skeletal/grain density
    if bulk_density_g_ml is not None and skeletal_density_g_ml is not None:
        if bulk_density_g_ml >= skeletal_density_g_ml:
            _add(
                "FAIL",
                "DENSITY_LOGIC_HARD",
                "Bulk density cannot be >= grain (skeletal) density (implies negative porosity).",
                {"rho_bulk": bulk_density_g_ml, "rho_grain": skeletal_density_g_ml},
            )
        else:
            warn_diff = limits.get("grain_bulk_warn_diff_g_ml", 0.1)
            if grain_density_diff is not None and grain_density_diff < warn_diff:
                _add(
                    "WARN",
                    "DENSITY_LOGIC_WARN",
                    "Very small (rho_grain - rho_bulk); porosity may be near-zero or measurement uncertain.",
                    grain_density_diff,
                )

    # Logical check: Swanson vs k_air discrepancy
    if k_ratio is not None:
        if k_ratio < limits.get("k_discrep_ratio_warn_low", 0.2) or k_ratio > limits.get("k_discrep_ratio_warn_high", 5.0):
            _add(
                "WARN",
                "K_DISCREP_WARN",
                "k_swanson / k_air discrepancy is large; Swanson model may not be calibrated for this rock.",
                k_ratio,
            )

    # Logical check: cementation exponent m (if computed)
    if m_cementation is not None and np.isfinite(m_cementation):
        if m_cementation < limits.get("m_cementation_warn_low", 1.1) or m_cementation > limits.get(
            "m_cementation_warn_high", 3.0
        ):
            _add(
                "WARN",
                "ARCHIE_M_WARN",
                "Cementation exponent m is outside typical range (1.1–3.0).",
                m_cementation,
            )

    # --------------------------
    # Grade + recommendation
    # --------------------------
    n_fail = sum(1 for it in issues if it.get("level") == "FAIL")
    n_warn = sum(1 for it in issues if it.get("level") == "WARN")

    if n_fail > 0:
        grade = "FAIL"
        rec = "DISCARD"
    elif n_warn > 0:
        grade = "WARN"
        rec = "REVIEW"
    else:
        grade = "PASS"
        rec = "ACCEPT"

    petro_qc_done = True if (meta or df is not None) else False

    out = {
        "petro_qc_done": petro_qc_done,
        "petro_qc_grade": grade,
        "petro_qc_recommendation": rec,
        "petro_qc_issues": issues,
        "petro_qc_n_fail": n_fail,
        "petro_qc_n_warn": n_warn,
        # Derived
        "phi_pct_used": porosity_pct,
        "k_air_md_used": k_air_md,
        "threshold_pressure_psia_used": threshold_psia,
        "threshold_pressure_method": threshold_method,
        "threshold_pressure_detail": threshold_detail,
        "max_shg_sat_pct": max_shg_sat_pct,
        "plateau_slope_ml_g_psia": plateau_slope_ml_g_psia,
        "plateau_fill_flag": plateau_fill_flag,
        "backbone_pressure_psia": backbone_psia,
        "backbone_sat_eff_frac": (backbone_info.get("sat_eff_frac") if isinstance(backbone_info, dict) else None),
        "grain_density_diff_g_ml": grain_density_diff,
        "conf_pknee_psia": conf_pknee_psia,
        "conf_vol_ml_g": conf_vol_ml_g,
        "conformance_vol_pct": conformance_vol_pct,
        "k_swanson_md": k_swanson_md,
        "k_swanson_kair_ratio": k_ratio,
        "m_cementation": m_cementation,
    }
    return out


def petroqc_issues_to_text(issues: List[Dict[str, Any]]) -> str:
    """Single-line reasons string for the PetroQC table."""
    if not issues:
        return ""
    parts: List[str] = []
    for it in issues:
        lvl = str(it.get("level", "")).strip()
        code = str(it.get("code", "")).strip()
        msg = str(it.get("message", "")).strip()
        val = it.get("value", None)
        if val is not None:
            parts.append(f"{lvl}: {code} - {msg} [value={val}]")
        else:
            parts.append(f"{lvl}: {code} - {msg}")
    return " | ".join(parts)


def build_petroqc_row(sample: Dict[str, Any]) -> Dict[str, Any]:
    """Build the row dict for the PetroQC multi-sample grid."""
    meta = sample.get("meta", {}) or {}
    res = sample.get("results", {}) or {}

    well = sample.get("well") or meta.get("well_guess") or meta.get("well") or ""
    core = meta.get("core_guess") or meta.get("core") or ""
    sample_id = meta.get("sample_id") or sample.get("filename") or ""
    depth_m = meta.get("depth_m")

    # Prefer values used by PetroQC (overrides respected)
    porosity_pct = res.get("phi_pct_used", meta.get("porosity_pct"))
    perm_md = res.get("k_air_md_used", meta.get("permeability_md"))

    threshold_psia = res.get("threshold_pressure_psia_used", meta.get("threshold_pressure_psia"))
    tort_factor = meta.get("tortuosity_factor")

    rho_bulk = meta.get("bulk_density_g_ml")
    rho_grain = meta.get("skeletal_density_g_ml")

    stem_used = meta.get("stem_volume_used_pct")

    max_sat = res.get("max_shg_sat_pct")
    rho_diff = res.get("grain_density_diff_g_ml")
    conf_pct = res.get("conformance_vol_pct")

    grade = res.get("petro_qc_grade", "")
    reasons = petroqc_issues_to_text(res.get("petro_qc_issues", []))

    decision = str(res.get("sample_decision", "PENDING")).upper()
    exclude_from_shm = bool(decision == "DISCARDED" or str(grade).upper() == "FAIL")

    return {
        "id": sample.get("id"),
        "well": well,
        "core": core,
        "sample_id": sample_id,
        "depth_m": depth_m,
        "porosity_pct": porosity_pct,
        "permeability_md": perm_md,
        "k_pnm_md": res.get("k_pnm_md"),
        "pnm_ci": res.get("pnm_ci"),
        "thomeer_mode": res.get("thomeer_mode"),
        "thomeer_pd": _safe_float(res.get("thomeer_pd_psia")),
        "thomeer_g": _safe_float(res.get("thomeer_G")),
        "thomeer_bv": _safe_float(res.get("thomeer_bvinf_pct")),
        "thomeer_pd1": _safe_float(res.get("thomeer_pd1_psia")),
        "thomeer_g1": _safe_float(res.get("thomeer_G1")),
        "thomeer_bv1": _safe_float(res.get("thomeer_bvinf1_pct")),
        "thomeer_pd2": _safe_float(res.get("thomeer_pd2_psia")),
        "thomeer_g2": _safe_float(res.get("thomeer_G2")),
        "thomeer_bv2": _safe_float(res.get("thomeer_bvinf2_pct")),
        "thomeer_macro_frac": _safe_float(res.get("thomeer_macro_frac")),
        "thomeer_bimodal_qc": res.get("thomeer_bimodal_qc", ""),
        "thomeer_pd_sep_log10": res.get("thomeer_pd_sep_log10", np.nan),
        "k_winland_md_total": _safe_float(res.get("k_winland_md_total")),
        "k_winland_md_macro": _safe_float(res.get("k_winland_md_macro")),
        "threshold_pressure_psia": threshold_psia,
        "tortuosity_factor": tort_factor,
        "bulk_density_g_ml": rho_bulk,
        "skeletal_density_g_ml": rho_grain,
        "stem_volume_used_pct": stem_used,
        "max_shg_sat_pct": max_sat,
        "grain_density_diff_g_ml": rho_diff,
        "conformance_vol_pct": conf_pct,
        "qc_flag": grade,
        "qc_reasons": reasons,
        "exclude_from_shm": exclude_from_shm,
    }


def run_petroqc_on_library(
    library: List[Dict[str, Any]],
    params: Dict[str, Any],
) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]], Dict[str, int]]:
    """Run PetroQC on all samples in the library and return (updated_library, rows, counts)."""
    library = library or []
    params = params or DEFAULT_PARAMS

    updated: List[Dict[str, Any]] = []
    rows: List[Dict[str, Any]] = []
    counts = {"PASS": 0, "WARN": 0, "FAIL": 0}

    for s in library:
        meta = s.get("meta", {}) or {}
        res0 = dict(s.get("results", {}) or {})

        df = None
        try:
            df = _ensure_schema(pd.DataFrame(s.get("data", [])))
        except Exception:
            df = None

        qc = petrophysical_qaqc(meta, params, df=df, res=res0)
        res0.update(qc)

        grade = str(res0.get("petro_qc_grade", "")).upper()
        if grade in counts:
            counts[grade] += 1

        # Exclude flag used by SHM/selection logic
        decision = str(res0.get("sample_decision", "PENDING")).upper()
        res0["exclude_from_shm"] = bool(decision == "DISCARDED" or grade == "FAIL")

        s2 = dict(s)
        s2["results"] = res0
        updated.append(s2)

        rows.append(build_petroqc_row(s2))

    return updated, rows, counts
def apply_qaqc_flags(df: pd.DataFrame) -> pd.DataFrame:
    df = _ensure_schema(df.copy())

    p = df["Pressure"]
    c = df["CumVol"]
    i = df["IncVol"]

    p_prev = p.shift(1)
    c_prev = c.shift(1)

    df["Flag_Pressure_Down"] = (p.notna() & p_prev.notna() & (p < p_prev)).map(lambda x: "Y" if x else "")
    df["Flag_Cum_Down"] = (c.notna() & c_prev.notna() & (c < c_prev)).map(lambda x: "Y" if x else "")
    df["Flag_Inc_Neg_Fail"] = (i.notna() & (i < 0)).map(lambda x: "Y" if x else "")

    return df

def scrub_data(df: pd.DataFrame) -> pd.DataFrame:
    df = _ensure_schema(df.copy())
    df = df.loc[df["Pressure"].notna()].copy()

    df = df.sort_values("Pressure", ascending=True).reset_index(drop=True)

    # Clip CumVol
    df["CumVol"] = df["CumVol"].clip(lower=0)

    # Enforce monotonic cumulative (non-decreasing)
    df["CumVol"] = df["CumVol"].cummax()

    # Recompute incremental from cumulative
    df["IncVol"] = df["CumVol"].diff()
    df.loc[df.index[0], "IncVol"] = df.loc[df.index[0], "CumVol"]
    df["IncVol"] = df["IncVol"].fillna(0).clip(lower=0)

    # Clear flags
    for c in FLAG_COLS:
        df[c] = ""

    return df





def _pressure_at_cumvol_target_psia(
    pressures_psia: np.ndarray,
    cumvol_ml_g: np.ndarray,
    v_target_ml_g: float,
) -> Optional[float]:
    """
    Helper: given monotonic cumulative intrusion vs pressure, return pressure at a target volume.
    Uses linear interpolation in log10(P) (standard for MICP curves on log-pressure axes).
    """
    try:
        p = np.asarray(pressures_psia, dtype=float)
        v = np.asarray(cumvol_ml_g, dtype=float)
        if p.size < 2:
            return None
        # Ensure valid
        m = np.isfinite(p) & np.isfinite(v) & (p > 0)
        p = p[m]
        v = v[m]
        if p.size < 2:
            return None

        # Sort by pressure
        order = np.argsort(p)
        p = p[order]
        v = v[order]

        # Force monotonic non-decreasing cumulative volume
        v = np.maximum.accumulate(np.clip(v, 0.0, None))

        vmax = float(np.nanmax(v)) if v.size else 0.0
        if not np.isfinite(vmax) or vmax <= 0:
            return None

        v_target = float(v_target_ml_g)
        if not np.isfinite(v_target):
            return None
        v_target = max(0.0, min(v_target, vmax))

        # Find first index where v >= v_target
        i = int(np.searchsorted(v, v_target, side="left"))
        if i <= 0:
            return float(p[0])
        if i >= p.size:
            return float(p[-1])

        p0, p1 = float(p[i - 1]), float(p[i])
        v0, v1 = float(v[i - 1]), float(v[i])
        if not (np.isfinite(p0) and np.isfinite(p1) and p0 > 0 and p1 > 0):
            return float(p[i])

        if v1 <= v0 or not np.isfinite(v0) or not np.isfinite(v1):
            return float(p[i])

        t = (v_target - v0) / (v1 - v0)
        t = float(max(0.0, min(1.0, t)))
        logp = math.log10(p0) + t * (math.log10(p1) - math.log10(p0))
        return float(10 ** logp)
    except Exception:
        return None


def compute_threshold_pressure_psia(
    df: pd.DataFrame,
    params: Dict[str, Any],
    frac_vmax: Optional[float] = None,
) -> Tuple[Optional[float], Dict[str, Any]]:
    """
    Compute Threshold Pressure (Pth / entry pressure) from an intrusion curve.

    Default method (recommended here): "acceleration" / ΔS detection:
      - Compute Hg saturation S (%) from cumulative intrusion.
      - Compute discrete derivative ΔS between consecutive points.
      - Pick the first pressure where ΔS exceeds a threshold (e.g., 1.0%),
        with an optional validation that the next increment is still positive
        (helps reject single-point noise).

    Fallback (legacy): volume-fraction-of-Vmax method (e.g., 2% of Vmax).

    Notes:
      - If your dataframe is already conformance-corrected (CumVol cleaned),
        this computation naturally works on the effective saturation.
      - Returned pressure is in psia.
    """
    method = str(params.get("pth_method", "ds_accel")).strip().lower()
    ds_jump_pct = float(params.get("pth_ds_jump_pct", 1.0))
    ds_validate_pct = float(params.get("pth_ds_validate_pct", 0.10))

    # Legacy fraction method config
    frac_legacy = float(
        frac_vmax
        if frac_vmax is not None
        else params.get("pth_frac_vmax", 0.02)
    )

    d = df.copy()
    pcol = "Pressure"
    vcol = "CumVol"
    if pcol not in d.columns or vcol not in d.columns:
        return None, {"method": "none", "detail": "missing columns", "pcol": pcol, "vcol": vcol}

    d = d[[pcol, vcol]].dropna()
    if d.empty:
        return None, {"method": "none", "detail": "empty dataframe"}

    # Clean arrays
    p = d[pcol].astype(float).to_numpy()
    v = d[vcol].astype(float).to_numpy()
    m = np.isfinite(p) & np.isfinite(v) & (p > 0)
    p = p[m]
    v = v[m]
    if p.size < 3:
        return None, {"method": "none", "detail": "insufficient points"}

    order = np.argsort(p)
    p = p[order]
    v = v[order]
    v = np.maximum.accumulate(np.clip(v, 0.0, None))
    vmax = float(np.nanmax(v)) if v.size else 0.0
    if not np.isfinite(vmax) or vmax <= 0:
        return None, {"method": "none", "detail": "vmax<=0"}

    sat_pct = (v / vmax) * 100.0

    # --- Method A: ΔS acceleration ---
    if method in {"ds_accel", "accel", "ds", "acceleration"}:
        ds = np.diff(sat_pct)
        # ds[i-1] corresponds to step from i-1 -> i
        for i in range(1, p.size):
            ds_i = float(ds[i - 1]) if (i - 1) < ds.size else 0.0
            if not np.isfinite(ds_i):
                continue
            if ds_i > ds_jump_pct:
                # Optional: validate next increment is still rising (avoid one-point spikes)
                if i < (p.size - 1):
                    ds_next = float(sat_pct[i + 1] - sat_pct[i])
                    if np.isfinite(ds_next) and ds_next < ds_validate_pct:
                        continue
                return float(p[i]), {
                    "method": "ds_accel",
                    "detail": f"first ΔS>{ds_jump_pct:.3g}%",
                    "ds_jump_pct": ds_jump_pct,
                    "ds_validate_pct": ds_validate_pct,
                    "i": int(i),
                    "sat_pct_at_pth": float(sat_pct[i]),
                    "vmax_ml_g": vmax,
                }

        # If not found, fall back to legacy fraction-of-Vmax
        method = "vfrac_vmax_fallback"

    # --- Method B: legacy volume fraction of Vmax ---
    frac = float(max(0.0, min(1.0, frac_legacy)))
    vth = frac * vmax
    pth = _pressure_at_cumvol_target_psia(p, v, vth)
    return pth, {
        "method": "vfrac_vmax" if method != "vfrac_vmax_fallback" else "vfrac_vmax_fallback",
        "detail": "interp_logP @ v_target",
        "vmax_ml_g": vmax,
        "vth_ml_g": vth,
        "frac": frac,
    }


def compute_backbone_pressure_psia(
    df: pd.DataFrame,
    params: Dict[str, Any],
    sat_eff_frac: Optional[float] = None,
) -> Tuple[Optional[float], Dict[str, Any]]:
    """
    Proxy for 'backbone / fractal region' onset: pressure at a given effective saturation
    fraction (default ~45%) after conformance cleaning.

    This is a pragmatic (and fast) implementation:
      P_backbone = P where CumVol reaches sat_eff_frac * Vmax (after any conformance correction).
    """
    frac = float(
        sat_eff_frac if sat_eff_frac is not None else params.get("backbone_sat_eff_frac", 0.45)
    )
    frac = float(max(0.0, min(1.0, frac)))

    d = df.copy()
    if "Pressure" not in d.columns or "CumVol" not in d.columns:
        return None, {"method": "none", "detail": "missing columns"}

    d = d[["Pressure", "CumVol"]].dropna()
    if d.empty:
        return None, {"method": "none", "detail": "empty dataframe"}

    p = d["Pressure"].astype(float).to_numpy()
    v = d["CumVol"].astype(float).to_numpy()
    m = np.isfinite(p) & np.isfinite(v) & (p > 0)
    p = p[m]
    v = v[m]
    if p.size < 2:
        return None, {"method": "none", "detail": "insufficient points"}

    order = np.argsort(p)
    p = p[order]
    v = v[order]
    v = np.maximum.accumulate(np.clip(v, 0.0, None))
    vmax = float(np.nanmax(v)) if v.size else 0.0
    if not np.isfinite(vmax) or vmax <= 0:
        return None, {"method": "none", "detail": "vmax<=0"}

    v_target = frac * vmax
    pbb = _pressure_at_cumvol_target_psia(p, v, v_target)
    return pbb, {
        "method": "sat_eff_frac",
        "sat_eff_frac": frac,
        "vmax_ml_g": vmax,
        "v_target_ml_g": v_target,
    }


def compute_plateau_tail_slope(
    df: pd.DataFrame,
    params: Dict[str, Any],
) -> Tuple[Optional[float], Dict[str, Any]]:
    """
    Compute the end-slope of cumulative intrusion vs pressure in the last N points.
    Used as a 'did we reach plateau / fully filled sample?' check.

    Slope units: (mL/g) / psia.
    """
    n = int(params.get("plateau_tail_n", 5))
    n = max(2, n)

    d = df.copy()
    if "Pressure" not in d.columns or "CumVol" not in d.columns:
        return None, {"method": "none", "detail": "missing columns"}

    d = d[["Pressure", "CumVol"]].dropna()
    if d.empty:
        return None, {"method": "none", "detail": "empty dataframe"}

    p = d["Pressure"].astype(float).to_numpy()
    v = d["CumVol"].astype(float).to_numpy()
    m = np.isfinite(p) & np.isfinite(v) & (p > 0)
    p = p[m]
    v = v[m]
    if p.size < n:
        return None, {"method": "none", "detail": "insufficient points", "n": n}

    order = np.argsort(p)
    p = p[order]
    v = v[order]
    v = np.maximum.accumulate(np.clip(v, 0.0, None))

    p_tail = p[-n:]
    v_tail = v[-n:]
    dp = float(p_tail[-1] - p_tail[0])
    if not np.isfinite(dp) or dp <= 0:
        return None, {"method": "none", "detail": "dp<=0"}

    dv = float(v_tail[-1] - v_tail[0])
    slope = dv / dp
    return float(slope), {
        "method": "tail_slope",
        "n": n,
        "dp": dp,
        "dv": dv,
    }
def detect_conformance_knee(
    df: pd.DataFrame,
    params: Optional[Dict[str, Any]] = None,
) -> Tuple[Optional[float], Optional[float], str]:
    """
    Conformance / closure point used for *data cleaning* (NOT Pd).

    Default behavior ("auto"):
      1) Acceleration method (recommended):
         - Build saturation S (%) from cumulative intrusion (CumVol / Vmax * 100).
         - Compute discrete derivative ΔS between consecutive points.
         - Find the first step where ΔS exceeds the configured jump threshold
           (params["pth_ds_jump_pct"], default 1.0 percentage points),
           with an optional validation that the next increment is still positive
           (params["pth_ds_validate_pct"], default 0.10 percentage points).
         - Define Conformance as the point immediately *before* that jump:
             P_conf = P[i-1], V_conf = V[i-1]

         This matches the definition: "Conformance = all volume injected before Pth".

      2) Fallback legacy knee detector (only if the acceleration method fails):
         - Looks for an early low-pressure jump within a small saturation cap.

    Returns:
      (P_conf_psia, V_conf_ml_g, method_tag)
    """

    # Be defensive: some callbacks may call this helper without passing params.
    # Merge DEFAULT_PARAMS with any user overrides.
    if not isinstance(params, dict):
        params = {}
    _p = dict(DEFAULT_PARAMS)
    _p.update(params)
    params = _p

    conf_mode = str(params.get("conf_method", "auto")).strip().lower()

    # Prepare arrays
    try:
        d = df[["Pressure", "CumVol"]].copy().dropna()
    except Exception:
        return None, None, "none"

    if d.empty:
        return None, None, "none"

    p = d["Pressure"].astype(float).to_numpy()
    v = d["CumVol"].astype(float).to_numpy()

    m = np.isfinite(p) & np.isfinite(v) & (p > 0)
    p = p[m]
    v = v[m]
    if p.size < 3:
        return None, None, "none"

    order = np.argsort(p)
    p = p[order]
    v = v[order]
    v = np.maximum.accumulate(np.clip(v, 0.0, None))

    vmax = float(np.nanmax(v)) if v.size else 0.0
    if not np.isfinite(vmax) or vmax <= 0:
        return None, None, "none"

    sat_pct = (v / vmax) * 100.0

    # ---------------------------------------------------------
    # A) Acceleration method (preferred)
    # ---------------------------------------------------------
    if conf_mode in {"auto", "pre_pth", "pth", "ds_accel", "accel", "acceleration"}:
        ds_jump_pct = float(params.get("pth_ds_jump_pct", 1.0))
        ds_validate_pct = float(params.get("pth_ds_validate_pct", 0.10))
        ds = np.diff(sat_pct)  # ds[i-1] = S[i]-S[i-1]

        for i in range(1, p.size):
            ds_i = float(ds[i - 1]) if (i - 1) < ds.size else 0.0
            if not np.isfinite(ds_i):
                continue
            if ds_i > ds_jump_pct:
                # Validate next step still rises (reject one-point noise)
                if i < (p.size - 1):
                    ds_next = float(sat_pct[i + 1] - sat_pct[i])
                    if np.isfinite(ds_next) and ds_next < ds_validate_pct:
                        continue

                # Conformance = previous point
                j = max(0, i - 1)
                return float(p[j]), float(v[j]), "pre_pth"

        if conf_mode in {"pre_pth", "pth"}:
            return None, None, "none"

    # ---------------------------------------------------------
    # B) Legacy fallback (kept for robustness)
    # ---------------------------------------------------------
    try:
        cap_frac = float(params.get("conf_max_sat_frac", 0.010))
        jump_frac = float(params.get("conf_jump_sat_frac", 0.005))
        cap_sat_pct = cap_frac * 100.0
        jump_sat_pct = jump_frac * 100.0

        prev = float(sat_pct[0])
        for i in range(1, sat_pct.size):
            s = float(sat_pct[i])
            ds_i = s - prev
            if (s <= cap_sat_pct) and (ds_i > jump_sat_pct):
                return float(p[i]), float(v[i]), "legacy_jump"
            prev = s

        # Fallback: last point within the cap (or first point)
        idx_cap = int(np.searchsorted(sat_pct, cap_sat_pct, side="right") - 1)
        idx_cap = max(0, min(idx_cap, p.size - 1))
        return float(p[idx_cap]), float(v[idx_cap]), "legacy_cap"
    except Exception:
        return None, None, "none"


def apply_conformance_correction(df: pd.DataFrame, pknee_psia: float, v_conf_ml_g: float) -> pd.DataFrame:
    """
    Simple conformance correction:
      - subtract constant V_conf from CumVol,
      - clip to >=0 and enforce monotonic,
      - recompute IncVol from corrected CumVol.

    NOTE: This is an approximation. Always review visually; in some rocks, early intrusion is real macroporosity.
    """
    d = _ensure_schema(df.copy())
    d = d.loc[d["Pressure"].notna()].copy()
    d = d.sort_values("Pressure", ascending=True).reset_index(drop=True)

    try:
        v0 = float(v_conf_ml_g)
    except Exception:
        return d

    if not (np.isfinite(v0) and v0 >= 0):
        return d

    d["CumVol"] = pd.to_numeric(d["CumVol"], errors="coerce") - v0
    d["CumVol"] = d["CumVol"].fillna(0).clip(lower=0)

    # Keep cumulative monotonic after shifting
    d["CumVol"] = d["CumVol"].cummax()

    # Recompute incremental
    d["IncVol"] = d["CumVol"].diff()
    if len(d) > 0:
        d.loc[d.index[0], "IncVol"] = d.loc[d.index[0], "CumVol"]
    d["IncVol"] = d["IncVol"].fillna(0).clip(lower=0)

    # Clear QA/QC flags (user can rerun QA/QC)
    for c in FLAG_COLS:
        d[c] = ""

    return d


def cumvol_at_pressure(df: pd.DataFrame, pc_psia: float) -> Optional[float]:
    """Interpolate CumVol at a given pressure (psia), using log10(P) interpolation."""
    try:
        pk = float(pc_psia)
    except Exception:
        return None
    if not np.isfinite(pk) or pk <= 0:
        return None

    d = _ensure_schema(df.copy())
    d = d.loc[d["Pressure"].notna()].copy()
    d = d.sort_values("Pressure", ascending=True)

    # Keep only valid positive pressures and cumulative values
    p = d["Pressure"].to_numpy(dtype=float)
    c = d["CumVol"].to_numpy(dtype=float)
    m = np.isfinite(p) & np.isfinite(c) & (p > 0)
    if m.sum() < 2:
        return None

    p = p[m]
    c = c[m]

    # De-duplicate pressures (take max CumVol at each P)
    dd = pd.DataFrame({"P": p, "C": c}).groupby("P", as_index=False)["C"].max()
    p_u = dd["P"].to_numpy(dtype=float)
    c_u = dd["C"].to_numpy(dtype=float)
    if len(p_u) < 2:
        return float(c_u[0]) if len(c_u) == 1 else None

    pmin = float(np.nanmin(p_u))
    pmax = float(np.nanmax(p_u))
    pk = min(max(pk, pmin), pmax)

    logp = np.log10(p_u)
    logk = math.log10(pk)
    v = float(np.interp(logk, logp, c_u))
    return v

def _compute_radius_um(pc_psia: np.ndarray, sigma_npm: float, theta_deg: float) -> np.ndarray:
    pc_psia = np.asarray(pc_psia, dtype=float)
    pc_pa = pc_psia * 6894.757293168  # Pa
    theta_rad = np.deg2rad(theta_deg)
    r_m = -(2.0 * sigma_npm * np.cos(theta_rad)) / (pc_pa + 1e-30)
    r_um = r_m * 1e6
    return r_um

def _compute_saturation(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    cum = df["CumVol"].to_numpy(dtype=float)
    cum_max = np.nanmax(cum) if np.isfinite(cum).any() else np.nan
    if not np.isfinite(cum_max) or cum_max <= 0:
        df["HgSat"] = np.nan
        df["Sw"] = np.nan
        return df
    hgsat = cum / cum_max
    df["HgSat"] = np.clip(hgsat, 0, 1)
    df["Sw"] = 1.0 - df["HgSat"]
    return df

def _compute_psd(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    r = df["r_um"].to_numpy(dtype=float)
    r = np.clip(r, 1e-12, None)
    logr = np.log10(r)

    dlog = np.abs(np.diff(logr, prepend=logr[0]))
    inc = df["IncVol"].to_numpy(dtype=float)
    inc = np.clip(inc, 0, None)  # for PSD we use non-negative increments

    dVdlogr = np.divide(inc, dlog, out=np.full_like(inc, np.nan, dtype=float), where=dlog > 0)

    total = np.nanmax(df["CumVol"].to_numpy(dtype=float))
    if not np.isfinite(total) or total <= 0:
        df["dlogr"] = dlog
        df["dVdlogr"] = dVdlogr
        df["dVdlogr_norm"] = np.nan
        return df

    df["dlogr"] = dlog
    df["dVdlogr"] = dVdlogr
    df["dVdlogr_norm"] = dVdlogr / total
    return df

def recompute_derived(df: pd.DataFrame, params: Dict[str, Any]) -> pd.DataFrame:
    df = _ensure_schema(df.copy())
    df = _compute_saturation(df)

    sigma = float(params.get("sigma_hg_air_npm", DEFAULT_PARAMS["sigma_hg_air_npm"]))
    theta = float(params.get("theta_hg_air_deg", DEFAULT_PARAMS["theta_hg_air_deg"]))
    pc = df["Pressure"].to_numpy(dtype=float)
    df["r_um"] = _compute_radius_um(pc, sigma, theta)
    df["d_um"] = df["r_um"] * 2.0

    df = _compute_psd(df)

    # SHF columns (computed when reservoir params available)
    df = compute_shf(df, params)

    return df

def _interp_pc_at_fraction(df: pd.DataFrame, frac_target: float) -> Optional[float]:
    """Interpolate Pc (psia) at given HgSat fraction target."""
    df = df.copy()
    df = df.loc[df["Pressure"].notna() & df["CumVol"].notna()].copy()
    if df.empty:
        return None

    # Ensure saturation exists
    df = _compute_saturation(df)
    s = df["HgSat"].to_numpy(dtype=float)
    pc = df["Pressure"].to_numpy(dtype=float)

    m = np.isfinite(s) & np.isfinite(pc) & (pc > 0)
    s = s[m]
    pc = pc[m]
    if len(s) < 2:
        return None

    # Sort by saturation increasing for interpolation stability
    order = np.argsort(s)
    s = s[order]
    pc = pc[order]

    if frac_target <= s.min():
        return float(pc[np.argmin(s)])
    if frac_target >= s.max():
        return float(pc[np.argmax(s)])

    idx = np.searchsorted(s, frac_target, side="left")
    if idx <= 0 or idx >= len(s):
        return None

    s0, s1 = s[idx - 1], s[idx]
    pc0, pc1 = pc[idx - 1], pc[idx]

    # interpolate in log10 Pc (more typical for capillary pressure curves)
    logpc0, logpc1 = np.log10(pc0), np.log10(pc1)
    t = (frac_target - s0) / (s1 - s0 + 1e-30)
    logpc = logpc0 + t * (logpc1 - logpc0)
    return float(10 ** logpc)


def _collapse_depth_duplicates(depth: np.ndarray, x: np.ndarray, ndigits: int = 2, agg: str = "median"):
    """Collapse duplicate (or near-duplicate) depth samples into a single representative value.

    Why:
      Core plug datasets often contain multiple plugs at the *same* depth. If we connect
      those points as a line, Plotly draws misleading horizontal segments.

    Strategy:
      - Round depth to `ndigits` (meters) and aggregate `x` per rounded depth.
      - Default aggregation is median (robust to outliers).
    """
    if depth is None or x is None:
        return np.array([]), np.array([])

    d = np.asarray(depth, dtype=float)
    v = np.asarray(x, dtype=float)
    m = np.isfinite(d) & np.isfinite(v)
    d = d[m]
    v = v[m]
    if len(d) == 0:
        return np.array([]), np.array([])

    df = pd.DataFrame({"depth": d, "x": v})
    df["depth_bin"] = df["depth"].round(int(ndigits))

    if agg == "mean":
        g = df.groupby("depth_bin", as_index=False)["x"].mean()
    elif agg == "min":
        g = df.groupby("depth_bin", as_index=False)["x"].min()
    elif agg == "max":
        g = df.groupby("depth_bin", as_index=False)["x"].max()
    else:
        g = df.groupby("depth_bin", as_index=False)["x"].median()

    g = g.sort_values("depth_bin")
    return g["depth_bin"].to_numpy(dtype=float), g["x"].to_numpy(dtype=float)


def _auto_gap_break_m(depth: np.ndarray, min_gap: float = 10.0, max_gap: float = 50.0) -> float:
    """Heuristic depth gap (m) above which we break line connections."""
    d = np.asarray(depth, dtype=float)
    d = d[np.isfinite(d)]
    if len(d) < 4:
        return float(min_gap)
    d = np.sort(d)
    diffs = np.diff(d)
    diffs = diffs[np.isfinite(diffs) & (diffs > 0)]
    if len(diffs) == 0:
        return float(min_gap)

    q75 = float(np.percentile(diffs, 75))
    gap = max(min_gap, 3.0 * q75)
    gap = min(max_gap, gap)
    return float(gap)


def _split_by_depth_gap(depth: np.ndarray, x: np.ndarray, gap_m: float):
    """Split a (depth, x) series into segments where depth gaps exceed `gap_m`."""
    d = np.asarray(depth, dtype=float)
    v = np.asarray(x, dtype=float)
    if len(d) == 0:
        return []
    if len(d) == 1:
        return [(d, v)]

    order = np.argsort(d)
    d = d[order]
    v = v[order]

    segments = []
    start = 0
    for i in range(1, len(d)):
        if (d[i] - d[i - 1]) > gap_m:
            segments.append((d[start:i], v[start:i]))
            start = i
    segments.append((d[start:], v[start:]))
    return segments


def _interp_logk(depth: np.ndarray, k: np.ndarray, step_m: Optional[float] = None, max_points: int = 800):
    """Interpolate permeability along depth in log10(k) space."""
    d = np.asarray(depth, dtype=float)
    kk = np.asarray(k, dtype=float)
    m = np.isfinite(d) & np.isfinite(kk) & (kk > 0)
    d = d[m]
    kk = kk[m]
    if len(d) < 2:
        return d, kk

    order = np.argsort(d)
    d = d[order]
    kk = kk[order]

    diffs = np.diff(d)
    diffs = diffs[np.isfinite(diffs) & (diffs > 0)]
    if step_m is None:
        step_m = float(max(0.25, (np.median(diffs) / 4.0) if len(diffs) else 0.25))

    dmin, dmax = float(d.min()), float(d.max())
    if dmax <= dmin:
        return d, kk

    n = int((dmax - dmin) / float(step_m)) + 1
    n = max(len(d), min(int(max_points), n))
    d_new = np.linspace(dmin, dmax, n)

    logk = np.log10(kk)
    logk_new = np.interp(d_new, d, logk)
    k_new = np.power(10.0, logk_new)
    return d_new, k_new

def compute_r_um_at_hgsat_fraction(df: pd.DataFrame, params: Dict[str, Any], frac_target: float) -> Optional[float]:
    """Compute pore throat *radius* (µm) at a target cumulative Hg saturation fraction.

    Notes:
    - HgSat is derived from cumulative intrusion volume (CumVol) normalized by its maximum.
    - We interpolate Pc at the target fraction in *log10(Pc)* space, then convert Pc → r using Washburn.
    """
    try:
        frac = float(frac_target)
    except Exception:
        return None

    pc = _interp_pc_at_fraction(df, frac)
    if pc is None or pc <= 0:
        return None

    sigma = float(params.get("sigma_hg_air_npm", DEFAULT_PARAMS["sigma_hg_air_npm"]))
    theta = float(params.get("theta_hg_air_deg", DEFAULT_PARAMS["theta_hg_air_deg"]))
    r = float(_compute_radius_um(np.array([pc], dtype=float), sigma, theta)[0])
    if not np.isfinite(r) or r <= 0:
        return None
    return r


def compute_r35_um(df: pd.DataFrame, params: Dict[str, Any]) -> Optional[float]:
    """Classic Winland radius r35 (µm): radius at HgSat = 0.35."""
    return compute_r_um_at_hgsat_fraction(df, params, 0.35)

def compute_swanson(df: pd.DataFrame, params: Dict[str, Any], meta: Dict[str, Any]) -> Dict[str, Any]:
    """
    Swanson apex:
      Sp = max( Sb / Pc ), where Sb is % bulk volume occupied by mercury (requires porosity).
    We compute Sb as: Sb_pct_bulk = HgSat_frac * phi_pct
    """
    out: Dict[str, Any] = {}

    # Porosity priority: override -> meta -> None
    phi_pct = params.get("phi_override_pct")
    if phi_pct is None:
        phi_pct = meta.get("porosity_pct")
    try:
        phi_pct_f = float(phi_pct)
    except Exception:
        phi_pct_f = float("nan")

    if not np.isfinite(phi_pct_f) or phi_pct_f <= 0:
        out["swanson_error"] = "Porosity missing (set phi_override_pct or ensure Porosity is parsed)."
        return out

    df2 = df.copy()
    df2 = _compute_saturation(df2)

    pc = df2["Pressure"].to_numpy(dtype=float)
    sb_pct_bulk = df2["HgSat"].to_numpy(dtype=float) * phi_pct_f  # % bulk
    ratio = np.where((pc > 0) & np.isfinite(pc) & np.isfinite(sb_pct_bulk), sb_pct_bulk / pc, np.nan)

    if not np.isfinite(ratio).any():
        out["swanson_error"] = "Cannot compute Swanson (no valid Pc or saturation)."
        return out

    idx = int(np.nanargmax(ratio))
    sp = float(ratio[idx])
    pc_apex = float(pc[idx]) if np.isfinite(pc[idx]) else None
    sb_apex = float(sb_pct_bulk[idx]) if np.isfinite(sb_pct_bulk[idx]) else None

    a = float(params.get("swanson_a", DEFAULT_PARAMS["swanson_a"]))
    b = float(params.get("swanson_b", DEFAULT_PARAMS["swanson_b"]))
    k_est = float(a * (sp ** b)) if np.isfinite(sp) and sp > 0 else None

    out.update({
        "phi_pct_used": phi_pct_f,
        "swanson_sp": sp,
        "swanson_pc_apex_psia": pc_apex,
        "swanson_sb_apex_pct_bulk": sb_apex,
        "k_swanson_md": k_est,
    })
    return out

def compute_winland_k(
    df: pd.DataFrame,
    params: Dict[str, Any],
    meta: Dict[str, Any],
    res: Optional[Dict[str, Any]] = None,
) -> Dict[str, Any]:
    """Winland r35 crossplot + permeability estimate.

    Classic Winland:
      - r35 is computed from MICP at HgSat = 0.35 (via interpolation).
      - If porosity is available, estimate k from r35:
          log10(r35) = 0.732 + 0.588 log10(k) - 0.864 log10(phi)

    Macro‑Normalized Winland (bimodal only):
      Bimodal rocks can have a large micro‑pore volume that biases r35 downward (and therefore k).
      When the current sample has a bi‑modal Thomeer fit (res['thomeer_mode']=='bimodal'),
      we *attempt* a macro‑normalized r35 using the macro fraction:

        HgSat_target_total = 0.35 * macro_frac

      Guardrails are applied to avoid "runaway" macro-normalized results when the macro volume
      is very small or when the computed r35_macro is unrealistically large.
    """
    out: Dict[str, Any] = {}
    params = params or DEFAULT_PARAMS
    res = res or {}

    # --- helpers
    def _k_from_r35_phi(r_um: float, phi_pct: float) -> Optional[float]:
        try:
            r_um_f = float(r_um)
            phi_pct_f = float(phi_pct)
        except Exception:
            return None
        if not (np.isfinite(r_um_f) and r_um_f > 0 and np.isfinite(phi_pct_f) and phi_pct_f > 0):
            return None
        # Invert: log r35 = 0.732 + 0.588 log k - 0.864 log phi
        logk = (math.log10(r_um_f) - 0.732 + 0.864 * math.log10(phi_pct_f)) / 0.588
        k = float(10 ** logk)
        if not np.isfinite(k) or k <= 0:
            return None
        return k

    # Porosity used by Winland (override > meta)
    phi_pct = params.get("phi_override_pct")
    if phi_pct is None:
        phi_pct = meta.get("porosity_pct")

    phi_pct_f: Optional[float] = None
    try:
        if phi_pct is not None:
            phi_pct_f = float(phi_pct)
            if not np.isfinite(phi_pct_f) or phi_pct_f <= 0:
                phi_pct_f = None
    except Exception:
        phi_pct_f = None

    if phi_pct_f is not None:
        out["phi_pct_used"] = float(phi_pct_f)

    # --- 1) Classic r35 (total)
    r35_total = compute_r_um_at_hgsat_fraction(df, params, 0.35)
    if r35_total is not None:
        out["r35_um_total"] = float(r35_total)
        out["r35_total_target_hgsat_frac"] = 0.35

        if phi_pct_f is not None:
            k_total = _k_from_r35_phi(r35_total, phi_pct_f)
            if k_total is not None:
                out["k_winland_md_total"] = float(k_total)

    # Default used values (classic)
    r35_used = r35_total
    k_used = out.get("k_winland_md_total")
    out["winland_mode"] = "classic"

    # --- 2) Macro‑normalized r35 (bimodal only, with guardrails)
    try:
        mode = str(res.get("thomeer_mode") or "").lower()
    except Exception:
        mode = ""

    if mode == "bimodal":
        # Guardrails (user-configurable)
        frac_min = float(params.get("winland_macro_frac_min", 0.15))
        frac_max = float(params.get("winland_macro_frac_max", 0.85))
        r35_macro_max = float(params.get("winland_r35_macro_max_um", 50.0))

        # Macro fraction from bimodal Thomeer
        f = res.get("thomeer_macro_frac")
        try:
            f = float(f)
        except Exception:
            f = float("nan")

        if not (np.isfinite(f) and 0.0 < f < 1.0):
            out["winland_warning"] = "Macro-normalized unavailable (missing macro_frac) → using classic r35."
        elif not (frac_min <= float(f) <= frac_max):
            out["winland_warning"] = (
                f"Macro-normalized unstable (macro_frac={float(f):.3g} outside [{frac_min:.2g}, {frac_max:.2g}]) "
                "→ using classic r35."
            )
        else:
            frac_target = 0.35 * float(f)
            r35_macro = compute_r_um_at_hgsat_fraction(df, params, frac_target)

            if r35_macro is None or not np.isfinite(float(r35_macro)) or float(r35_macro) <= 0:
                out["winland_warning"] = "Macro-normalized unavailable (insufficient saturation range) → using classic r35."
            else:
                # Keep candidate for diagnostics (not used unless it passes guardrails)
                out["r35_um_macro_candidate"] = float(r35_macro)
                out["r35_macro_target_hgsat_frac"] = float(frac_target)

                if float(r35_macro) > r35_macro_max:
                    out["winland_warning"] = (
                        f"Macro-normalized unstable (r35_macro={float(r35_macro):.3g} µm > {r35_macro_max:.0f} µm) "
                        "→ using classic r35."
                    )

                    # Optional candidate k (diagnostic only)
                    if phi_pct_f is not None:
                        k_macro_cand = _k_from_r35_phi(float(r35_macro), phi_pct_f)
                        if k_macro_cand is not None:
                            out["k_winland_md_macro_candidate"] = float(k_macro_cand)
                else:
                    # Accept macro-normalized
                    out["r35_um_macro"] = float(r35_macro)
                    out["winland_mode"] = "macro-normalized"
                    r35_used = float(r35_macro)

                    if phi_pct_f is not None:
                        k_macro = _k_from_r35_phi(float(r35_macro), phi_pct_f)
                        if k_macro is not None:
                            out["k_winland_md_macro"] = float(k_macro)
                            # Backward-compatible alias used by some multi-sample helpers
                            out["k_winland_macro_md"] = float(k_macro)
                            k_used = float(k_macro)

    # Final outputs (what the rest of the app uses)
    if r35_used is None or not np.isfinite(float(r35_used)) or float(r35_used) <= 0:
        out["winland_error"] = "Cannot compute r35 (insufficient saturation range)."
        return out

    out["r35_um"] = float(r35_used)
    if k_used is not None:
        try:
            if np.isfinite(float(k_used)) and float(k_used) > 0:
                out["k_winland_md"] = float(k_used)
        except Exception:
            pass

    # Also expose classic k under expected alias if needed
    if out.get("k_winland_md_total") is not None and out.get("k_winland_total_md") is None:
        out["k_winland_total_md"] = out.get("k_winland_md_total")

    out["winland_label"] = "Winland (Macro‑Normalized)" if out.get("winland_mode") == "macro-normalized" else "Winland (Classic)"
    return out
# ---------------------------
# Pore Network Modeling (PNM) — Fast statistical model
# ---------------------------

def compute_pnm_fast(
    df: pd.DataFrame,
    params: Dict[str, Any],
    meta: Dict[str, Any],
    res: Optional[Dict[str, Any]] = None,
) -> Dict[str, Any]:
    """Compute a lightweight statistical PNM permeability estimate.

    This is a fast, robust, explainable proxy driven by the throat-size distribution (PSD).
    It is NOT a full 3D network reconstruction. It estimates an effective permeability:

        k_pnm ≈ K0 * (1/tau_eff) * f(z) * Σ (w_i * r_i^2)

    where:
        - w_i are normalized PSD weights (dV/dlog(r) normalized to 1),
        - r_i are throat radii in microns,
        - z is coordination number,
        - tau_eff is effective tortuosity,
        - K0 is a scale factor (optionally calibrated using k_core if available).

    Output keys:
        - k_pnm_md
        - pnm_ci (k_pnm / k_core if core permeability exists)
        - pnm_z, pnm_tau_eff, pnm_constriction, pnm_k0
        - pnm_method = 'pnm_fast'
    """
    out: Dict[str, Any] = {}
    params = params or DEFAULT_PARAMS
    res = res or {}

    df = _ensure_schema(df)
    df = recompute_derived(df, params)

    # Need pore-throat radius for PSD
    if "r_um" not in df.columns or df["r_um"].isna().all():
        try:
            df = add_pore_throat_columns(df, params)
        except Exception:
            return out

    # Build PSD proxy: dV/dlog10(r) using incremental intrusion (mL/g)
    try:
        r = df["r_um"].to_numpy(dtype=float)
        inc = df["IncVol"].to_numpy(dtype=float)
        ok = np.isfinite(r) & np.isfinite(inc) & (r > 0) & (inc >= 0)
        r = r[ok]
        inc = inc[ok]
        if len(r) < 5:
            return out

        # Use log-spaced bins to build stable PSD
        log_r = np.log10(r)
        nbins = int(params.get("pnm_psd_bins", 40))
        edges = np.linspace(np.nanmin(log_r), np.nanmax(log_r), nbins + 1)
        bin_idx = np.digitize(log_r, edges) - 1
        bin_idx = np.clip(bin_idx, 0, nbins - 1)

        dv = np.zeros(nbins, dtype=float)
        r_mid = np.zeros(nbins, dtype=float)
        for i in range(nbins):
            mask = bin_idx == i
            if not np.any(mask):
                dv[i] = 0.0
                r_mid[i] = 10 ** ((edges[i] + edges[i + 1]) / 2)
            else:
                dv[i] = float(np.nansum(inc[mask]))
                r_mid[i] = 10 ** ((edges[i] + edges[i + 1]) / 2)

        # Normalize weights
        dv_sum = float(np.nansum(dv))
        if dv_sum <= 0:
            return out
        w = dv / dv_sum

        # Effective geometric moment (um^2)
        m2 = float(np.nansum(w * (r_mid ** 2)))
        if not np.isfinite(m2) or m2 <= 0:
            return out
    except Exception:
        return out

    # PNM parameters
    z = float(params.get("pnm_coordination_z", 4.0))
    z = max(1.5, min(12.0, z))
    tau_eff = params.get("pnm_tau_eff")
    if tau_eff is None:
        # Use measured tortuosity if available, else default
        tau_eff = meta.get("tortuosity") or meta.get("tortuosity_factor") or 2.5
    tau_eff = float(tau_eff) if tau_eff else 2.5
    tau_eff = max(1.0, min(50.0, tau_eff))

    constr = float(params.get("pnm_constriction", 0.5))
    constr = max(0.05, min(1.0, constr))

    # Connectivity factor f(z)
    fz = (z - 1.0) / z if z > 1 else 0.1
    fz = max(0.05, min(1.0, fz))

    # Scale factor K0 (mD / um^2). Optionally calibrate using k_core.
    k_core = params.get("k_override_md") or meta.get("permeability_md")
    if k_core is not None:
        try:
            k_core = float(k_core)
        except Exception:
            k_core = None

    k0 = params.get("pnm_k0_md_per_um2")
    if k0 is None:
        k0 = 50.0  # reasonable default scale
    k0 = float(k0)

    # Optional per-sample calibration
    if params.get("pnm_calibrate_to_core", True) and k_core and k_core > 0:
        # Keep fz, tau_eff fixed, solve for k0 to match k_core
        try:
            k0_cal = (k_core * tau_eff) / (fz * m2)
            if np.isfinite(k0_cal) and k0_cal > 0:
                # Avoid runaway calibration
                k0 = float(np.clip(k0_cal, 0.1, 1e6))
                out["pnm_k0_calibrated"] = True
        except Exception:
            pass

    # Compute k_pnm
    k_pnm = k0 * (1.0 / tau_eff) * fz * m2

    out.update(
        {
            "k_pnm_md": float(k_pnm),
            "pnm_method": "pnm_fast",
            "pnm_z": float(z),
            "pnm_tau_eff": float(tau_eff),
            "pnm_constriction": float(constr),
            "pnm_k0_md_per_um2": float(k0),
            "pnm_m2_um2": float(m2),
        }
    )
    if k_core and k_core > 0:
        out["pnm_ci"] = float(k_pnm / k_core)

    return out

def rock_type_from_r35(r35_um: Optional[float], params: Dict[str, Any]) -> str:
    if r35_um is None or not np.isfinite(r35_um):
        return "RT? (No r35)"
    bins = params.get("rt_bins_um", DEFAULT_PARAMS["rt_bins_um"])
    labels = params.get("rt_labels", DEFAULT_PARAMS["rt_labels"])
    try:
        bins_f = [float(x) for x in bins]
    except Exception:
        bins_f = DEFAULT_PARAMS["rt_bins_um"]
    # bins: [0.01,0.1,1,10] => 5 labels
    # <=0.01: labels[0], (0.01,0.1]: labels[1], ...
    x = float(r35_um)
    if x <= bins_f[0]:
        return labels[0]
    if x <= bins_f[1]:
        return labels[1]
    if x <= bins_f[2]:
        return labels[2]
    if x <= bins_f[3]:
        return labels[3]
    return labels[4]

# --- Thomeer fit
def thomeer_vb(pc_psia: np.ndarray, vb_inf: float, pd_psia: float, G: float) -> np.ndarray:
    """
    Thomeer (3-parameter) model (log10):
      (Vb(Pc) / Vb∞) = exp( -G / log10(Pc / Pd) )
    """
    pc_psia = np.asarray(pc_psia, dtype=float)
    pd_psia = max(float(pd_psia), 1e-12)
    x = np.log10(np.maximum(pc_psia, pd_psia * (1.0 + 1e-9)) / pd_psia)
    return vb_inf * np.exp(-G / (x + 1e-30))

def fit_thomeer(df: pd.DataFrame, params: Dict[str, Any], meta: Dict[str, Any]) -> Dict[str, Any]:
    out: Dict[str, Any] = {}

    d = df.copy()
    d = recompute_derived(d, params)

    pc = d["Pressure"].to_numpy(dtype=float)
    m = np.isfinite(pc) & (pc > 0)
    if m.sum() < 10:
        out["thomeer_error"] = "Not enough valid Pc points."
        return out

    d = d.loc[m].copy()
    pc = d["Pressure"].to_numpy(dtype=float)

    # Prefer bulk volume fraction if porosity exists, else use HgSat fraction as proxy
    phi_pct = params.get("phi_override_pct")
    if phi_pct is None:
        phi_pct = meta.get("porosity_pct")
    phi_frac = None
    try:
        phi_frac = float(phi_pct) / 100.0
        if not np.isfinite(phi_frac) or phi_frac <= 0:
            phi_frac = None
    except Exception:
        phi_frac = None

    if phi_frac is not None:
        vb = d["HgSat"].to_numpy(dtype=float) * phi_frac  # v/v bulk
        vb_label = "Vb (bulk fraction)"
        vb_inf_upper = max(phi_frac * 1.25, 0.05)
    else:
        vb = d["HgSat"].to_numpy(dtype=float)  # 0..1 proxy
        vb_label = "HgSat (proxy for Vb)"
        vb_inf_upper = 1.25

    m2 = np.isfinite(vb) & (vb >= 0) & np.isfinite(pc) & (pc > 0)
    pc = pc[m2]
    vb = vb[m2]
    if len(pc) < 10:
        out["thomeer_error"] = "Not enough valid points after cleaning."
        return out

    # Initial guesses
    vb_inf0 = float(np.nanmax(vb)) if np.isfinite(np.nanmax(vb)) else 1.0
    vb_inf0 = max(min(vb_inf0, vb_inf_upper), 1e-3)

    # Pd ~ first Pc where Vb reaches ~1% of vb_inf0
    try:
        idx_pd = int(np.argmax(vb >= 0.01 * vb_inf0))
        pd0 = float(pc[idx_pd])
    except Exception:
        pd0 = float(np.nanmin(pc))

    pd0 = max(pd0, 1e-6)
    G0 = 0.1

    if curve_fit is None:
        out["thomeer_error"] = "SciPy not available for fitting."
        return out

    try:
        popt, pcov = curve_fit(
            thomeer_vb,
            pc,
            vb,
            p0=[vb_inf0, pd0, G0],
            bounds=([0.0, 1e-6, 1e-4], [vb_inf_upper, np.nanmax(pc) * 10.0, 5.0]),
            maxfev=20000,
        )
        vb_inf, pd_fit, G_fit = [float(x) for x in popt]

        vb_pred = thomeer_vb(pc, vb_inf, pd_fit, G_fit)
        ss_res = float(np.nansum((vb - vb_pred) ** 2))
        ss_tot = float(np.nansum((vb - np.nanmean(vb)) ** 2))
        r2 = 1.0 - ss_res / ss_tot if ss_tot > 0 else np.nan

        out.update({
            "thomeer_vb_label": vb_label,
            "thomeer_vb_inf": vb_inf,
            "thomeer_pd_psia": pd_fit,
            "thomeer_G": G_fit,
            "thomeer_r2": r2,
        })
    except Exception as e:
        out["thomeer_error"] = f"Fit failed: {e}"

    return out




def _r2(y_true, y_pred) -> float:
    """Compute coefficient of determination (R²).

    Notes
    -----
    * Uses only finite (non-NaN/inf) pairs.
    * Returns NaN if there are fewer than 2 finite points or if the variance of
      *y_true* is zero.
    """
    yt = np.asarray(y_true, dtype=float)
    yp = np.asarray(y_pred, dtype=float)
    m = np.isfinite(yt) & np.isfinite(yp)
    if m.sum() < 2:
        return float('nan')
    yt = yt[m]
    yp = yp[m]
    ss_res = float(np.sum((yt - yp) ** 2))
    ss_tot = float(np.sum((yt - np.mean(yt)) ** 2))
    if ss_tot <= 0:
        return float('nan')
    return 1.0 - ss_res / ss_tot


def fit_thomeer_bimodal(pc, vb, vb_upper=None, initial=None):
    """Fit a *two pore-system* (macro + micro) Thomeer model.

    Key difference vs older implementations:
      - **Closure** is treated as a *prior conformance/closure correction* (volume correction).
      - **Pd1 and Pd2 are both fitted** (physical entry pressures), i.e. Pd1 is *not* forced to equal closure.

    Parameters
    ----------
    pc : array-like
        Capillary pressure (psia), must be > 0.
    vb : array-like
        Cumulative mercury intrusion expressed as a fraction of bulk volume (0-1).
    vb_upper : float, optional
        Upper bound for vb_total. If you pass porosity fraction (phi/100),
        it will constrain vb_total to a physically meaningful range.
    initial : dict, optional
        Optional initial guesses: vb_total, frac, pd1, G1, pd2, G2.

    Returns
    -------
    dict
        Keys compatible with the rest of the app (thomeer_*).
    """
    pc = np.asarray(pc, dtype=float)
    vb = np.asarray(vb, dtype=float)

    mask = np.isfinite(pc) & np.isfinite(vb) & (pc > 0) & (vb >= 0)
    pc = pc[mask]
    vb = vb[mask]

    if pc.size < 8:
        return {"thomeer_error": "Not enough valid points for bimodal fit", "thomeer_mode": "bimodal"}

    # Sort by Pc (intrusion curves should be monotonic in Pc).
    order = np.argsort(pc)
    pc = pc[order]
    vb = vb[order]

    # Enforce non-decreasing vb to reduce instabilities from noisy lab data.
    vb = np.maximum.accumulate(vb)

    vb_max = float(np.nanmax(vb)) if vb.size else 0.0
    if vb_upper is None or (not np.isfinite(vb_upper)) or vb_upper <= 0:
        vb_upper = vb_max if vb_max > 0 else 1.0
    vb_upper = float(vb_upper)

    vb_total_init = vb_max if vb_max > 0 else vb_upper
    vb_total_init = float(min(vb_upper, vb_total_init))

    if initial is None:
        initial = {}

    # Helper: pick Pc at a given cumulative fill fraction of vb_total_init.
    def _pc_at_vb_frac(frac):
        if vb_total_init <= 0:
            return np.nan
        target = float(frac) * vb_total_init
        idx = int(np.searchsorted(vb, target, side="left"))
        if idx < 0 or idx >= pc.size:
            return np.nan
        return float(pc[idx])

    frac_init = float(initial.get("frac", 0.50))
    frac_init = float(np.clip(frac_init, 0.05, 0.95))

    pd1_init = float(initial.get("pd1", np.nan))
    pd2_init = float(initial.get("pd2", np.nan))

    # Data-driven guesses (macro enters early, micro later).
    if not np.isfinite(pd1_init):
        pd1_init = _pc_at_vb_frac(0.05)
    if not np.isfinite(pd2_init):
        pd2_init = _pc_at_vb_frac(0.60)

    if not np.isfinite(pd1_init):
        pd1_init = float(np.nanmin(pc))
    if not np.isfinite(pd2_init):
        pd2_init = float(np.nanmax(pc))

    pd1_init = max(1e-6, float(pd1_init))
    pd2_init = max(pd1_init * 1.10, float(pd2_init))
    if pd2_init <= pd1_init:
        pd2_init = pd1_init * 10.0

    G1_init = float(initial.get("G1", 0.50))
    G2_init = float(initial.get("G2", 1.00))

    vb_total_guess = float(initial.get("vb_total", vb_total_init))
    vb_total_guess = max(0.0, vb_total_guess)

    p0 = [vb_total_guess, frac_init, pd1_init, G1_init, pd2_init, G2_init]

    pc_max = float(np.nanmax(pc))
    lower = [0.0, 0.0, 1e-6, 1e-4, 1e-6, 1e-4]
    upper = [vb_upper * 1.2, 1.0, pc_max * 10.0, 5.0, pc_max * 10.0, 5.0]

    def model(p, vb_total, frac, pd1, G1, pd2, G2):
        vb1 = vb_total * frac
        vb2 = vb_total * (1.0 - frac)
        return thomeer_vb(p, vb1, pd1, G1) + thomeer_vb(p, vb2, pd2, G2)

    try:
        popt, _ = curve_fit(
            model,
            pc,
            vb,
            p0=p0,
            bounds=(lower, upper),
            maxfev=60000,
        )
    except Exception as e:
        return {"thomeer_error": f"Bimodal fit failed: {e}", "thomeer_mode": "bimodal"}

    vb_total_fit, frac_fit, pd1_fit, G1_fit, pd2_fit, G2_fit = map(float, popt)

    # Enforce ordering: Pd1 <= Pd2 (macro=lower entry pressure).
    if pd1_fit > pd2_fit:
        pd1_fit, pd2_fit = pd2_fit, pd1_fit
        G1_fit, G2_fit = G2_fit, G1_fit
        frac_fit = 1.0 - frac_fit

    vb_macro = vb_total_fit * frac_fit
    vb_micro = vb_total_fit * (1.0 - frac_fit)

    vb_fit = model(pc, vb_total_fit, frac_fit, pd1_fit, G1_fit, pd2_fit, G2_fit)
    r2 = _r2(vb, vb_fit)

    out = {
        "thomeer_mode": "bimodal",
        "thomeer_bimodal": True,
        "thomeer_r2": r2,
        "thomeer_vb_total": vb_total_fit,
        "thomeer_macro_frac": frac_fit,

        "thomeer_vb_inf1": vb_macro,
        "thomeer_vb_inf2": vb_micro,
        "thomeer_pd1_psia": pd1_fit,
        "thomeer_pd2_psia": pd2_fit,
        "thomeer_G1": G1_fit,
        "thomeer_G2": G2_fit,

        # Convenience / backward compatible fields (use macro as the 'primary' system).
        "thomeer_vb_inf": vb_total_fit,
        "thomeer_pd_psia": pd1_fit,
        "thomeer_G": G1_fit,
        "thomeer_pd_locked": False,

        # Bulk-volume percent outputs (useful for UI + export).
        "thomeer_bvinf_pct": vb_total_fit * 100.0,
        "thomeer_bvinf1_pct": vb_macro * 100.0,
        "thomeer_bvinf2_pct": vb_micro * 100.0,
    }
    return out

def fit_thomeer_fixed_pd(df: pd.DataFrame, pd_psia: float, params: Dict[str, Any], meta: Dict[str, Any]) -> Dict[str, Any]:
    """Fit Thomeer with *fixed* Pd (linked to closure correction).
    We fit only (Vb∞, G) while keeping Pd constant.
    """
    out: Dict[str, Any] = {}

    d = df.copy()
    d = recompute_derived(d, params)

    pc = d["Pressure"].to_numpy(dtype=float)
    m = np.isfinite(pc) & (pc > 0)
    if m.sum() < 10:
        out["thomeer_error"] = "Not enough valid Pc points."
        return out
    d = d.loc[m].copy()
    pc = d["Pressure"].to_numpy(dtype=float)

    # vb series as in fit_thomeer
    phi_pct = params.get("phi_override_pct")
    if phi_pct is None:
        phi_pct = meta.get("porosity_pct")

    phi_frac = None
    try:
        phi_frac = float(phi_pct) / 100.0
        if not np.isfinite(phi_frac) or phi_frac <= 0:
            phi_frac = None
    except Exception:
        phi_frac = None

    if phi_frac is not None:
        vb = d["HgSat"].to_numpy(dtype=float) * phi_frac
        vb_label = "Bv (bulk fraction)"
        vb_inf_upper = max(phi_frac * 1.25, 0.05)
    else:
        vb = d["HgSat"].to_numpy(dtype=float)
        vb_label = "HgSat (proxy for Bv)"
        vb_inf_upper = 1.25

    m2 = np.isfinite(vb) & (vb >= 0) & np.isfinite(pc) & (pc > 0)
    pc = pc[m2]
    vb = vb[m2]
    if len(pc) < 10:
        out["thomeer_error"] = "Not enough valid points after cleaning."
        return out

    # Fixed Pd sanity
    try:
        pd_fixed = float(pd_psia)
    except Exception:
        pd_fixed = float(np.nanmin(pc))
    pd_fixed = max(pd_fixed, 1e-6)
    # Make sure Pd is within (min..max)
    pd_fixed = min(max(pd_fixed, float(np.nanmin(pc))), float(np.nanmax(pc)))

    if curve_fit is None:
        out["thomeer_error"] = "SciPy not available for fitting."
        return out

    # Initial guesses
    vb_inf0 = float(np.nanmax(vb)) if np.isfinite(np.nanmax(vb)) else 1.0
    vb_inf0 = max(min(vb_inf0, vb_inf_upper), 1e-3)
    G0 = 0.1

    try:
        popt, _ = curve_fit(
            lambda pc_in, vb_inf_in, G_in: thomeer_vb(pc_in, vb_inf_in, pd_fixed, G_in),
            pc,
            vb,
            p0=[vb_inf0, G0],
            bounds=([0.0, 1e-4], [vb_inf_upper, 5.0]),
            maxfev=20000,
        )
        vb_inf_fit, G_fit = [float(x) for x in popt]

        vb_pred = thomeer_vb(pc, vb_inf_fit, pd_fixed, G_fit)
        ss_res = float(np.nansum((vb - vb_pred) ** 2))
        ss_tot = float(np.nansum((vb - np.nanmean(vb)) ** 2))
        r2 = 1.0 - ss_res / ss_tot if ss_tot > 0 else np.nan

        out.update({
            "thomeer_vb_label": vb_label,
            "thomeer_vb_inf": vb_inf_fit,
            "thomeer_bvinf_pct": vb_inf_fit * 100.0,
            "thomeer_pd_psia": pd_fixed,
            "thomeer_G": G_fit,
            "thomeer_r2": r2,
            "thomeer_mode": "auto_fixed_pd",
            "thomeer_pd_locked": True,
        })
    except Exception as e:
        out["thomeer_error"] = f"Fit failed: {e}"

    return out



# --- Thomeer (bi-modal) fit
def fit_thomeer_bimodal_fixed_pd1(df: pd.DataFrame, pd1_psia: float, params: Dict[str, Any], meta: Dict[str, Any]) -> Dict[str, Any]:
    """Fit a *bi-modal* Thomeer model using a fixed Pd for the macro system (Pd1).

    This is useful when the intrusion curve suggests two largely independent pore systems:
      - Macro system (reservoir): lower Pd, typically controls permeability.
      - Micro system (matrix/clay): higher Pd, adds volume but contributes little to flow.

    Model (sum of two Thomeer components):
        Vb(Pc) = Vb1_inf * exp(-G1 / log10(Pc/Pd1)) + Vb2_inf * exp(-G2 / log10(Pc/Pd2))

    Implementation details:
    - Pd1 is fixed (typically tied to the Closure/Conformance knee).
    - We fit: (Vb_total, frac_macro, G1, Pd2, G2) where:
        Vb1_inf = Vb_total * frac_macro
        Vb2_inf = Vb_total * (1 - frac_macro)
      This guarantees Vb1_inf + Vb2_inf = Vb_total (bounded by porosity-derived upper bound).
    """
    out: Dict[str, Any] = {}

    d = df.copy()
    d = recompute_derived(d, params)

    pc = d["Pressure"].to_numpy(dtype=float)
    m = np.isfinite(pc) & (pc > 0)
    if m.sum() < 12:
        out["thomeer_error"] = "Not enough valid Pc points for bi-modal fit."
        return out

    d = d.loc[m].copy()
    pc = d["Pressure"].to_numpy(dtype=float)

    # vb series as in unimodal fits
    phi_pct = params.get("phi_override_pct")
    if phi_pct is None:
        phi_pct = meta.get("porosity_pct")

    phi_frac = None
    try:
        phi_frac = float(phi_pct) / 100.0
        if not np.isfinite(phi_frac) or phi_frac <= 0:
            phi_frac = None
    except Exception:
        phi_frac = None

    if phi_frac is not None:
        vb = d["HgSat"].to_numpy(dtype=float) * phi_frac
        vb_label = "Bv (bulk fraction)"
        vb_inf_upper = max(phi_frac * 1.25, 0.05)
    else:
        vb = d["HgSat"].to_numpy(dtype=float)
        vb_label = "HgSat (proxy for Bv)"
        vb_inf_upper = 1.25

    m2 = np.isfinite(vb) & (vb >= 0) & np.isfinite(pc) & (pc > 0)
    pc = pc[m2]
    vb = vb[m2]
    if len(pc) < 12:
        out["thomeer_error"] = "Not enough valid points after cleaning for bi-modal fit."
        return out

    # Fixed Pd1 sanity
    try:
        pd1_fixed = float(pd1_psia)
    except Exception:
        pd1_fixed = float(np.nanmin(pc))
    pd1_fixed = max(pd1_fixed, 1e-6)
    pd1_fixed = min(max(pd1_fixed, float(np.nanmin(pc))), float(np.nanmax(pc)))

    if curve_fit is None:
        out["thomeer_error"] = "SciPy not available for bi-modal fitting."
        return out

    # Helper: pick Pc at a given cumulative fraction (for initial guesses)
    def _pc_at_vb_frac(frac: float) -> float:
        frac = float(frac)
        frac = max(0.0, min(1.0, frac))
        # sort by Pc ascending
        order = np.argsort(pc)
        pc_s = pc[order]
        vb_s = vb[order]
        v_target = frac * float(np.nanmax(vb_s))
        idx = int(np.argmax(vb_s >= v_target))
        try:
            return float(pc_s[idx])
        except Exception:
            return float(np.nanmax(pc_s))

    vb_total0 = float(np.nanmax(vb)) if np.isfinite(np.nanmax(vb)) else 1.0
    vb_total0 = max(min(vb_total0, vb_inf_upper), 1e-3)

    # Candidate starting points (robust multi-start)
    frac_guesses = [0.25, 0.40, 0.55, 0.70]
    G1_guesses = [0.05, 0.10, 0.20, 0.35]
    G2_guesses = [0.35, 0.60, 1.00, 1.60]

    # Pd2 should be > Pd1; use late-stage points as seeds
    pd2_candidates = sorted({
        max(pd1_fixed * 1.5, _pc_at_vb_frac(0.70)),
        max(pd1_fixed * 2.0, _pc_at_vb_frac(0.85)),
        max(pd1_fixed * 3.0, _pc_at_vb_frac(0.93)),
        max(pd1_fixed * 5.0, _pc_at_vb_frac(0.97)),
        max(pd1_fixed * 10.0, float(np.nanmax(pc)) * 0.2),
    })
    pd2_candidates = [p for p in pd2_candidates if np.isfinite(p) and p > pd1_fixed]
    if not pd2_candidates:
        pd2_candidates = [max(pd1_fixed * 3.0, float(np.nanmax(pc)) * 0.2)]

    pd2_lower = max(pd1_fixed * 1.0001, float(np.nanmin(pc)))
    pd2_upper = float(np.nanmax(pc)) * 10.0
    if pd2_upper <= pd2_lower:
        pd2_upper = pd2_lower * 10.0

    bounds_lo = [0.0, 0.0, 1e-4, pd2_lower, 1e-4]  # (vb_total, frac_macro, G1, Pd2, G2)
    bounds_hi = [vb_inf_upper, 1.0, 5.0, pd2_upper, 5.0]

    best = None  # (r2, popt)

    def _model(pc_in: np.ndarray, vb_total: float, frac_macro: float, G1: float, pd2: float, G2: float) -> np.ndarray:
        return (
            thomeer_vb(pc_in, vb_total * frac_macro, pd1_fixed, G1)
            + thomeer_vb(pc_in, vb_total * (1.0 - frac_macro), pd2, G2)
        )

    for frac0 in frac_guesses:
        for G10 in G1_guesses:
            for pd20 in pd2_candidates:
                for G20 in G2_guesses:
                    p0 = [vb_total0, frac0, G10, pd20, G20]
                    try:
                        popt, _ = curve_fit(
                            _model,
                            pc,
                            vb,
                            p0=p0,
                            bounds=(bounds_lo, bounds_hi),
                            maxfev=50000,
                        )
                        vb_total_fit, frac_fit, G1_fit, pd2_fit, G2_fit = [float(x) for x in popt]

                        vb_pred = _model(pc, vb_total_fit, frac_fit, G1_fit, pd2_fit, G2_fit)
                        ss_res = float(np.nansum((vb - vb_pred) ** 2))
                        ss_tot = float(np.nansum((vb - np.nanmean(vb)) ** 2))
                        r2 = 1.0 - ss_res / ss_tot if ss_tot > 0 else np.nan

                        if not np.isfinite(r2):
                            continue

                        if (best is None) or (r2 > best[0]):
                            best = (r2, (vb_total_fit, frac_fit, G1_fit, pd2_fit, G2_fit))
                    except Exception:
                        continue

    if best is None:
        out["thomeer_error"] = "Bi-modal Thomeer fit failed (no converged solutions)."
        return out

    r2_best, (vb_total_fit, frac_fit, G1_fit, pd2_fit, G2_fit) = best
    frac_fit = float(np.clip(frac_fit, 0.0, 1.0))
    vb1 = vb_total_fit * frac_fit
    vb2 = vb_total_fit * (1.0 - frac_fit)

    out.update({
        "thomeer_vb_label": vb_label,

        # Backward-compatible keys (represent the macro system + total Bv∞)
        "thomeer_mode": "bimodal",
        "thomeer_pd_locked": True,
        "thomeer_pd_psia": float(pd1_fixed),
        "thomeer_G": float(G1_fit),
        "thomeer_vb_inf": float(vb_total_fit),
        "thomeer_bvinf_pct": float(vb_total_fit) * 100.0,
        "thomeer_r2": float(r2_best),

        # Explicit bi-modal parameters
        "thomeer_pd1_psia": float(pd1_fixed),
        "thomeer_G1": float(G1_fit),
        "thomeer_vb_inf1": float(vb1),
        "thomeer_bvinf1_pct": float(vb1) * 100.0,

        "thomeer_pd2_psia": float(pd2_fit),
        "thomeer_G2": float(G2_fit),
        "thomeer_vb_inf2": float(vb2),
        "thomeer_bvinf2_pct": float(vb2) * 100.0,

        "thomeer_vb_inf_total": float(vb_total_fit),
        "thomeer_macro_frac": float(frac_fit),
    })

    return out
# -------------------------------------------------
# Bi-modal pore-system hinting (auto-detection)
# -------------------------------------------------
def _detect_psd_peaks(df: pd.DataFrame, params: Dict[str, Any]) -> Dict[str, Any]:
    """
    Detect multiple pore systems using the PSD proxy (normalized dV/dlog(r)).

    Returns a dict with:
      - peaks: list of peaks (sorted by log10(r) descending)
      - n_peaks: number of selected, well-separated peaks
      - max_y: max of smoothed PSD
      - min_height: threshold used
    """
    try:
        d = recompute_derived(df.copy(), params or DEFAULT_PARAMS)
    except Exception:
        d = df.copy()

    if "dVdlogr_norm" not in d.columns or "r_um" not in d.columns:
        return {"peaks": [], "n_peaks": 0, "max_y": 0.0, "min_height": 0.0}

    y = pd.to_numeric(d["dVdlogr_norm"], errors="coerce").to_numpy(dtype=float)
    x = pd.to_numeric(d["r_um"], errors="coerce").to_numpy(dtype=float)

    m = np.isfinite(y) & np.isfinite(x) & (x > 0) & (y >= 0)
    if int(m.sum()) < 7:
        return {"peaks": [], "n_peaks": 0, "max_y": float(np.nanmax(y)) if len(y) else 0.0, "min_height": 0.0}

    x = x[m]
    y = y[m]

    # Sort by pore-throat radius so local maxima are meaningful.
    order = np.argsort(x)
    x = x[order]
    y = y[order]

    # Smooth (rolling median + rolling mean) to reduce noisy spikes.
    win = int((params or {}).get("bimodal_peak_smooth_window", DEFAULT_PARAMS.get("bimodal_peak_smooth_window", 7)))
    win = max(3, min(win, 31))  # keep it reasonable
    ys = pd.Series(y).rolling(window=win, center=True, min_periods=1).median()
    ys = ys.rolling(window=max(3, win // 2), center=True, min_periods=1).mean().to_numpy(dtype=float)

    max_y = float(np.nanmax(ys)) if len(ys) else 0.0
    if not np.isfinite(max_y) or max_y <= 0:
        return {"peaks": [], "n_peaks": 0, "max_y": max_y, "min_height": 0.0}

    min_frac = float((params or {}).get("bimodal_peak_min_frac", DEFAULT_PARAMS.get("bimodal_peak_min_frac", 0.35)))
    min_frac = float(np.clip(min_frac, 0.05, 0.95))
    min_height = max_y * min_frac

    # Candidate local maxima (simple, robust)
    candidates: List[int] = []
    for i in range(1, len(ys) - 1):
        if ys[i] > ys[i - 1] and ys[i] >= ys[i + 1] and ys[i] >= min_height:
            candidates.append(i)

    if not candidates:
        return {"peaks": [], "n_peaks": 0, "max_y": max_y, "min_height": min_height}

    # Enforce separation in log10(r) so we don't count shoulder/roughness as a second system
    sep = float((params or {}).get("bimodal_peak_min_sep_log10r", DEFAULT_PARAMS.get("bimodal_peak_min_sep_log10r", 0.40)))
    sep = float(np.clip(sep, 0.05, 2.0))
    lx = np.log10(np.maximum(x, 1e-30))

    # Select strongest peaks first, keep only well-separated ones
    candidates = sorted(candidates, key=lambda i: ys[i], reverse=True)
    selected: List[int] = []
    for i in candidates:
        if all(abs(lx[i] - lx[j]) >= sep for j in selected):
            selected.append(i)
        if len(selected) >= 3:
            break

    selected = sorted(selected, key=lambda i: lx[i], reverse=True)

    peaks = []
    for i in selected:
        peaks.append({
            "r_um": float(x[i]),
            "log10r": float(lx[i]),
            "psd_norm": float(ys[i]),
        })

    return {"peaks": peaks, "n_peaks": len(peaks), "max_y": max_y, "min_height": min_height}


def compute_bimodal_flags(df: pd.DataFrame, params: Dict[str, Any], res: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    """
    Compute an automatic hint to try the bi-modal Thomeer fit.

    Trigger logic (suggestion):
      - Unimodal Thomeer R² is below threshold, OR
      - PSD proxy shows >= 2 well-separated peaks.

    Returned dict is JSON-serializable (safe for dcc.Store).
    """
    params = params or DEFAULT_PARAMS
    res = res or {}

    mode = (res.get("thomeer_mode") or "unimodal").lower().strip()
    r2_thr = float(params.get("bimodal_hint_r2_threshold", DEFAULT_PARAMS.get("bimodal_hint_r2_threshold", 0.93)))

    reasons: List[str] = []
    peaks_payload: List[Dict[str, Any]] = []

    # 1) R² criterion (only meaningful after a unimodal fit)
    r2_low = False
    r2_val = None
    try:
        r2_val = float(res.get("thomeer_r2"))
        if np.isfinite(r2_val) and mode != "bimodal" and r2_val < r2_thr:
            r2_low = True
            reasons.append(f"Low unimodal Thomeer R² ({r2_val:.3f} < {r2_thr:.3f})")
    except Exception:
        r2_val = None

    # 2) Peak criterion (PSD proxy)
    pk = _detect_psd_peaks(df, params)
    two_peaks = bool(pk.get("n_peaks", 0) >= 2)
    if two_peaks:
        peaks_payload = pk.get("peaks", []) or []
        # Build a compact description of the main two peaks
        try:
            ptxt = ", ".join([f"r≈{p['r_um']:.3g}µm" for p in peaks_payload[:2]])
            reasons.append(f"Two PSD peaks detected ({ptxt})")
        except Exception:
            reasons.append("Two PSD peaks detected")

    detected = bool(r2_low or two_peaks)

    # Confidence (simple scoring): 1 criterion -> MEDIUM, 2 criteria -> HIGH
    bimodal_score = int(bool(r2_low)) + int(bool(two_peaks))
    bimodal_confidence = "HIGH" if bimodal_score >= 2 else ("MEDIUM" if bimodal_score == 1 else "LOW")

    # Peak separation diagnostics (if available)
    bimodal_peak_sep_log10r = None
    bimodal_peak_ratio = None
    if len(peaks_payload) >= 2:
        try:
            bimodal_peak_sep_log10r = abs(float(peaks_payload[0].get("log10r")) - float(peaks_payload[1].get("log10r")))
        except Exception:
            bimodal_peak_sep_log10r = None
        try:
            p1 = float(peaks_payload[0].get("psd_norm"))
            p2 = float(peaks_payload[1].get("psd_norm"))
            if np.isfinite(p1) and np.isfinite(p2) and p1 > 0:
                bimodal_peak_ratio = float(p2 / p1)
        except Exception:
            bimodal_peak_ratio = None

    # Suggestion only if user has NOT already run the bi-modal fit
    hint = bool(detected and mode != "bimodal")

    hint_text = ""
    if hint:
        hint_text = f"Possible bi-modal pore system ({bimodal_confidence}) → try Fit Thomeer (Bimodal)."
        if reasons:
            hint_text += " Reasons: " + "; ".join(reasons) + "."

    return {
        "bimodal_detected": bool(detected),
        "bimodal_hint": bool(hint),
        "bimodal_reasons": reasons,
        "bimodal_peaks": peaks_payload,
        "bimodal_r2_threshold": r2_thr,
        "bimodal_r2_value": float(r2_val) if (r2_val is not None and np.isfinite(r2_val)) else None,
        "bimodal_score": int(bimodal_score),
        "bimodal_confidence": str(bimodal_confidence),
        "bimodal_peak_sep_log10r": float(bimodal_peak_sep_log10r) if (bimodal_peak_sep_log10r is not None and np.isfinite(bimodal_peak_sep_log10r)) else None,
        "bimodal_peak_ratio": float(bimodal_peak_ratio) if (bimodal_peak_ratio is not None and np.isfinite(bimodal_peak_ratio)) else None,
        "bimodal_hint_text": hint_text,
    }




# --- SHF & J-function
def _bimodal_hints(df: pd.DataFrame, params: Dict[str, Any], res_store: Dict[str, Any]) -> Dict[str, Any]:
    """Return auxiliary, JSON‑serializable hints for the Thomeer bimodal workflow.

    This helper is intentionally defensive: it **must not** raise, because it is called from a Dash
    callback (Fit Thomeer (Bimodal)). Earlier builds referenced this function but didn't define it,
    which triggered a 500 (Internal Server Error).

    The returned dict can be merged into the sample's ``results`` store.
    """
    out: Dict[str, Any] = {}
    try:
        flags = compute_bimodal_flags(df, params)

        # Copy a stable subset (keep JSON-serializable types).
        for k in (
            "bimodal_candidate",
            "bimodal_reason",
            "pd_sep_log10",
            "min_sep_log10",
            "bimodal_peak_ratio",
            "bimodal_peaks",
            "bimodal_hint_title",
            "bimodal_hint_text",
            "bimodal_hint_level",
        ):
            if k in flags:
                v = flags.get(k)
                # Convert numpy scalars to plain Python numbers for JSON serialization.
                if isinstance(v, (np.floating, np.integer)):
                    v = float(v)
                out[k] = v

        # Optional: add a compact fit summary if we already have fit metrics.
        if isinstance(res_store, dict):
            parts: List[str] = []
            r2 = res_store.get("thomeer_bimodal_r2")
            if _is_finite(r2):
                parts.append(f"R²={float(r2):.3f}")
            macro = res_store.get("thomeer_macro_frac")
            if _is_finite(macro):
                parts.append(f"macro={float(macro):.3f}")
            if parts:
                out["bimodal_fit_summary"] = ", ".join(parts)

    except Exception as e:
        # Never break callbacks because of hint generation.
        out["bimodal_hint_error"] = f"{type(e).__name__}: {e}"

    return out

def compute_shf(df: pd.DataFrame, params: Dict[str, Any]) -> pd.DataFrame:
    """
    Saturation height function (very standard scaling):
      Pc_res = Pc_hg_air * (σ_res cosθ_res) / (σ_hg_air cosθ_hg_air)
      Height = Pc_res / (Δρ g)
    Pc in Pa, Height in m.
    """
    d = df.copy()
    pc_psia = d["Pressure"].to_numpy(dtype=float)
    pc_pa = pc_psia * 6894.757293168

    sigma_hg = float(params.get("sigma_hg_air_npm", DEFAULT_PARAMS["sigma_hg_air_npm"]))
    theta_hg = math.radians(float(params.get("theta_hg_air_deg", DEFAULT_PARAMS["theta_hg_air_deg"])))
    sigma_res = float(params.get("sigma_res_npm", DEFAULT_PARAMS["sigma_res_npm"]))
    theta_res = math.radians(float(params.get("theta_res_deg", DEFAULT_PARAMS["theta_res_deg"])))

    # Use absolute value of cos to avoid sign flips in scaling
    num = sigma_res * abs(math.cos(theta_res))
    den = sigma_hg * abs(math.cos(theta_hg)) + 1e-30
    scale = num / den

    pc_res_pa = pc_pa * scale
    d["Pc_res_pa"] = pc_res_pa

    rho_w = float(params.get("rho_w_kgm3", DEFAULT_PARAMS["rho_w_kgm3"]))
    rho_hc = float(params.get("rho_hc_kgm3", DEFAULT_PARAMS["rho_hc_kgm3"]))
    drho = abs(rho_w - rho_hc)
    g = 9.80665
    if drho <= 0:
        d["Height_m"] = np.nan
    else:
        d["Height_m"] = pc_res_pa / (drho * g)

    return d

def compute_j_function(df: pd.DataFrame, params: Dict[str, Any], meta: Dict[str, Any], results: Dict[str, Any]) -> Dict[str, Any]:
    out: Dict[str, Any] = {}

    # Need phi and k
    phi_pct = params.get("phi_override_pct")
    if phi_pct is None:
        phi_pct = meta.get("porosity_pct")
    k_md = params.get("k_override_md")
    if k_md is None:
        k_md = meta.get("permeability_md")
    if k_md is None:
        k_md = results.get("k_swanson_md")
    if k_md is None:
        k_md = results.get("k_winland_md")

    try:
        phi_frac = float(phi_pct) / 100.0
        k_md_f = float(k_md)
    except Exception:
        out["j_error"] = "Need porosity and permeability (set overrides or ensure meta is parsed)."
        return out

    if not np.isfinite(phi_frac) or phi_frac <= 0 or not np.isfinite(k_md_f) or k_md_f <= 0:
        out["j_error"] = "Invalid porosity/permeability values."
        return out

    d = recompute_derived(df.copy(), params)
    pc_res = d["Pc_res_pa"].to_numpy(dtype=float)

    # k: mD -> m^2
    k_m2 = k_md_f * 9.869233e-16

    sigma_res = float(params.get("sigma_res_npm", DEFAULT_PARAMS["sigma_res_npm"]))
    theta_res = math.radians(float(params.get("theta_res_deg", DEFAULT_PARAMS["theta_res_deg"])))
    denom = sigma_res * abs(math.cos(theta_res)) + 1e-30

    J = pc_res * math.sqrt(k_m2 / phi_frac) / denom
    out["J"] = J.tolist()
    return out

# --- Clustering (no sklearn dependency)
def _kmeans(X: np.ndarray, k: int, n_init: int = 8, max_iter: int = 100, seed: int = 7) -> Tuple[np.ndarray, np.ndarray]:
    rng = np.random.default_rng(seed)
    best_inertia = np.inf
    best_labels = np.zeros(X.shape[0], dtype=int)
    best_centroids = np.zeros((k, X.shape[1]), dtype=float)

    for _ in range(n_init):
        # init centroids from random samples
        idx = rng.choice(X.shape[0], size=k, replace=False)
        centroids = X[idx].copy()

        for _it in range(max_iter):
            # assign
            dist2 = ((X[:, None, :] - centroids[None, :, :]) ** 2).sum(axis=2)
            labels = dist2.argmin(axis=1)
            # update
            new_centroids = np.vstack([X[labels == j].mean(axis=0) if np.any(labels == j) else centroids[j] for j in range(k)])
            if np.allclose(new_centroids, centroids, atol=1e-6, rtol=1e-6):
                centroids = new_centroids
                break
            centroids = new_centroids

        inertia = float(((X - centroids[labels]) ** 2).sum())
        if inertia < best_inertia:
            best_inertia = inertia
            best_labels = labels.copy()
            best_centroids = centroids.copy()

    return best_labels, best_centroids

def cluster_library(library: List[Dict[str, Any]], params: Dict[str, Any]) -> List[Dict[str, Any]]:
    # Build feature matrix: [phi, log10(r35), log10(k)]
    feats = []
    idx_map = []

    for i, s in enumerate(library):
        meta = s.get("meta", {}) or {}
        res = s.get("results", {}) or {}

        phi = params.get("phi_override_pct") or meta.get("porosity_pct")
        r35 = res.get("r35_um")
        k = params.get("k_override_md") or meta.get("permeability_md") or res.get("k_swanson_md") or res.get("k_winland_md")

        try:
            phi_f = float(phi)
            r35_f = float(r35)
            k_f = float(k)
        except Exception:
            continue

        if not (np.isfinite(phi_f) and phi_f > 0 and np.isfinite(r35_f) and r35_f > 0 and np.isfinite(k_f) and k_f > 0):
            continue

        feats.append([phi_f, math.log10(r35_f), math.log10(k_f)])
        idx_map.append(i)

    if len(feats) < 2:
        return library

    X = np.asarray(feats, dtype=float)
    k = int(params.get("cluster_k", DEFAULT_PARAMS["cluster_k"]))
    k = max(2, min(k, len(feats)))

    labels, centroids = _kmeans(X, k=k)
    for lbl, idx in zip(labels, idx_map):
        library[idx].setdefault("results", {})
        library[idx]["results"]["cluster"] = int(lbl) + 1  # 1-based
    return library

# ---------------------------
# Plotly figures (interactive)
# ---------------------------
import plotly.graph_objects as go

from plotly.subplots import make_subplots
def fig_intrusion(df: pd.DataFrame, ui: Dict[str, Any]) -> go.Figure:
    d = df.copy()
    fig = go.Figure()

    x = d["Pressure"]
    y1 = d["CumVol"]
    fig.add_trace(go.Scatter(x=x, y=y1, mode="lines+markers", name="Cumulative", line=dict(color=COLORS["curve2"])))

    if ui.get("overlay_inc", False):
        y2 = d["IncVol"]
        fig.add_trace(go.Scatter(x=x, y=y2, mode="lines+markers", name="Incremental", yaxis="y2", line=dict(color=COLORS["curve1"])))

        fig.update_layout(
            yaxis2=dict(title="IncVol (mL/g)", overlaying="y", side="right", showgrid=False),
        )

    fig.update_layout(
        template="plotly_dark",
        margin=dict(l=45, r=45, t=35, b=45),
        paper_bgcolor=COLORS["panel_bg"],
        plot_bgcolor=COLORS["app_bg"],
        font=dict(color=COLORS["text"]),
        xaxis_title="Pressure (psia)",
        yaxis_title="CumVol (mL/g)",
        legend=dict(orientation="h"),
    )

    if ui.get("xlog", True):
        fig.update_xaxes(type="log")
    else:
        fig.update_xaxes(type="linear")

    return fig

def fig_pc_sw(df: pd.DataFrame, ui: Dict[str, Any]) -> go.Figure:
    d = df.copy()
    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=d["Pressure"], y=d["Sw"],
        mode="lines+markers",
        name="Pc vs Sw",
        line=dict(color=COLORS["curve1"]),
    ))
    fig.update_layout(
        template="plotly_dark",
        margin=dict(l=45, r=45, t=35, b=45),
        paper_bgcolor=COLORS["panel_bg"],
        plot_bgcolor=COLORS["app_bg"],
        font=dict(color=COLORS["text"]),
        xaxis_title="Pressure (psia)",
        yaxis_title="Sw (fraction)",
        legend=dict(orientation="h"),
    )
    fig.update_xaxes(type="log" if ui.get("xlog", True) else "linear")
    fig.update_yaxes(range=[0, 1])
    return fig

def fig_psd(df: pd.DataFrame, ui: Dict[str, Any]) -> go.Figure:
    d = df.copy()
    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=d["r_um"], y=d["dVdlogr_norm"],
        mode="lines",
        name="dV/dlog(r) (norm)",
        line=dict(color=COLORS["curve2"]),
    ))
    fig.update_layout(
        template="plotly_dark",
        margin=dict(l=45, r=45, t=35, b=45),
        paper_bgcolor=COLORS["panel_bg"],
        plot_bgcolor=COLORS["app_bg"],
        font=dict(color=COLORS["text"]),
        xaxis_title="Pore throat radius r (µm)",
        yaxis_title="dV/dlog(r) (normalized)",
        legend=dict(orientation="h"),
    )
    fig.update_xaxes(type="log")
    return fig

def fig_thomeer(df: pd.DataFrame, results: Dict[str, Any], params: Dict[str, Any], meta: Dict[str, Any], ui: Dict[str, Any]) -> go.Figure:
    d = df.copy()
    d = recompute_derived(d, params)

    pc = d["Pressure"].to_numpy(dtype=float)
    # compute vb series as in fit
    phi_pct = params.get("phi_override_pct") or meta.get("porosity_pct")
    vb_label = results.get("thomeer_vb_label", "Vb (proxy)")
    try:
        phi_frac = float(phi_pct) / 100.0
        if not np.isfinite(phi_frac) or phi_frac <= 0:
            raise ValueError
        vb = d["HgSat"].to_numpy(dtype=float) * phi_frac
    except Exception:
        vb = d["HgSat"].to_numpy(dtype=float)

    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=pc, y=vb, mode="markers",
        name="Data",
        marker=dict(color=COLORS["curve2"], size=6),
    ))

    # fit curve if available
    mode = (results or {}).get("thomeer_mode", "unimodal")

    # Bi-modal fit (macro + micro)
    if (
        mode == "bimodal"
        and all(k in results for k in ["thomeer_pd1_psia", "thomeer_G1", "thomeer_vb_inf1", "thomeer_pd2_psia", "thomeer_G2", "thomeer_vb_inf2"])
    ):
        pd1 = float(results["thomeer_pd1_psia"])
        G1 = float(results["thomeer_G1"])
        vb1 = float(results["thomeer_vb_inf1"])
        pd2 = float(results["thomeer_pd2_psia"])
        G2 = float(results["thomeer_G2"])
        vb2 = float(results["thomeer_vb_inf2"])

        pc_grid = np.logspace(np.log10(max(1e-6, np.nanmin(pc[pc > 0]))), np.log10(np.nanmax(pc)), 250)
        vb_fit1 = thomeer_vb(pc_grid, vb1, pd1, G1)
        vb_fit2 = thomeer_vb(pc_grid, vb2, pd2, G2)
        vb_fit = vb_fit1 + vb_fit2

        fig.add_trace(go.Scatter(
            x=pc_grid, y=vb_fit, mode="lines",
            name="Thomeer fit (bi-modal)",
            line=dict(color=COLORS["curve3"], width=2),
        ))
        fig.add_trace(go.Scatter(
            x=pc_grid, y=vb_fit1, mode="lines",
            name="Macro component",
            line=dict(color=COLORS["curve1"], width=1, dash="dot"),
        ))
        fig.add_trace(go.Scatter(
            x=pc_grid, y=vb_fit2, mode="lines",
            name="Micro component",
            line=dict(color=COLORS["accent"], width=1, dash="dot"),
        ))

    # Uni-modal fit (single Thomeer)
    elif all(k in results for k in ["thomeer_vb_inf", "thomeer_pd_psia", "thomeer_G"]):
        vb_inf = float(results["thomeer_vb_inf"])
        pd_psia = float(results["thomeer_pd_psia"])
        G = float(results["thomeer_G"])

        pc_grid = np.logspace(np.log10(max(1e-6, np.nanmin(pc[pc > 0]))), np.log10(np.nanmax(pc)), 200)
        vb_fit = thomeer_vb(pc_grid, vb_inf, pd_psia, G)
        fig.add_trace(go.Scatter(
            x=pc_grid, y=vb_fit, mode="lines",
            name="Thomeer fit",
            line=dict(color=COLORS["curve3"], width=2),
        ))

    fig.update_layout(
        template="plotly_dark",
        margin=dict(l=45, r=45, t=35, b=45),
        paper_bgcolor=COLORS["panel_bg"],
        plot_bgcolor=COLORS["app_bg"],
        font=dict(color=COLORS["text"]),
        xaxis_title="Pressure (psia)",
        yaxis_title=vb_label,
        legend=dict(orientation="h"),
    )
    fig.update_xaxes(type="log" if ui.get("xlog", True) else "linear")
    return fig

def fig_shf(df: pd.DataFrame, params: dict, ui: dict) -> go.Figure:
    """Saturation Height Function for a single sample.

    Uses df columns: Sw, Height_m (computed from Pc_res and density contrast).
    UI option: ui['shf_axis'] in {'height_m','height_ft','depth_m','depth_ft'}.
    """
    ui = ui or {}
    params = params or DEFAULT_PARAMS

    axis = (ui.get("shf_axis") or "height_m").strip().lower()

    fig = go.Figure()
    if df is None or df.empty:
        fig.update_layout(title="Saturation Height Function (no data)", template="plotly_dark")
        return fig

    if "Sw" not in df.columns or "Height_m" not in df.columns:
        fig.update_layout(title="Saturation Height Function (missing Sw/Height)", template="plotly_dark")
        return fig

    sw = np.asarray(df["Sw"].values, dtype=float)
    h_m = np.asarray(df["Height_m"].values, dtype=float)
    m = np.isfinite(sw) & np.isfinite(h_m)
    sw = sw[m]
    h_m = h_m[m]
    if sw.size < 3:
        fig.update_layout(title="Saturation Height Function (insufficient points)", template="plotly_dark")
        return fig

    fwl = _safe_float(params.get("fwl_depth_m", None))

    note = None
    if axis == "height_ft":
        x = h_m * 3.28084
        x_title = "Height above FWL (ft)"
    elif axis == "depth_m":
        if fwl is None:
            x = h_m
            x_title = "Height above FWL (m)"
            note = "Depth axis requested, but FWL depth is not set. Open IFT/Angle Params and set FWL depth."
        else:
            x = (fwl + h_m) if fwl < 0 else (fwl - h_m)
            x_title = "Depth (m)"
    elif axis == "depth_ft":
        if fwl is None:
            x = h_m * 3.28084
            x_title = "Height above FWL (ft)"
            note = "Depth axis requested, but FWL depth is not set. Open IFT/Angle Params and set FWL depth."
        else:
            depth_m = (fwl + h_m) if fwl < 0 else (fwl - h_m)
            x = depth_m * 3.28084
            x_title = "Depth (ft)"
    else:
        x = h_m
        x_title = "Height above FWL (m)"

    # Sort by x (so the curve looks clean)
    order = np.argsort(x)
    x = x[order]
    sw = sw[order]

    fig.add_trace(go.Scatter(x=x, y=sw, mode="lines+markers", name="Sw (field)"))

    fig.update_layout(
        title="Saturation Height Function (SHF)",
        xaxis_title=x_title,
        yaxis_title="Sw (fraction)",
        template="plotly_dark",
        margin=dict(l=40, r=20, t=60, b=50),
    )
    fig.update_yaxes(range=[0, 1])

    if note:
        fig.add_annotation(
            text=note,
            xref="paper",
            yref="paper",
            x=0.01,
            y=0.99,
            showarrow=False,
            align="left",
            font=dict(size=12),
            bgcolor="rgba(0,0,0,0.35)",
            bordercolor="rgba(255,255,255,0.25)",
            borderwidth=1,
        )

    return fig
def fig_winland_crossplot(library: List[Dict[str, Any]]) -> go.Figure:
    # Scatter of k vs r35, colored by rock type / cluster if present
    pts = []
    for s in library:
        res = s.get("results", {}) or {}
        meta = s.get("meta", {}) or {}
        sid = s.get("sample_id", s.get("filename", "sample"))
        r35 = res.get("r35_um")
        k = meta.get("permeability_md") or res.get("k_swanson_md") or res.get("k_winland_md")
        rt = res.get("rock_type", "RT?")
        cl = res.get("cluster")
        if r35 is None or k is None:
            continue
        try:
            r35_f = float(r35)
            k_f = float(k)
        except Exception:
            continue
        if not (np.isfinite(r35_f) and r35_f > 0 and np.isfinite(k_f) and k_f > 0):
            continue
        pts.append((sid, r35_f, k_f, rt, cl))

    fig = go.Figure()
    if pts:
        dfp = pd.DataFrame(pts, columns=["Sample", "r35_um", "k_md", "RockType", "Cluster"])
        # Use cluster if available, else RockType
        color_col = "Cluster" if dfp["Cluster"].notna().any() else "RockType"
        groups = sorted(dfp[color_col].dropna().unique(), key=lambda x: str(x))
        palette = [COLORS["curve2"], COLORS["curve1"], COLORS["curve3"], COLORS["accent"], "#AAAAAA"]

        for gi, g in enumerate(groups):
            sub = dfp[dfp[color_col] == g]
            fig.add_trace(go.Scatter(
                x=sub["r35_um"],
                y=sub["k_md"],
                mode="markers+text",
                text=sub["Sample"],
                textposition="top center",
                name=f"{color_col}: {g}",
                marker=dict(size=10, color=palette[gi % len(palette)], line=dict(width=0)),
            ))

    fig.update_layout(
        template="plotly_dark",
        margin=dict(l=60, r=45, t=35, b=55),
        paper_bgcolor=COLORS["panel_bg"],
        plot_bgcolor=COLORS["app_bg"],
        font=dict(color=COLORS["text"]),
        xaxis_title="r35 (µm)",
        yaxis_title="Permeability k (mD)",
        legend=dict(orientation="h"),
    )
    fig.update_xaxes(type="log")
    fig.update_yaxes(type="log")
    return fig

# ---------------------------
# PDF report
# ---------------------------
def _mpl_setup_dark(ax):
    ax.set_facecolor(COLORS["app_bg"])
    ax.figure.set_facecolor(COLORS["panel_bg"])
    ax.tick_params(colors=COLORS["text"])
    for spine in ax.spines.values():
        spine.set_color("#555")
    ax.xaxis.label.set_color(COLORS["text"])
    ax.yaxis.label.set_color(COLORS["text"])
    ax.title.set_color(COLORS["text"])
    ax.grid(True, alpha=0.25)

def _save_mpl_png(fig, dpi=140) -> bytes:
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=dpi, bbox_inches="tight")
    plt.close(fig)
    return buf.getvalue()

def _plot_pc_sw_png(df: pd.DataFrame) -> bytes:
    fig, ax = plt.subplots(figsize=(6.8, 3.4))
    _mpl_setup_dark(ax)
    ax.plot(df["Pressure"], df["Sw"], marker="o", linewidth=1.7, markersize=3.5, color=COLORS["curve1"])
    ax.set_xscale("log")
    ax.set_ylim(0, 1)
    ax.set_xlabel("Pressure (psia)")
    ax.set_ylabel("Sw (fraction)")
    ax.set_title("Pc vs Sw")
    return _save_mpl_png(fig)

def _plot_psd_png(df: pd.DataFrame) -> bytes:
    fig, ax = plt.subplots(figsize=(6.8, 3.4))
    _mpl_setup_dark(ax)
    ax.plot(df["r_um"], df["dVdlogr_norm"], linewidth=1.7, color=COLORS["curve2"])
    ax.set_xscale("log")
    ax.set_xlabel("Pore throat radius r (µm)")
    ax.set_ylabel("dV/dlog(r) (normalized)")
    ax.set_title("Pore Size Distribution (normalized)")
    return _save_mpl_png(fig)

def _plot_thomeer_png(df: pd.DataFrame, results: Dict[str, Any], params: Dict[str, Any], meta: Dict[str, Any]) -> bytes:
    d = recompute_derived(df.copy(), params)
    pc = d["Pressure"].to_numpy(dtype=float)

    # vb series
    phi_pct = params.get("phi_override_pct") or meta.get("porosity_pct")
    try:
        phi_frac = float(phi_pct) / 100.0
        if not np.isfinite(phi_frac) or phi_frac <= 0:
            raise ValueError
        vb = d["HgSat"].to_numpy(dtype=float) * phi_frac
        vb_label = "Vb (bulk fraction)"
    except Exception:
        vb = d["HgSat"].to_numpy(dtype=float)
        vb_label = "HgSat (proxy)"

    fig, ax = plt.subplots(figsize=(6.8, 3.4))
    _mpl_setup_dark(ax)
    ax.scatter(pc, vb, s=14, color=COLORS["curve2"], label="Data")
    ax.set_xscale("log")
    ax.set_xlabel("Pressure (psia)")
    ax.set_ylabel(vb_label)
    ax.set_title("Thomeer Fit")

    if all(k in results for k in ["thomeer_vb_inf", "thomeer_pd_psia", "thomeer_G"]):
        vb_inf = float(results["thomeer_vb_inf"])
        pd_psia = float(results["thomeer_pd_psia"])
        G = float(results["thomeer_G"])
        pc_grid = np.logspace(np.log10(max(1e-6, np.nanmin(pc[pc > 0]))), np.log10(np.nanmax(pc)), 200)
        vb_fit = thomeer_vb(pc_grid, vb_inf, pd_psia, G)
        ax.plot(pc_grid, vb_fit, color=COLORS["curve3"], linewidth=2.0, label="Fit")
        ax.legend(facecolor=COLORS["panel_bg"], edgecolor="#666", labelcolor=COLORS["text"])

    return _save_mpl_png(fig)

def _plot_shf_png(df: pd.DataFrame) -> bytes:
    fig, ax = plt.subplots(figsize=(6.8, 3.4))
    _mpl_setup_dark(ax)
    ax.plot(df["Height_m"], df["Sw"], marker="o", linewidth=1.7, markersize=3.5, color=COLORS["curve1"])
    ax.set_ylim(0, 1)
    ax.set_xlabel("Height (m)")
    ax.set_ylabel("Sw (fraction)")
    ax.set_title("Saturation Height Function (SHF)")
    return _save_mpl_png(fig)

def build_pdf_report_bytes(sample: Dict[str, Any], params: Dict[str, Any], library: List[Dict[str, Any]]) -> bytes:
    df = pd.DataFrame(sample.get("data", []))
    df = _ensure_schema(df)
    df = recompute_derived(df, params)

    meta = sample.get("meta", {}) or {}
    res = sample.get("results", {}) or {}
    sid = sample.get("sample_id", sample.get("filename", "sample"))
    well = sample.get("well", meta.get("well_guess", ""))
    filename = sample.get("filename", "")

    # Generate plot PNGs
    pc_sw_png = _plot_pc_sw_png(df)
    psd_png = _plot_psd_png(df)
    th_png = _plot_thomeer_png(df, res, params, meta)
    shf_png = _plot_shf_png(df)

    # Build report
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=letter, leftMargin=0.65 * inch, rightMargin=0.65 * inch, topMargin=0.6 * inch, bottomMargin=0.6 * inch)
    styles = getSampleStyleSheet()
    styleN = styles["Normal"]
    styleH = styles["Heading1"]
    styleH2 = styles["Heading2"]

    story = []
    story.append(Paragraph(f"<b>GeoPore Analytics — MICP Report</b>", styleH))
    story.append(Paragraph(f"Version: {VERSION}", styleN))
    story.append(Paragraph(f"Well: <b>{well}</b> &nbsp;&nbsp; Sample: <b>{sid}</b>", styleN))
    if filename:
        story.append(Paragraph(f"Source file: {filename}", styleN))
    story.append(Paragraph(f"Generated: {_dt.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", styleN))
    story.append(Spacer(1, 0.2 * inch))

    # Meta table
    meta_rows = [
        ["Metric", "Value"],
        ["Porosity (%)", f"{meta.get('porosity_pct', '')}"],
        ["Permeability (mD)", f"{meta.get('permeability_md', '')}"],
        ["Bulk density (g/mL)", f"{meta.get('bulk_density_g_ml', '')}"],
        ["Apparent/Skeletal density (g/mL)", f"{meta.get('skeletal_density_g_ml', '')}"],
        ["Stem volume (mL)", f"{meta.get('stem_volume_ml', '')}"],
        ["Stem volume used (%)", f"{meta.get('stem_volume_used_pct', '')}"],
        ["Tortuosity", f"{meta.get('tortuosity', '')}"],
        ["Tortuosity factor", f"{meta.get('tortuosity_factor', '')}"],
        ["Formation factor", f"{meta.get('formation_factor', '')}"],
        ["Threshold Pressure (psia)", f"{res.get('threshold_pressure_psia_used', meta.get('threshold_pressure_psia', ''))}"],
        ["Backbone / Fractal P (psia)", f"{res.get('backbone_pressure_psia', '')}"],
        ["Threshold P method", f"{res.get('threshold_pressure_method', '')}"],
        ["Petro QC grade", f"{res.get('petro_qc_grade', '')}"],
        ["Petro QC recommendation", f"{res.get('petro_qc_recommendation', '')}"],
        ["Sample decision", f"{res.get('sample_decision', 'PENDING')}"],
        ["Discard reason", f"{res.get('discard_reason', '')}"],
    ]
    tbl_style = TableStyle([

        ("BACKGROUND", (0, 0), (-1, 0), rl_colors.HexColor(COLORS["panel_bg"])),
        ("TEXTCOLOR", (0, 0), (-1, 0), rl_colors.HexColor(COLORS["text"])),
        ("GRID", (0, 0), (-1, -1), 0.25, rl_colors.HexColor("#555555")),
        ("BACKGROUND", (0, 1), (-1, -1), rl_colors.HexColor(COLORS["app_bg"])),
        ("TEXTCOLOR", (0, 1), (-1, -1), rl_colors.HexColor(COLORS["text"])),
        ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
        ("FONTSIZE", (0, 0), (-1, -1), 9),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ])
    t = Table(meta_rows, colWidths=[2.2 * inch, 3.8 * inch])
    t.setStyle(tbl_style)
    story.append(Paragraph("<b>Metadata</b>", styleH2))
    story.append(t)
    story.append(Spacer(1, 0.2 * inch))

    # Results table
    mode_th = res.get("thomeer_mode")
    res_rows = [
        ["Metric", "Value"],
        ["Winland mode", f"{res.get('winland_mode', '')}"],
        ["r35 used (µm)", f"{res.get('r35_um', '')}"],
        ["r35 total (µm)", f"{res.get('r35_um_total', '')}"],
        ["r35 macro-norm (µm)", f"{res.get('r35_um_macro', '')}"],
        ["Rock Type", f"{res.get('rock_type', '')}"],
        ["k_swanson (mD)", f"{res.get('k_swanson_md', '')}"],
        ["k_winland used (mD)", f"{res.get('k_winland_md', '')}"],
        ["k_winland total (mD)", f"{res.get('k_winland_md_total', '')}"],
        ["k_winland macro (mD)", f"{res.get('k_winland_md_macro', '')}"],
        ["Swanson Sp", f"{res.get('swanson_sp', '')}"],
        ["Thomeer mode", f"{mode_th or ''}"],
    ]

    if mode_th == "bimodal":
        res_rows += [
            ["Thomeer Pd1 (psia)", f"{res.get('thomeer_pd1_psia', '')}"],
            ["Thomeer G1", f"{res.get('thomeer_G1', '')}"],
            ["Thomeer Vb∞1", f"{res.get('thomeer_vb_inf1', '')}"],
            ["Thomeer Pd2 (psia)", f"{res.get('thomeer_pd2_psia', '')}"],
            ["Thomeer G2", f"{res.get('thomeer_G2', '')}"],
            ["Thomeer Vb∞2", f"{res.get('thomeer_vb_inf2', '')}"],
            ["Thomeer macro frac", f"{res.get('thomeer_macro_frac', '')}"],
            ["Thomeer R²", f"{res.get('thomeer_r2', '')}"],
        ]
    else:
        res_rows += [
            ["Thomeer Pd (psia)", f"{res.get('thomeer_pd_psia', '')}"],
            ["Thomeer G", f"{res.get('thomeer_G', '')}"],
            ["Thomeer Vb∞", f"{res.get('thomeer_vb_inf', '')}"],
            ["Thomeer R²", f"{res.get('thomeer_r2', '')}"],
        ]

    t2 = Table(res_rows, colWidths=[2.2 * inch, 3.8 * inch])
    t2.setStyle(tbl_style)
    story.append(Paragraph("<b>Key Results</b>", styleH2))
    story.append(t2)
    story.append(Spacer(1, 0.2 * inch))

    # Figures
    def add_img(title: str, png_bytes: bytes):
        story.append(Paragraph(f"<b>{title}</b>", styleH2))
        img = RLImage(io.BytesIO(png_bytes))
        img.drawHeight = 3.0 * inch
        img.drawWidth = 6.8 * inch
        story.append(img)
        story.append(Spacer(1, 0.2 * inch))

    add_img("Pc vs Sw", pc_sw_png)
    add_img("PSD — dV/dlog(r) (normalized)", psd_png)
    add_img("Thomeer Fit", th_png)
    add_img("SHF — Sw vs Height", shf_png)

    # Optional: multi-sample summary table
    if len(library) > 1:
        story.append(Paragraph("<b>Multi‑sample Summary</b>", styleH2))
        rows = [["Sample", "Decision", "PetroQC", "r35 (µm)", "k (mD)", "RockType", "Cluster"]]
        for s in library:
            sid2 = s.get("sample_id", s.get("filename", "sample"))
            res2 = s.get("results", {}) or {}
            meta2 = s.get("meta", {}) or {}
            r35 = res2.get("r35_um", "")
            k = meta2.get("permeability_md") or res2.get("k_swanson_md") or res2.get("k_winland_md") or ""
            rt = res2.get("rock_type", "")
            cl = res2.get("cluster", "")
            rows.append([sid2, str(res2.get("sample_decision", "PENDING")), str(res2.get("petro_qc_grade", "")), str(r35), str(k), str(rt), str(cl)])
        t3 = Table(rows, colWidths=[2.0 * inch, 0.9 * inch, 0.8 * inch, 0.9 * inch, 0.9 * inch, 1.4 * inch, 0.6 * inch])
        t3.setStyle(tbl_style)
        story.append(t3)

    doc.build(story)
    return buf.getvalue()

# ---------------------------
# ---------------------------
# Multi-sample Figures
# ---------------------------

def _iter_samples(library):
    """Yield sample dicts from the in-memory library."""
    if not library:
        return []
    out = []
    for s in library:
        if isinstance(s, dict) and s.get("data"):
            out.append(s)
    return out



def _coerce_library_list(library):
    """Coerce the `store-library` payload into a list of sample dicts.

    Across versions, `store-library` has held:
      - a list of sample dicts (current/preferred)
      - a project/session dict containing a `library` key (older versions / imported storage)
      - occasionally a JSON string (e.g., corrupted localStorage)

    The remove-sample workflow needs a predictable list.
    """

    if library is None:
        return []

    # tolerate JSON string payloads
    if isinstance(library, str):
        try:
            library = json.loads(library)
        except Exception:
            return []

    # project dict wrapper
    if isinstance(library, dict):
        if isinstance(library.get("library"), list):
            return [s for s in library.get("library", []) if isinstance(s, dict)]
        if isinstance(library.get("samples"), list):
            return [s for s in library.get("samples", []) if isinstance(s, dict)]
        # a single sample dict
        if "sample_id" in library or "id" in library:
            return [library]
        # best-effort: first list field
        for k in ("data", "items", "rows"):
            v = library.get(k)
            if isinstance(v, list):
                return [s for s in v if isinstance(s, dict)]
        return []

    if isinstance(library, (list, tuple)):
        out = []
        for it in library:
            if it is None:
                continue
            if isinstance(it, str):
                try:
                    it = json.loads(it)
                except Exception:
                    continue
            if isinstance(it, dict):
                out.append(it)
        return out

    return []


def _df_from_sample(sample):
    try:
        df = pd.DataFrame(sample.get("data", []))
        if df is None:
            return pd.DataFrame()
        return df
    except Exception:
        return pd.DataFrame()


def _sample_label(sample):
    """Return a human-friendly label for legends/hover.

    Library items store `sample_id` and `filename` (not `id`).
    """
    if not sample:
        return "sample"

    # Prefer the canonical identifiers used by this app
    for key in ("sample_id", "label", "id", "filename"):
        v = sample.get(key) if isinstance(sample, dict) else None
        if v is None:
            continue
        if isinstance(v, str) and v.strip() == "":
            continue
        return str(v)

    meta = sample.get("meta", {}) if isinstance(sample, dict) else {}
    if isinstance(meta, dict):
        for key in ("sample_id", "Sample ID", "SampleID"):
            v = meta.get(key)
            if v is None:
                continue
            if isinstance(v, str) and v.strip() == "":
                continue
            return str(v)

    return "sample"




def _is_excluded_sample_for_multisample(sample, params):
    """Return True if a sample should be hidden from multi-sample plots by default.

    Rules:
    - User decision DISCARDED -> excluded
    - PetroQC hard-fail (grade FAIL) -> excluded
    - If PetroQC results are not present yet, we attempt a lightweight PetroQC evaluation
      (best effort; failures will default to *not* excluding).
    """
    try:
        if not isinstance(sample, dict):
            return False
        res = sample.get('results', {}) or {}
        if isinstance(res, dict):
            dec = (res.get('sample_decision') or '').strip().upper()
            if dec == 'DISCARDED':
                return True
            if bool(res.get('exclude_from_shm')):
                return True
            grade = (res.get('petro_qc_grade') or '').strip().upper()
            if grade == 'FAIL':
                return True

        # If PetroQC has not been run yet, attempt a best-effort evaluation
        meta = sample.get('meta', {}) or {}
        if not isinstance(meta, dict):
            meta = {}
        df = _df_from_sample(sample)
        if df is None or df.empty:
            return False
        df = _ensure_schema(df)
        res_tmp = dict(res) if isinstance(res, dict) else {}
        try:
            petrophysical_qaqc(meta, params or {}, df=df, res=res_tmp)
        except Exception:
            return False
        grade2 = (res_tmp.get('petro_qc_grade') or '').strip().upper()
        return grade2 == 'FAIL'
    except Exception:
        return False

def _select_k_md(meta, res):
    """Pick a representative permeability (mD) for plots (core if present, else MICP-derived)."""
    # prefer measured/core permeability
    k_core = meta.get("permeability_md")
    try:
        if k_core is not None and float(k_core) > 0:
            return float(k_core), "core"
    except Exception:
        pass

    # then Swanson
    k_sw = res.get("k_swanson_md")
    try:
        if k_sw is not None and float(k_sw) > 0:
            return float(k_sw), "swanson"
    except Exception:
        pass

    # then Winland (macro if available, else total)
    for key, tag in [
        ("k_winland_macro_md", "winland_macro"),
        ("k_winland_md", "winland"),
    ]:
        try:
            v = res.get(key)
            if v is not None and float(v) > 0:
                return float(v), tag
        except Exception:
            continue

    return None, "n/a"


def fig_ms_pc_overlay(library_ms, ui, highlight_id=None):
    """Multi-sample Pc vs Sw overlay."""
    ui = ui or {}
    samples = _iter_samples(library_ms)
    fig = go.Figure()

    xlog = bool(ui.get("xlog", True))

    for s in samples:
        df = _df_from_sample(s)
        if df.empty:
            continue
        try:
            d = df.copy()

            # Ensure derived columns exist
            if "Sw" not in d.columns or "HgSat" not in d.columns:
                d = recompute_derived(d, DEFAULT_PARAMS)

            # Use intrusion branch only: remove points flagged as pressure decreasing
            # (Flag_Pressure_Down is set to 'Y' when P_i < P_{i-1}).
            if "Flag_Pressure_Down" in d.columns:
                f = d["Flag_Pressure_Down"].fillna("").astype(str).str.strip().str.upper()
                d = d[f != "Y"]

            if "Pressure" not in d.columns or "Sw" not in d.columns:
                continue

            d = d.sort_values("Pressure")
            x = d["Pressure"].astype(float).values
            y = d["Sw"].astype(float).values
            if len(x) < 2:
                continue

            name = _sample_label(s)
            is_hi = (highlight_id is not None and s.get("sample_id") == highlight_id)

            fig.add_trace(
                go.Scatter(
                    x=x,
                    y=y,
                    mode="lines",
                    name=name,
                    line=dict(width=3 if is_hi else 1),
                    opacity=1.0 if is_hi else 0.65,
                    hovertemplate="Sample=%{fullData.name}<br>P=%{x:.3g} psia<br>Sw=%{y:.3f}<extra></extra>",
                )
            )
        except Exception:
            continue

    fig.update_layout(
        template="plotly_dark",
        margin=dict(l=50, r=25, t=45, b=120),
        legend=dict(orientation="h", yanchor="top", y=-0.22, xanchor="left", x=0),
        title="Multi-Sample Capillary Pressure Overlay (Pc vs Sw)",
    )
    fig.update_xaxes(title="Pressure (psia)", type="log" if xlog else "linear")
    fig.update_yaxes(title="Sw (fraction)", range=[0, 1])
    return fig


def fig_ms_cum_intrusion(library_ms, ui, highlight_id=None):
    """Multi-sample cumulative intrusion vs pressure."""
    ui = ui or {}
    samples = _iter_samples(library_ms)
    fig = go.Figure()
    xlog = bool(ui.get("xlog", True))

    for s in samples:
        df = _df_from_sample(s)
        if df.empty:
            continue
        try:
            d = df.copy()

            # Use intrusion branch only (exclude pressure-down / extrusion branch)
            if "Flag_Pressure_Down" in d.columns:
                f = d["Flag_Pressure_Down"].fillna("").astype(str).str.strip().str.upper()
                d = d[f != "Y"]

            if "Pressure" not in d.columns or "CumVol" not in d.columns:
                continue

            d = d.sort_values("Pressure")
            x = d["Pressure"].astype(float).values
            y = d["CumVol"].astype(float).values
            if len(x) < 2:
                continue

            name = _sample_label(s)
            is_hi = (highlight_id is not None and s.get("sample_id") == highlight_id)
            fig.add_trace(
                go.Scatter(
                    x=x,
                    y=y,
                    mode="lines",
                    name=name,
                    line=dict(width=3 if is_hi else 1),
                    opacity=1.0 if is_hi else 0.65,
                    hovertemplate="Sample=%{fullData.name}<br>P=%{x:.3g} psia<br>Cum=%{y:.4g} mL/g<extra></extra>",
                )
            )
        except Exception:
            continue

    fig.update_layout(
        template="plotly_dark",
        margin=dict(l=50, r=25, t=45, b=120),
        legend=dict(orientation="h", yanchor="top", y=-0.22, xanchor="left", x=0),
        title="Multi-Sample Cumulative Intrusion Profile",
    )
    fig.update_xaxes(title="Pressure (psia)", type="log" if xlog else "linear")
    fig.update_yaxes(title="Cumulative Intrusion (mL/g)")
    return fig


def fig_ms_psd_compare(library_ms, ui, highlight_id=None):
    """Multi-sample PSD (dV/dlog(r)) comparison."""
    ui = ui or {}
    samples = _iter_samples(library_ms)
    fig = go.Figure()
    xlog = bool(ui.get("xlog", True))

    for s in samples:
        df = _df_from_sample(s)
        if df.empty:
            continue
        try:
            d = df.copy()
            if "dVdlogr_norm" not in d.columns or "r_um" not in d.columns:
                d = recompute_derived(d, DEFAULT_PARAMS)
            if "r_um" not in d.columns or "dVdlogr_norm" not in d.columns:
                continue
            x = d["r_um"].astype(float).values
            y = d["dVdlogr_norm"].astype(float).values
            name = _sample_label(s)
            is_hi = (highlight_id is not None and s.get("sample_id") == highlight_id)
            fig.add_trace(
                go.Scatter(
                    x=x,
                    y=y,
                    mode="lines",
                    name=name,
                    line=dict(width=3 if is_hi else 1),
                    opacity=1.0 if is_hi else 0.65,
                    hovertemplate="Sample=%{fullData.name}<br>r=%{x:.4g} µm<br>dV/dlog(r)=%{y:.4g}<extra></extra>",
                )
            )
        except Exception:
            continue

    fig.update_layout(
        template="plotly_dark",
        margin=dict(l=55, r=25, t=45, b=120),
        legend=dict(orientation="h", yanchor="top", y=-0.22, xanchor="left", x=0),
        title="Pore Throat Size Distribution (PSD) Comparison",
    )
    fig.update_xaxes(title="Pore throat radius (µm)", type="log" if xlog else "linear")
    fig.update_yaxes(title="Normalized dV/dlog(r)")
    return fig


def fig_ms_phi_k_crossplot(library_ms, params, highlight_id=None):
    """Porosity vs Permeability (multi-sample)."""
    params = params or DEFAULT_PARAMS
    samples = _iter_samples(library_ms)

    rows = []
    for s in samples:
        meta = s.get("meta", {}) or {}
        res = s.get("results", {}) or {}
        phi = meta.get("porosity_pct")
        try:
            phi = float(phi) if phi is not None else None
        except Exception:
            phi = None

        k, ktag = _select_k_md(meta, res)
        if phi is None or k is None:
            continue

        rows.append(
            dict(
                id=s.get("sample_id"),
                label=_sample_label(s),
                porosity_pct=phi,
                k_md=k,
                k_method=ktag,
                rock_type=res.get("rock_type") or meta.get("rock_type"),
            )
        )

    dfp = pd.DataFrame(rows)
    fig = go.Figure()
    if dfp.empty:
        fig.update_layout(
            template="plotly_dark",
            title="Porosity vs. Permeability Crossplot (no data)",
            margin=dict(l=50, r=25, t=45, b=120),
        )
        return fig

    # One trace per sample so the legend shows sample names (industry-style QC).
    for _, r in dfp.iterrows():
        sid = r.get("id")
        label = r.get("label")
        size = 14 if (highlight_id is not None and sid == highlight_id) else 9
        fig.add_trace(
            go.Scatter(
                x=[r.get("porosity_pct")],
                y=[r.get("k_md")],
                mode="markers",
                name=label,
                marker=dict(size=size),
                customdata=[r.get("k_method")],
                hovertemplate=(
                    "Sample=%{fullData.name}<br>"
                    "ϕ=%{x:.2f}%<br>"
                    "k=%{y:.3g} mD<br>"
                    "k_method=%{customdata}<extra></extra>"
                ),
                showlegend=True,
            )
        )

    fig.update_layout(
        template="plotly_dark",
        title="Porosity vs. Permeability Crossplot",
        margin=dict(l=55, r=25, t=45, b=140),
        legend=dict(
            orientation="h",
            yanchor="top",
            y=-0.22,
            xanchor="left",
            x=0.0,
            font=dict(size=10),
        ),
    )
    fig.update_xaxes(title="Porosity (%)")
    fig.update_yaxes(title="Permeability (mD)", type="log")
    return fig


def fig_ms_j_function(library_ms, params, highlight_id=None):
    """Leverett J-function plot (multi-sample)."""
    params = params or DEFAULT_PARAMS
    samples = _iter_samples(library_ms)
    fig = go.Figure()

    for s in samples:
        df = _df_from_sample(s)
        if df.empty:
            continue
        meta = s.get("meta", {}) or {}
        res = s.get("results", {}) or {}

        # Compute J for this sample
        try:
            d = recompute_derived(df.copy(), params)
            if d.empty or "Sw" not in d.columns:
                continue

            phi_pct = meta.get("porosity_pct")
            phi = float(phi_pct) / 100.0 if phi_pct is not None else None
            if phi is None or phi <= 0 or phi >= 1:
                continue

            k_md, _ = _select_k_md(meta, res)
            if k_md is None or k_md <= 0:
                continue

            sigma = float(params.get("sigma_res_Nm", 0.03))
            theta = math.radians(float(params.get("theta_res_deg", 140.0)))
            cos_t = math.cos(theta) if abs(math.cos(theta)) > 1e-12 else 1e-12

            # Pc at reservoir conditions (Pa) stored as Pc_res_pa
            if "Pc_res_pa" not in d.columns:
                continue

            J = d["Pc_res_pa"].astype(float).values * math.sqrt(k_md * 9.869233e-16 / phi) / (sigma * cos_t)
            Sw = d["Sw"].astype(float).values

            name = _sample_label(s)
            is_hi = (highlight_id is not None and s.get("sample_id") == highlight_id)
            fig.add_trace(
                go.Scatter(
                    x=Sw,
                    y=J,
                    mode="lines",
                    name=name,
                    line=dict(width=3 if is_hi else 1),
                    opacity=1.0 if is_hi else 0.65,
                    hovertemplate="Sample=%{fullData.name}<br>Sw=%{x:.3f}<br>J=%{y:.3g}<extra></extra>",
                )
            )
        except Exception:
            continue

    fig.update_layout(
        template="plotly_dark",
        title="Leverett J-Function (Multi-Sample)",
        margin=dict(l=55, r=25, t=45, b=120),
        legend=dict(orientation="h", yanchor="top", y=-0.22, xanchor="left", x=0),
    )
    fig.update_xaxes(title="Sw (fraction)", range=[0, 1])
    fig.update_yaxes(title="J (dimensionless)")
    return fig


def _safe_float(x):
    try:
        if x is None or (isinstance(x, str) and x.strip() == ""):
            return None
        return float(x)
    except Exception:
        return None


def _sample_phi_k(sample: dict):
    """Return (phi_frac, k_md) for a sample using whatever metadata is available."""
    meta = sample.get("meta", {}) or {}
    res = sample.get("results", {}) or {}

    # Porosity is stored in % in this app; convert to fraction for calculations
    phi_pct = (
        _safe_float(meta.get("porosity_pct"))
        or _safe_float(res.get("phi_pct_used"))
        or _safe_float(res.get("porosity_pct"))
    )
    phi = (phi_pct / 100.0) if (phi_pct is not None) else None

    # Permeability in mD
    k_md = (
        _safe_float(meta.get("permeability_md"))
        or _safe_float(res.get("permeability_md"))
        or _safe_float(res.get("perm_md"))
        or _safe_float(res.get("k_air_md"))
        or _safe_float(res.get("k_swanson_md"))
        or _safe_float(res.get("k_winland_md"))
        or _safe_float(res.get("k_pnm_md"))
    )

    return phi, k_md


def _compute_sample_j_curve(sample: dict, params: dict):
    """Return (Sw, J) arrays for a single sample, or (None, None) if unavailable."""
    params = params or DEFAULT_PARAMS
    phi, k_md = _sample_phi_k(sample)
    if phi is None or k_md is None or phi <= 0 or k_md <= 0:
        return None, None

    # NOTE: the app stores the MICP table as list-of-dicts in "data" (processed)
    # and "data_raw" (unprocessed). Earlier versions used the key "raw".
    # SHM/J-curves require the actual MICP table.
    records = sample.get("data", None)
    if records is None:
        records = sample.get("data_raw", None)
    if records is None:
        records = sample.get("raw", None)

    df = pd.DataFrame(records) if records is not None else pd.DataFrame()
    if df.empty:
        return None, None

    d = recompute_derived(df.copy(), params)
    if d is None or d.empty:
        return None, None

    if "Sw" not in d.columns or "Pc_res_pa" not in d.columns:
        return None, None

    sw = np.asarray(d["Sw"].values, dtype=float)
    pc_res_pa = np.asarray(d["Pc_res_pa"].values, dtype=float)

    sigma = float(params.get("sigma_res_npm") or DEFAULT_PARAMS.get("sigma_res_npm") or 0.03)
    theta_deg = float(params.get("theta_res_deg") or DEFAULT_PARAMS.get("theta_res_deg") or 30.0)
    cos_t = math.cos(math.radians(theta_deg))
    cos_t = abs(cos_t) if abs(cos_t) > 1e-9 else 1e-9

    # mD -> m^2
    k_m2 = float(k_md) * 9.869233e-16

    j = pc_res_pa * np.sqrt(k_m2 / float(phi)) / (sigma * cos_t)

    # Clean / sort by Sw ascending
    m = np.isfinite(sw) & np.isfinite(j)
    sw = sw[m]
    j = j[m]
    if sw.size < 4:
        return None, None

    order = np.argsort(sw)
    sw = sw[order]
    j = j[order]

    # Drop duplicate Sw (keep first)
    sw_unique, idx_unique = np.unique(sw, return_index=True)
    j_unique = j[idx_unique]

    if sw_unique.size < 4:
        return None, None

    return sw_unique, j_unique
def build_shm_pseudo_curves(library: dict, params: dict, n_points: int = 41, sw_min: float = 0.05, sw_max: float = 0.98):
    """Build pseudo (average) SHM curves per Rock Type.

    Returns a dict:
      {rock_type: {sw: [...], j: [...], height_m: [...], height_ft: [...], n: int, phi_rep: float, k_rep_md: float}}
    """
    params = params or DEFAULT_PARAMS
    lib = (library or {}) if isinstance(library, dict) else {}

    sw_grid = np.linspace(float(sw_min), float(sw_max), int(n_points))

    # group by rock type if available
    groups = {}
    for sid, s in lib.items():
        if not isinstance(s, dict):
            continue
        res = s.get("results", {}) or {}
        rt = res.get("rock_type") or res.get("cluster") or "All"
        rt = str(rt)
        groups.setdefault(rt, []).append(s)

    out = {}
    for rt, samples in groups.items():
        j_stack = []
        phis = []
        ks = []
        used = 0

        for s in samples:
            sw, j = _compute_sample_j_curve(s, params)
            if sw is None or j is None:
                continue

            # Interpolate on Sw grid, but only within sample range
            j_interp = np.full_like(sw_grid, np.nan, dtype=float)
            m = (sw_grid >= sw.min()) & (sw_grid <= sw.max())
            if np.any(m):
                j_interp[m] = np.interp(sw_grid[m], sw, j)

            if np.isfinite(j_interp).sum() < 5:
                continue

            j_stack.append(j_interp)
            phi, k_md = _sample_phi_k(s)
            if phi is not None:
                phis.append(phi)
            if k_md is not None:
                ks.append(k_md)
            used += 1

        if used < 1 or len(j_stack) < 1:
            continue

        j_arr = np.vstack(j_stack)
        j_med = np.nanmedian(j_arr, axis=0)

        # representative phi/k for this RT (median of available)
        phi_rep = float(np.nanmedian(phis)) if len(phis) else None
        k_rep_md = float(np.nanmedian(ks)) if len(ks) else None
        if phi_rep is None or k_rep_md is None or phi_rep <= 0 or k_rep_md <= 0:
            continue

        sigma = float(params.get("sigma_res_npm") or DEFAULT_PARAMS.get("sigma_res_npm") or 0.03)
        theta_deg = float(params.get("theta_res_deg") or DEFAULT_PARAMS.get("theta_res_deg") or 30.0)
        cos_t = math.cos(math.radians(theta_deg))
        cos_t = abs(cos_t) if abs(cos_t) > 1e-9 else 1e-9

        k_m2 = float(k_rep_md) * 9.869233e-16
        pc_res_pa = j_med * (sigma * cos_t) / np.sqrt(k_m2 / float(phi_rep))

        rho_w = float(params.get("rho_w_kgm3") or DEFAULT_PARAMS.get("rho_w_kgm3") or 1000.0)
        rho_hc = float(params.get("rho_hc_kgm3") or DEFAULT_PARAMS.get("rho_hc_kgm3") or 800.0)
        delta_rho = max(rho_w - rho_hc, 1e-6)
        g = 9.80665

        height_m = pc_res_pa / (delta_rho * g)
        height_ft = height_m * 3.28084

        out[rt] = {
            "sw": sw_grid.tolist(),
            "j": j_med.tolist(),
            "height_m": height_m.tolist(),
            "height_ft": height_ft.tolist(),
            "n": int(used),
            "phi_rep": phi_rep,
            "k_rep_md": k_rep_md,
        }

    return out


def fig_ms_shm_curves(library: dict, params: dict, ui: dict):
    """Pseudo SHM curves (Sw vs Height/Depth) per Rock Type."""
    ui = ui or {}
    params = params or DEFAULT_PARAMS
    axis = (ui.get("shf_axis") or "height_m").strip().lower()

    curves = build_shm_pseudo_curves(library, params, n_points=51)

    fig = go.Figure()
    if not curves:
        fig.update_layout(title="SHM Pseudo-curves (no rock types / insufficient data)", template="plotly_dark")
        return fig

    # Optional: absolute depth reference (FWL)
    fwl = params.get("fwl_depth_m", None)
    fwl = _safe_float(fwl)

    for rt, c in curves.items():
        sw = np.asarray(c.get("sw", []), dtype=float)
        h_m = np.asarray(c.get("height_m", []), dtype=float)
        h_ft = np.asarray(c.get("height_ft", []), dtype=float)

        if sw.size < 3:
            continue

        x = None
        x_title = ""

        if axis == "height_ft":
            x = h_ft
            x_title = "Height above FWL (ft)"
        elif axis == "depth_m" and fwl is not None:
            depth_m = (fwl + h_m) if fwl < 0 else (fwl - h_m)
            x = depth_m
            x_title = "Depth (m)"
        elif axis == "depth_ft" and fwl is not None:
            depth_m = (fwl + h_m) if fwl < 0 else (fwl - h_m)
            x = depth_m * 3.28084
            x_title = "Depth (ft)"
        else:
            x = h_m
            x_title = "Height above FWL (m)"

        # Sort by x increasing for a clean curve
        m = np.isfinite(x) & np.isfinite(sw)
        x = x[m]
        sw2 = sw[m]
        if x.size < 3:
            continue
        o = np.argsort(x)
        x = x[o]
        sw2 = sw2[o]

        phi_rep = c.get("phi_rep", None)
        k_rep_md = c.get("k_rep_md", None)
        if phi_rep is not None and k_rep_md is not None:
            label = f"{rt} (n={c.get('n', 0)}, φ~{phi_rep*100:.1f}%, k~{k_rep_md:.1f} mD)"
        else:
            label = f"{rt} (n={c.get('n', 0)})"

        fig.add_trace(go.Scatter(x=x, y=sw2, mode="lines", name=label))

    fig.update_layout(
        title="Saturation Height Modeling — Pseudo-curves by Rock Type",
        xaxis_title=x_title,
        yaxis_title="Sw (fraction)",
        template="plotly_dark",
        legend=dict(orientation="h", yanchor="bottom", y=-0.22, xanchor="left", x=0),
        margin=dict(l=40, r=20, t=60, b=90),
    )
    fig.update_yaxes(range=[0, 1])
    return fig
def fig_ms_g_vs_pd(library_ms, highlight_id=None):
    """Pore geometry factor (G) vs Pd crossplot."""
    samples = _iter_samples(library_ms)
    rows = []

    for s in samples:
        res = s.get("results", {}) or {}
        sid = s.get("sample_id")
        label = _sample_label(s)

        mode = res.get("thomeer_mode", "unimodal")
        if mode == "bimodal":
            for tag in ["macro", "micro"]:
                pd_key = "thomeer_pd1_psia" if tag == "macro" else "thomeer_pd2_psia"
                g_key = "thomeer_G1" if tag == "macro" else "thomeer_G2"
                pdv = res.get(pd_key)
                gv = res.get(g_key)
                try:
                    if pdv is None or gv is None:
                        continue
                    pdv = float(pdv)
                    gv = float(gv)
                    if pdv <= 0:
                        continue
                    rows.append(
                        dict(
                            id=sid,
                            label=label,
                            pd_psia=pdv,
                            G=gv,
                            component=tag,
                        )
                    )
                except Exception:
                    continue
        else:
            try:
                pdv = res.get("thomeer_pd_psia")
                gv = res.get("thomeer_G")
                if pdv is None or gv is None:
                    continue
                pdv = float(pdv)
                gv = float(gv)
                if pdv <= 0:
                    continue
                rows.append(dict(id=sid, label=label, pd_psia=pdv, G=gv, component="uni"))
            except Exception:
                continue

    dfp = pd.DataFrame(rows)
    fig = go.Figure()
    if dfp.empty:
        fig.update_layout(
            template="plotly_dark",
            title="Pore Geometry Factor (G) vs Pd (no Thomeer data)",
            margin=dict(l=55, r=25, t=45, b=120),
        )
        return fig

    # One trace per sample so the legend shows sample names.
    # We keep component information via marker symbols + hover.
    sym_map = {"uni": "circle", "macro": "diamond", "micro": "square"}
    for sid, sub in dfp.groupby("id"):
        try:
            label = str(sub["label"].iloc[0])
        except Exception:
            label = str(sid)

        comps = sub["component"].tolist()
        symbols = [sym_map.get(c, "circle") for c in comps]
        size = 14 if (highlight_id is not None and sid == highlight_id) else 9

        fig.add_trace(
            go.Scatter(
                x=sub["pd_psia"],
                y=sub["G"],
                mode="markers",
                name=label,
                marker=dict(symbol=symbols, size=size),
                customdata=comps,
                hovertemplate=(
                    "Sample=%{fullData.name}<br>"
                    "comp=%{customdata}<br>"
                    "Pd=%{x:.3g} psia<br>"
                    "G=%{y:.3g}<extra></extra>"
                ),
                showlegend=True,
            )
        )

    fig.update_layout(
        template="plotly_dark",
        title="Pore Geometry Factor (G) vs Pd",
        margin=dict(l=55, r=25, t=45, b=140),
        legend=dict(
            orientation="h",
            yanchor="top",
            y=-0.22,
            xanchor="left",
            x=0.0,
            font=dict(size=10),
        ),
    )
    fig.update_xaxes(title="Pd (psia)", type="log")
    fig.update_yaxes(title="G (dimensionless)")
    return fig


def fig_ms_petro_logs(library_ms, params):
    """Multi-sample petrophysical log with DEPTH on Y (increasing downward)."""
    params = params or DEFAULT_PARAMS
    samples = _iter_samples(library_ms)

    rows = []
    for s in samples:
        meta = s.get("meta", {}) or {}
        res = s.get("results", {}) or {}

        depth = meta.get("depth_m")
        try:
            depth = float(depth) if depth is not None else None
        except Exception:
            depth = None
        if depth is None:
            continue

        phi = meta.get("porosity_pct")
        try:
            phi = float(phi) if phi is not None else None
        except Exception:
            phi = None

        k_md, ktag = _select_k_md(meta, res)

        rows.append(
            dict(
                depth_m=depth,
                phi_pct=phi,
                k_md=k_md,
                k_method=ktag,
                label=_sample_label(s),
            )
        )

    dfp = pd.DataFrame(rows)
    if dfp.empty:
        fig = go.Figure()
        fig.update_layout(
            template="plotly_dark",
            title="Petrophysical Property Logs (no depth data)",
            margin=dict(l=60, r=60, t=45, b=45),
        )
        return fig

    dfp = dfp.sort_values("depth_m")

    fig = go.Figure()

    # Porosity track (bottom x-axis)
    if dfp["phi_pct"].notna().any():
        fig.add_trace(
            go.Scatter(
                x=dfp["phi_pct"],
                y=dfp["depth_m"],
                mode="markers+lines",
                name="Porosity (%)",
                text=dfp["label"],
                hovertemplate="Sample=%{text}<br>Depth=%{y:.2f} m<br>ϕ=%{x:.2f}%<extra></extra>",
            )
        )

    # Permeability track (top x-axis, log)
    if dfp["k_md"].notna().any():
        fig.add_trace(
            go.Scatter(
                x=dfp["k_md"],
                y=dfp["depth_m"],
                mode="markers+lines",
                name="Permeability (mD)",
                text=dfp["label"],
                xaxis="x2",
                hovertemplate="Sample=%{text}<br>Depth=%{y:.2f} m<br>k=%{x:.3g} mD<extra></extra>",
            )
        )

    fig.update_layout(
        template="plotly_dark",
        title="Petrophysical Property Logs (Depth)",
        margin=dict(l=60, r=60, t=55, b=55),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0),
        yaxis=dict(title="Depth (m)", autorange="reversed"),
        xaxis=dict(title="Porosity (%)", side="bottom"),
        xaxis2=dict(
            title="Permeability (mD)",
            overlaying="x",
            side="top",
            type="log",
            showgrid=False,
        ),
    )
    return fig




def _collect_ms_k_profiles(library, logn4_store=None):
    """Build a depth-indexed DataFrame with permeability/porosity profiles across the current library.

    This is used by the Multi-Sample 'k Profile' plot (well-log style). It intentionally tolerates
    partially-populated sample dictionaries and will return NaNs for missing fields.
    """
    library = _coerce_library_list(library)

    rows = []
    for s in library:
        if not isinstance(s, dict):
            continue

        meta = s.get("meta", {}) or {}
        res = s.get("results", {}) or {}

        # Reuse PetroQC row builder for consistent 'used' values where available
        try:
            qc_row = build_petroqc_row(s)
        except Exception:
            qc_row = {}

        depth_m = _safe_float(qc_row.get("depth_m", meta.get("depth_m")))
        if depth_m is None or (isinstance(depth_m, float) and np.isnan(depth_m)):
            # If depth is unknown, skip the point (a log track without depth is not meaningful)
            continue

        por_pct = _safe_float(qc_row.get("porosity_pct", res.get("phi_pct_used", meta.get("porosity_pct"))))
        phi_frac = (por_pct / 100.0) if (por_pct is not None and not np.isnan(por_pct)) else np.nan

        # Measured (core / lab) k
        k_air_md = _safe_float(qc_row.get("permeability_md", res.get("k_air_md_used", meta.get("permeability_md"))))

        # Model/predicted permeabilities (if computed)
        k_pnm_md = _safe_float(res.get("k_pnm_md"))
        k_pnm_fast_md = _safe_float(res.get("k_pnm_fast_md") or res.get("k_pnm_md_fast") or res.get("k_pnm_fast"))
        k_swanson_md = _safe_float(res.get("k_swanson_md"))
        k_winland_md = _safe_float(
            res.get("k_winland_md")
            or res.get("k_winland_md_total")
            or res.get("k_winland_md_macro")
        )

        # Optional external/core permeability variants (if you import them elsewhere)
        k_klinkenberg_md = _safe_float(
            res.get("k_klinkenberg_md")
            or meta.get("k_klinkenberg_md")
            or meta.get("klinkenberg_md")
        )
        k_helium_md = _safe_float(
            res.get("k_helium_md")
            or meta.get("k_helium_md")
            or meta.get("helium_md")
        )

        # External Core Logs permeabilities (if stored in meta/results)
        permamb_md = _safe_float(
            res.get("permamb_md")
            or meta.get("permamb_md")
            or meta.get("PermAmb")
            or meta.get("Perm_Amb")
        )
        permob_md = _safe_float(
            res.get("permob_md")
            or meta.get("permob_md")
            or meta.get("PermOB")
            or meta.get("Perm_OB")
        )

        # Best-effort enrichment from External Core Logs (e.g., Helium / Klinkenberg permeability sheets)
        depth_ft = _safe_float(meta.get("depth_ft"))
        if depth_ft is None and depth_m is not None:
            # Meta depth is commonly stored in meters; external logs are typically in feet
            try:
                depth_ft = float(depth_m) / 0.3048
            except Exception:
                depth_ft = None
        ext = _best_external_perm_at_depth(logn4_store, depth_ft) if logn4_store else {}
        if ext:
            if permamb_md is None:
                permamb_md = ext.get("permamb_md")
            if permob_md is None:
                permob_md = ext.get("permob_md")
            if k_air_md is None:
                k_air_md = ext.get("k_air_md")
            if k_klinkenberg_md is None:
                k_klinkenberg_md = ext.get("k_klinkenberg_md")
            if k_helium_md is None:
                k_helium_md = ext.get("k_helium_md") or ext.get("k_air_md")

        rows.append(
            {
                "depth_m": depth_m,
                "sample_id": _sample_label(s),
                "phi_frac": phi_frac,
                "k_pnm_md": k_pnm_md,
                "k_pnm_fast_md": k_pnm_fast_md,
                "k_swanson_md": k_swanson_md,
                "k_winland_md": k_winland_md,
                "k_air_md": k_air_md,
                "k_klinkenberg_md": k_klinkenberg_md,
                "k_helium_md": k_helium_md,
                "permamb_md": permamb_md,
                "permob_md": permob_md,
            }
        )

    df = pd.DataFrame(rows)
    if df.empty:
        return df

    # Depth-sorted so lines don't zig-zag
    df = df.sort_values("depth_m").reset_index(drop=True)


    # ---- column aliases (keeps plotting code stable across versions) ----
    if "k_core_md" not in df.columns and "k_air_md" in df.columns:
        df["k_core_md"] = df["k_air_md"]
    if "k_permamb_md" not in df.columns and "permamb_md" in df.columns:
        df["k_permamb_md"] = df["permamb_md"]
    if "k_permob_md" not in df.columns and "permob_md" in df.columns:
        df["k_permob_md"] = df["permob_md"]
    if "k_winland_macro_md" not in df.columns and "k_winland_md" in df.columns:
        df["k_winland_macro_md"] = df["k_winland_md"]
    return df

def fig_ms_k_profile(library, mode: str = "logk", fill_models: bool = False, logn4_store: dict | None = None):
    """Multi-sample k profile rendered as well-log style tracks.

    Tracks (default):
      - Measured k (core / Helium / Klinkenberg / N4 Perm)
      - Empirical models (Swanson / Winland)
      - PNM (PNM k)

    Notes
    -----
    - If mode == 'logk': model traces are drawn as lines (visual log(k) interpolation on log-x axis).
    - If mode == 'raw': everything is shown as raw points (markers).
    - Depth gaps larger than `kprof_gap_break_m` (from UI store when present) are broken so lines don't
      artificially connect across missing intervals.
    """

    profiles = _collect_ms_k_profiles(library, logn4_store=logn4_store)

    if profiles is None or profiles.empty:
        return _empty_fig("No k profile data available.")

    # --- Column map (DataFrame column names) ---
    keymap = {
        "Core k": "k_core_md",
        "Helium k": "k_helium_md",
        "Klinkenberg k": "k_klinkenberg_md",
        "PermOB (N4)": "permob_md",
        "PermAmb (N4)": "permamb_md",
        "PermOB (SC)": "permob_sc_md",
        "PermAmb (SC)": "permamb_sc_md",
        "Swanson": "k_swanson_md",
        "Winland (macro)": "k_winland_macro_md",
        "PNM k": "k_pnm_md",
    }

    # Track grouping (readability-first)
    track_defs = [
        ("Measured", ["Core k", "Helium k", "Klinkenberg k", "PermOB (N4)", "PermAmb (N4)", "PermOB (SC)", "PermAmb (SC)"]),
        ("Empirical models", ["Swanson", "Winland (macro)"]),
        ("PNM", ["PNM k"]),
    ]

    # Determine available series
    available_labels = []
    for label, col in keymap.items():
        if col in profiles.columns and profiles[col].notna().any():
            available_labels.append(label)

    if not available_labels:
        return _empty_fig("No k profile data available.")

    # Filter tracks to only those with at least one available series
    tracks = []
    for tname, labels in track_defs:
        present = [lbl for lbl in labels if lbl in available_labels]
        if present:
            tracks.append((tname, present))

    if not tracks:
        return _empty_fig("No k profile data available.")

    # Range for log-x axis
    k_cols_present = [keymap[lbl] for lbl in available_labels if keymap[lbl] in profiles.columns]
    all_vals = []
    for c in k_cols_present:
        v = profiles[c]
        if v is not None:
            all_vals.append(v)

    import numpy as np

    if all_vals:
        vv = np.concatenate([np.asarray(s.dropna().values, dtype=float) for s in all_vals if hasattr(s, 'dropna')])
        vv = vv[np.isfinite(vv) & (vv > 0)]
        if vv.size:
            k_min = float(np.nanpercentile(vv, 1))
            k_max = float(np.nanpercentile(vv, 99))
        else:
            k_min, k_max = 0.1, 1000.0
    else:
        k_min, k_max = 0.1, 1000.0

    # Safety
    if not np.isfinite(k_min) or k_min <= 0:
        k_min = 0.1
    if not np.isfinite(k_max) or k_max <= k_min:
        k_max = max(k_min * 10.0, 1000.0)

    # decade ticks (industry standard: 0.01 / 0.1 / 1 / 10 / 100 / 1000 ...)
    d0 = int(np.floor(np.log10(k_min)))
    d1 = int(np.ceil(np.log10(k_max)))
    decades = list(range(d0, d1 + 1))
    tick_vals = [10 ** d for d in decades]

    # Depth gap breaking threshold (optional UI store)
    gap_break_m = 6.0
    try:
        if isinstance(logn4_store, dict):
            ui = (logn4_store.get('_ui', {}) or {})
            gap_break_m = float(ui.get('kprof_gap_break_m', gap_break_m))
    except Exception:
        pass

    def _break_on_gaps(y_depth, x_k, meta_text, gap_m: float):
        """Insert None breaks when depth gaps are large (for line traces)."""
        if gap_m is None or gap_m <= 0 or len(y_depth) < 2:
            return x_k, y_depth, meta_text
        xo, yo, to = [], [], []
        last_y = None
        for xi, yi, ti in zip(x_k, y_depth, meta_text):
            if last_y is not None and yi is not None and abs(float(yi) - float(last_y)) > gap_m:
                xo.append(None); yo.append(None); to.append(None)
            xo.append(xi); yo.append(yi); to.append(ti)
            last_y = yi
        return xo, yo, to

    # Styling
    style_map = {
        "Core k": dict(color="#00E5FF", width=0, symbol="circle"),
        "Helium k": dict(color="#FFD54F", width=0, symbol="diamond"),
        "Klinkenberg k": dict(color="#FFA726", width=0, symbol="diamond"),
        "PermOB (N4)": dict(color="#90CAF9", width=0, symbol="square"),
        "PermAmb (N4)": dict(color="#64B5F6", width=0, symbol="square"),
        "PermOB (SC)": dict(color="#BBDEFB", width=0, symbol="square-open"),
        "PermAmb (SC)": dict(color="#64B5F6", width=0, symbol="square-open"),
        "Swanson": dict(color="#00FF66", width=2, dash="solid"),
        "Winland (macro)": dict(color="#00FF66", width=2, dash="dash"),
        "PNM k": dict(color="#FF00FF", width=2, dash="solid"),
    }

    measured_labels = {"Core k", "Helium k", "Klinkenberg k", "PermOB (N4)", "PermAmb (N4)", "PermOB (SC)", "PermAmb (SC)"}

    # Create track subplots
    fig = make_subplots(
        rows=1,
        cols=len(tracks),
        shared_yaxes=True,
        horizontal_spacing=0.035,
        subplot_titles=[t[0] for t in tracks],
    )

    # Add traces per track
    for cidx, (tname, labels) in enumerate(tracks, start=1):
        for label in labels:
            col = keymap[label]
            df = profiles[["depth_m", "sample_id", col]].dropna()
            if df.empty:
                continue

            y = df["depth_m"].astype(float).tolist()
            x = df[col].astype(float).tolist()
            txt = df["sample_id"].astype(str).tolist()

            # Keep only positive k values for log axis
            x2, y2, t2 = [], [], []
            for xi, yi, ti in zip(x, y, txt):
                if xi is None:
                    continue
                try:
                    if float(xi) <= 0:
                        continue
                except Exception:
                    continue
                x2.append(float(xi))
                y2.append(float(yi))
                t2.append(ti)

            if not x2:
                continue

            st = style_map.get(label, {})

            is_measured = label in measured_labels
            show_line = (not is_measured) and (mode == "logk")

            if show_line:
                x2, y2, t2 = _break_on_gaps(y2, x2, t2, gap_break_m)

            hover = (
                "Sample: %{text}<br>"
                "Depth: %{y:.2f} m<br>"
                "k: %{x:.3g} mD<extra></extra>"
            )

            if is_measured:
                fig.add_trace(
                    go.Scatter(
                        x=x2,
                        y=y2,
                        mode="markers",
                        name=label,
                        text=t2,
                        hovertemplate=hover,
                        marker=dict(size=6, color=st.get("color", "#FFFFFF"), symbol=st.get("symbol", "circle"), line=dict(width=0)),
                        showlegend=True,
                    ),
                    row=1,
                    col=cidx,
                )
            else:
                fig.add_trace(
                    go.Scatter(
                        x=x2,
                        y=y2,
                        mode="lines" if show_line else "markers",
                        name=label,
                        text=t2,
                        hovertemplate=hover,
                        line=dict(color=st.get("color", "#FFFFFF"), width=int(st.get("width", 2)), dash=st.get("dash", "solid")),
                        marker=dict(size=5, color=st.get("color", "#FFFFFF")),
                        fill="tozerox" if (fill_models and show_line) else None,
                        opacity=0.22 if (fill_models and show_line) else 1.0,
                        showlegend=True,
                    ),
                    row=1,
                    col=cidx,
                )

    # Axes configuration
    x_range = [np.log10(k_min), np.log10(k_max)]

    for cidx in range(1, len(tracks) + 1):
        fig.update_xaxes(
            type="log",
            range=x_range,
            tickvals=tick_vals,
            ticktext=[f"{v:g}" for v in tick_vals],
            showgrid=True,
            minor=dict(showgrid=True),
            title_text="Permeability (mD)",
            row=1,
            col=cidx,
        )

    # y-axis on first track only (shared y)
    fig.update_yaxes(
        title_text="Depth (m)",
        autorange="reversed",
        showgrid=True,
        row=1,
        col=1,
    )

    # Hide y tick labels on other tracks (clean log-track look)
    for cidx in range(2, len(tracks) + 1):
        fig.update_yaxes(showticklabels=False, row=1, col=cidx)

    fig.update_layout(
        template="plotly_dark",
        title="k Profile (Well-log tracks)",
        height=520,
        margin=dict(l=70, r=25, t=55, b=55),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0.0),
        hovermode="closest",
    )

    return fig
def fig_ms_hfu_log(library_ms, params):
    """Hydraulic Flow Unit (HFU) log using FZI-based classes (DEPTH on Y)."""
    params = params or DEFAULT_PARAMS
    samples = _iter_samples(library_ms)

    rows = []
    logfzi_vals = []
    for s in samples:
        meta = s.get("meta", {}) or {}
        res = s.get("results", {}) or {}

        depth = meta.get("depth_m")
        try:
            depth = float(depth) if depth is not None else None
        except Exception:
            depth = None
        if depth is None:
            continue

        phi_pct = meta.get("porosity_pct")
        try:
            phi = float(phi_pct) / 100.0 if phi_pct is not None else None
        except Exception:
            phi = None
        if phi is None or phi <= 0 or phi >= 0.999:
            continue

        k_md, _ = _select_k_md(meta, res)
        if k_md is None or k_md <= 0:
            continue

        try:
            rqi = 0.0314 * math.sqrt(k_md / phi)
            phiz = phi / (1.0 - phi)
            fzi = rqi / phiz if phiz > 0 else None
            if fzi is None or fzi <= 0:
                continue
            logfzi = math.log10(fzi)
        except Exception:
            continue

        logfzi_vals.append(logfzi)
        rows.append(
            dict(
                depth_m=depth,
                logFZI=logfzi,
                FZI=fzi,
                label=_sample_label(s),
            )
        )

    dfp = pd.DataFrame(rows)
    fig = go.Figure()
    if dfp.empty or len(logfzi_vals) < 2:
        fig.update_layout(
            template="plotly_dark",
            title="Hydraulic Flow Unit (HFU) Log (insufficient data)",
            margin=dict(l=60, r=30, t=45, b=45),
        )
        return fig

    dfp = dfp.sort_values("depth_m")

    qs = np.quantile(np.array(logfzi_vals), [0.2, 0.4, 0.6, 0.8])
    hfu = 1 + np.digitize(dfp["logFZI"].values, qs)
    dfp["HFU"] = hfu.astype(int)

    fig.add_trace(
        go.Scatter(
            x=dfp["HFU"],
            y=dfp["depth_m"],
            mode="markers+text",
            marker=dict(size=10, color=dfp["HFU"], colorscale="Turbo", showscale=True, colorbar=dict(title="HFU")),
            text=dfp["label"],
            textposition="top center",
            textfont=dict(size=10),
            customdata=np.stack([dfp["FZI"].values, dfp["logFZI"].values], axis=-1),
            hovertemplate="Sample=%{text}<br>Depth=%{y:.2f} m<br>HFU=%{x}<br>FZI=%{customdata[0]:.3g}<br>logFZI=%{customdata[1]:.3g}<extra></extra>",
            name="HFU",
        )
    )

    fig.update_layout(
        template="plotly_dark",
        title="Hydraulic Flow Unit (HFU) Log (FZI-based)",
        margin=dict(l=60, r=30, t=45, b=45),
    )
    fig.update_xaxes(title="HFU (class)", dtick=1)
    fig.update_yaxes(title="Depth (m)", autorange="reversed")
    return fig



# Dash UI
# ---------------------------
external_stylesheets = [dbc.themes.CYBORG]

app = Dash(
    __name__,
    external_stylesheets=external_stylesheets,
    meta_tags=[{"name": "viewport", "content": "width=device-width, initial-scale=0.75"}],  # 75% scaling
)

app.title = f"GeoPore Analytics — MICP v{VERSION}"

# Custom CSS (palette + 75% fallback zoom for some browsers)
CUSTOM_CSS = """
:root {{
  --app-bg: {COLORS["app_bg"]};
  --panel-bg: {COLORS["panel_bg"]};
  --accent: {COLORS["accent"]};
  --text: {COLORS["text"]};
}}
body {{
  background-color: var(--app-bg);
  color: var(--text);
  font-size: 19px;
  line-height: 1.35;
}}
/* Fallback zoom on desktop browsers that ignore viewport initial-scale */
#app-root {{
  zoom: 0.75;
}}
/* Right control panel: independent scroll so the plot stays static while you browse buttons */
.control-sidebar {{
  /* Compensate zoom (0.75) so the sidebar visually fills the viewport */
  height: calc(133.333vh - 80px);
  overflow-y: auto;
  padding-right: 10px;
  padding-bottom: 10px;
  scrollbar-gutter: stable;
}}}}
.control-sidebar::-webkit-scrollbar {{
  width: 10px;
}}
.control-sidebar::-webkit-scrollbar-thumb {{
  background: rgba(255,255,255,0.16);
  border-radius: 10px;
}}
.control-sidebar::-webkit-scrollbar-thumb:hover {{
  background: rgba(255,255,255,0.26);
}}

.card {{
  background-color: var(--panel-bg) !important;
  border: 1px solid rgba(255,255,255,0.08);
}}
.section-title {{
  font-weight: 700;
  letter-spacing: 0.4px;
  text-transform: none;
  color: var(--text);
  margin-bottom: 0.35rem;
}}
.btn-accent {{
  background-color: var(--accent) !important;
  border-color: var(--accent) !important;
  color: #FFFFFF !important;
  font-weight: 800;
  text-shadow: 0 1px 0 rgba(0,0,0,0.35);
}}
.btn-accent:hover {{
  filter: brightness(0.95);
}}
.btn-outline-accent {{
  border-color: var(--accent) !important;
  color: #FFFFFF !important;
  font-weight: 800;
  text-shadow: 0 1px 0 rgba(0,0,0,0.35);
}}
.btn-outline-accent:hover {{
  background-color: rgba(0,210,255,0.18) !important;
  color: #FFFFFF !important;
}}

.action-btn {{
  width: 100%;
  min-height: 40px;
  font-size: 0.95rem;
  display: flex !important;
  align-items: center;
  justify-content: center;
  text-align: center;
  padding: 0.35rem 0.60rem;
  border-radius: 0.40rem;
  white-space: nowrap;
  overflow: hidden;
  text-overflow: ellipsis;
}}
.topbar-btn {{
  min-height: 36px;
  min-width: 82px;
  padding: 0.30rem 0.80rem;
}}

.topbar-actions {{
  display: flex;
  gap: 10px;
  align-items: center;
  justify-content: flex-end;
  flex-wrap: wrap;
}}

.small-muted {{
  color: rgba(224,224,224,0.7);
  font-size: 0.9rem;
}}

.ag-theme-alpine-dark, .ag-theme-alpine {{
  font-size: 16px;
}}


/* Inputs / selects on dark background */
.form-control, .form-select, .Select-control, .Select-menu-outer, textarea {
  background-color: rgba(0,0,0,0.18) !important;
  color: var(--text) !important;
  border-color: rgba(255,255,255,0.15) !important;
}

/* --- Fix contrast for the Sample selector (native <select>) --- */
#sel-sample, #sel-sample.form-select {
  background-color: #0b0b0b !important;
  color: #ffffff !important;
}
#sel-sample option {
  background-color: #0b0b0b !important;
  color: #ffffff !important;
}

/* --- k Profile mode selector (buttons) --- */
.radio-kprof label {
  display: inline-flex;
  align-items: center;
  padding: 6px 10px;
  border: 1px solid rgba(255,255,255,0.18);
  border-radius: 10px;
  background: rgba(0,0,0,0.18);
  margin-right: 8px;
  cursor: pointer;
  user-select: none;
}
.radio-kprof input { margin-right: 6px; }
.radio-kprof label:hover { border-color: rgba(0,210,255,0.60); }
.radio-kprof input:checked + span { font-weight: 600; color: #ffffff; }

.form-control:focus, .form-select:focus, textarea:focus {
  box-shadow: 0 0 0 0.2rem rgba(0,210,255,0.15) !important;
  border-color: rgba(0,210,255,0.55) !important;
}

/* Workflow tracker */
.wf-steps {
  display: flex;
  flex-wrap: wrap;
  gap: 14px;
  align-items: center;
  margin-top: 6px;
  margin-bottom: 2px;
}
.wf-step {
  display: inline-flex;
  align-items: center;
  gap: 6px;
  font-size: 0.95rem;
}
.wf-box {
  width: 12px;
  height: 12px;
  border-radius: 3px;
  border: 1px solid rgba(255,255,255,0.25);
  display: inline-block;
}
.wf-on {
  background-color: var(--accent);
  border-color: var(--accent);
}
.wf-off {
  background-color: rgba(255,255,255,0.08);
}
.wf-kpi-label {
  font-size: 0.85rem;
  color: rgba(224,224,224,0.75);
  margin-bottom: 2px;
}
.kpi-input {
  background-color: rgba(0,0,0,0.18) !important;
  color: var(--text) !important;
  border-color: rgba(255,255,255,0.16) !important;
  font-weight: 800 !important;
}
.kpi-input:disabled {
  opacity: 1 !important;
}

/* Petrophysical QAQC */
.petro-qc-box {
  background-color: rgba(0,0,0,0.18);
  border: 1px solid rgba(255,255,255,0.10);
  border-radius: 0.50rem;
  padding: 0.60rem 0.70rem;
  font-size: 0.92rem;
  max-height: 220px;
  overflow: auto;
}
.petro-qc-item {
  display: flex;
  align-items: flex-start;
  gap: 8px;
  margin-bottom: 6px;
}
.petro-qc-code {
  font-family: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, "Liberation Mono", "Courier New", monospace;
  font-size: 0.78rem;
  padding: 2px 6px;
  border-radius: 6px;
  background-color: rgba(255,255,255,0.08);
  border: 1px solid rgba(255,255,255,0.12);
  white-space: nowrap;
}
.petro-qc-msg {
  flex: 1;
  color: rgba(224,224,224,0.92);
  line-height: 1.15rem;
}

.log-area {
  background-color: rgba(0,0,0,0.18) !important;
  color: var(--text) !important;
  border-color: rgba(255,255,255,0.12) !important;
  font-family: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, "Liberation Mono", "Courier New", monospace;
  font-size: 0.88rem;
}

/* Progress bar styling */
.progress {
  background-color: rgba(255,255,255,0.10) !important;
  height: 14px;
}
.progress-bar {
  background-color: var(--accent) !important;
  color: #1E1E2F !important;
  font-weight: 900;
}

/* Thomeer slider visibility (dcc.Slider / rc-slider) */
.rc-slider-mark-text {
  color: rgba(224,224,224,0.90) !important;
  font-size: 0.85rem;
}
.rc-slider-mark-text-active {
  color: #FFFFFF !important;
}
.rc-slider-rail {
  background-color: rgba(255,255,255,0.14) !important;
}
.rc-slider-track {
  background-color: var(--accent) !important;
}
.rc-slider-handle {
  border-color: var(--accent) !important;
  background-color: var(--panel-bg) !important;
  width: 16px;
  height: 16px;
  margin-top: -6px;
}
.rc-slider-dot {
  border-color: rgba(255,255,255,0.18) !important;
}
.rc-slider-dot-active {
  border-color: var(--accent) !important;
}
.rc-slider-tooltip-inner {
  background-color: rgba(0,0,0,0.85) !important;
  color: #FFFFFF !important;
  font-weight: 900;
  font-size: 0.85rem;
}
.rc-slider-tooltip-arrow {
  border-top-color: rgba(0,0,0,0.85) !important;
}

/* Thomeer readouts */
.thomeer-readout {
  color: #FFFFFF;
  font-size: 0.95rem;
  font-weight: 800;
  margin-top: 2px;
}
"""

# Resolve palette placeholders (avoid f-string brace escaping issues)
CUSTOM_CSS = CUSTOM_CSS.replace('{COLORS["app_bg"]}', COLORS["app_bg"])
CUSTOM_CSS = CUSTOM_CSS.replace('{COLORS["panel_bg"]}', COLORS["panel_bg"])
CUSTOM_CSS = CUSTOM_CSS.replace('{COLORS["accent"]}', COLORS["accent"])
CUSTOM_CSS = CUSTOM_CSS.replace('{COLORS["text"]}', COLORS["text"])
# Convert any doubled braces ({{ / }}) back to normal CSS braces
CUSTOM_CSS = CUSTOM_CSS.replace("{{", "{").replace("}}", "}")

# Inject CUSTOM_CSS without relying on html.Style (works across Dash versions)
app.index_string = f"""
<!DOCTYPE html>
<html>
  <head>
    {{%metas%}}
    <title>{{%title%}}</title>
    {{%favicon%}}
    {{%css%}}
    <style>{CUSTOM_CSS}</style>
  </head>
  <body>
    {{%app_entry%}}
    <footer>
      {{%config%}}
      {{%scripts%}}
      {{%renderer%}}
    </footer>
  </body>
</html>
"""


def fig_ms_k_pnm_crossplot(library_ms, params, highlight_id=None):
    """Core k vs PNM k (multi-sample, log-log)."""
    params = params or DEFAULT_PARAMS
    samples = _iter_samples(library_ms)

    rows = []
    for s in samples:
        meta = s.get("meta", {}) or {}
        res = s.get("results", {}) or {}

        # core permeability
        k_core = meta.get("permeability_md")
        try:
            k_core = float(k_core) if k_core is not None else None
        except Exception:
            k_core = None

        # PNM permeability (compute on-the-fly if missing)
        k_pnm = res.get("k_pnm_md")
        try:
            k_pnm = float(k_pnm) if k_pnm is not None else None
        except Exception:
            k_pnm = None

        if (k_pnm is None or not np.isfinite(k_pnm) or k_pnm <= 0) and s.get("data"):
            try:
                df = pd.DataFrame(s.get("data", []))
                df = _ensure_schema(df)
                pnm = compute_pnm_fast(df, params, meta, res)
                if pnm.get("k_pnm_md") is not None:
                    k_pnm = float(pnm.get("k_pnm_md"))
            except Exception:
                k_pnm = None

        if k_core is None or k_pnm is None:
            continue
        if (not np.isfinite(k_core)) or (not np.isfinite(k_pnm)) or k_core <= 0 or k_pnm <= 0:
            continue

        depth = meta.get("depth_m")
        try:
            depth = float(depth) if depth is not None else None
        except Exception:
            depth = None

        rows.append(
            dict(
                id=s.get("sample_id"),
                label=_sample_label(s),
                depth_m=depth,
                k_core=k_core,
                k_pnm=k_pnm,
            )
        )

    dfp = pd.DataFrame(rows)
    fig = go.Figure()
    if dfp.empty:
        fig.update_layout(
            template="plotly_dark",
            title="Core k vs PNM k (no data)",
            margin=dict(l=55, r=25, t=45, b=45),
        )
        fig.update_xaxes(title="Core Permeability (mD)", type="log")
        fig.update_yaxes(title="PNM Permeability (mD)", type="log")
        return fig

    # marker size highlight
    sizes = []
    for sid in dfp["id"].tolist():
        sizes.append(14 if (highlight_id is not None and sid == highlight_id) else 9)

    fig.add_trace(
        go.Scatter(
            x=dfp["k_core"],
            y=dfp["k_pnm"],
            mode="markers",
            marker=dict(size=sizes),
            text=dfp["label"],
            customdata=dfp["depth_m"],
            hovertemplate="Sample=%{text}<br>Depth=%{customdata:.2f} m<br>k_core=%{x:.3g} mD<br>k_pnm=%{y:.3g} mD<extra></extra>",
            name="Samples",
        )
    )

    # 1:1 reference
    try:
        xmin = float(np.nanmin(dfp["k_core"]))
        xmax = float(np.nanmax(dfp["k_core"]))
        ymin = float(np.nanmin(dfp["k_pnm"]))
        ymax = float(np.nanmax(dfp["k_pnm"]))
        lo = max(min(xmin, ymin), 1e-6)
        hi = max(xmax, ymax)
        fig.add_trace(
            go.Scatter(
                x=[lo, hi],
                y=[lo, hi],
                mode="lines",
                name="1:1",
                line=dict(color="rgba(255,255,255,0.35)", dash="dot"),
                hoverinfo="skip",
            )
        )
    except Exception:
        pass

    fig.update_layout(
        template="plotly_dark",
        title="Core k vs PNM k (multi-sample)",
        margin=dict(l=60, r=30, t=45, b=45),
    )
    fig.update_xaxes(title="Core Permeability (mD)", type="log")
    fig.update_yaxes(title="PNM Permeability (mD)", type="log")
    return fig


def fig_ms_ci_log(library_ms, params, highlight_id=None):
    """Connectivity index log track: log10(k_PNM/k_core) vs depth (DEPTH on Y, increasing downward)."""
    params = params or DEFAULT_PARAMS
    samples = _iter_samples(library_ms)

    rows = []
    for s in samples:
        meta = s.get("meta", {}) or {}
        res = s.get("results", {}) or {}

        depth = meta.get("depth_m")
        try:
            depth = float(depth) if depth is not None else None
        except Exception:
            depth = None
        if depth is None:
            continue

        k_core = meta.get("permeability_md")
        try:
            k_core = float(k_core) if k_core is not None else None
        except Exception:
            k_core = None

        ci = res.get("pnm_ci")
        try:
            ci = float(ci) if ci is not None else None
        except Exception:
            ci = None

        # If missing, compute from PNM on-the-fly
        if (ci is None or not np.isfinite(ci) or ci <= 0) and k_core and k_core > 0 and s.get("data"):
            try:
                df = pd.DataFrame(s.get("data", []))
                df = _ensure_schema(df)
                pnm = compute_pnm_fast(df, params, meta, res)
                k_pnm = pnm.get("k_pnm_md")
                if k_pnm is not None and float(k_pnm) > 0:
                    ci = float(k_pnm) / float(k_core)
            except Exception:
                ci = None

        if ci is None or not np.isfinite(ci) or ci <= 0:
            continue

        rows.append(
            dict(
                id=s.get("sample_id"),
                label=_sample_label(s),
                depth_m=depth,
                ci=ci,
                log_ci=float(np.log10(ci)),
            )
        )

    dfp = pd.DataFrame(rows)
    fig = go.Figure()
    if dfp.empty:
        fig.update_layout(
            template="plotly_dark",
            title="PNM Connectivity Index Log (no depth/CI data)",
            margin=dict(l=60, r=30, t=45, b=45),
        )
        fig.update_xaxes(title="log10(k_PNM/k_core)")
        fig.update_yaxes(title="Depth (m)", autorange="reversed")
        return fig

    dfp = dfp.sort_values("depth_m")

    sizes = []
    for sid in dfp["id"].tolist():
        sizes.append(12 if (highlight_id is not None and sid == highlight_id) else 8)

    fig.add_trace(
        go.Scatter(
            x=dfp["log_ci"],
            y=dfp["depth_m"],
            mode="markers+lines",
            marker=dict(size=sizes),
            text=dfp["label"],
            hovertemplate="Sample=%{text}<br>Depth=%{y:.2f} m<br>CI=%{customdata:.3g}<br>log10(CI)=%{x:.2f}<extra></extra>",
            customdata=dfp["ci"],
            name="log10(CI)",
        )
    )

    # Reference line at CI=1 -> log10(CI)=0
    try:
        fig.add_shape(type="line", x0=0, x1=0, y0=0, y1=1, xref="x", yref="paper",
                      line=dict(color="rgba(255,255,255,0.25)", dash="dot"))
    except Exception:
        pass

    fig.update_layout(
        template="plotly_dark",
        title="PNM Connectivity Index Log (Depth)",
        margin=dict(l=60, r=30, t=45, b=45),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0),
    )
    fig.update_xaxes(title="log10(k_PNM/k_core)")
    fig.update_yaxes(title="Depth (m)", autorange="reversed")
    return fig



def _btn(
    label: str,
    id_: str,
    outline: bool = False,
    size: str = "sm",
    style: Optional[Dict[str, Any]] = None,
    block: bool = True,
    title: Optional[str] = None,
):
    """Helper to create consistently styled buttons.
    - block=True: full-width buttons for the side panel
    - block=False: compact buttons for the top bar
    """
    cls_base = "btn-outline-accent" if outline else "btn-accent"
    cls_extra = "action-btn" if block else "topbar-btn"
    cls = f"{cls_base} {cls_extra}"
    return dbc.Button(label, id=id_, className=cls, size=size, style=style or {}, title=title or label)

# Grid columns
COLUMN_DEFS = [
    {"headerName": "Pressure (psia)", "field": "Pressure", "type": "numericColumn", "editable": True},
    {"headerName": "Cumulative Intrusion (mL/g)", "field": "CumVol", "type": "numericColumn", "editable": True},
    {"headerName": "Incremental Intrusion (mL/g)", "field": "IncVol", "type": "numericColumn", "editable": True},

    {"headerName": "Flag_Pressure_Down", "field": "Flag_Pressure_Down", "editable": False, "width": 150},
    {"headerName": "Flag_Cum_Down", "field": "Flag_Cum_Down", "editable": False, "width": 130},
    {"headerName": "Flag_Inc_Neg_Fail", "field": "Flag_Inc_Neg_Fail", "editable": False, "width": 150},

    # Derived (hidden by default)
    {"headerName": "HgSat", "field": "HgSat", "type": "numericColumn", "editable": False, "hide": True},
    {"headerName": "Sw", "field": "Sw", "type": "numericColumn", "editable": False, "hide": True},
    {"headerName": "r (µm)", "field": "r_um", "type": "numericColumn", "editable": False, "hide": True},
    {"headerName": "dV/dlog(r) norm", "field": "dVdlogr_norm", "type": "numericColumn", "editable": False, "hide": True},
    {"headerName": "Height (m)", "field": "Height_m", "type": "numericColumn", "editable": False, "hide": True},
]


PETROQC_COLUMN_DEFS = [
    {"headerName": "ID", "field": "id", "hide": True},
    {"headerName": "WELL", "field": "well", "minWidth": 120},
    {"headerName": "CORE", "field": "core", "minWidth": 140},
    {"headerName": "Sample ID", "field": "sample_id", "minWidth": 260, "flex": 2},
    {"headerName": "DEPTH (m)", "field": "depth_m", "type": "numericColumn", "minWidth": 120},
    {"headerName": "Porosity (%)", "field": "porosity_pct", "type": "numericColumn", "minWidth": 120},
    {"headerName": "Permeability (mD)", "field": "permeability_md", "type": "numericColumn", "minWidth": 150},
    {"headerName": "*Threshold Pressure (psia)", "field": "threshold_pressure_psia", "type": "numericColumn", "minWidth": 190},
    {"headerName": "Tortuosity factor", "field": "tortuosity_factor", "type": "numericColumn", "minWidth": 150},
    {"headerName": "Bulk Density @0.58 psia (g/mL)", "field": "bulk_density_g_ml", "type": "numericColumn", "minWidth": 230},
    {"headerName": "Apparent (skeletal) Density (g/mL)", "field": "skeletal_density_g_ml", "type": "numericColumn", "minWidth": 260},
    {"headerName": "Stem Volume Used (%)", "field": "stem_volume_used_pct", "type": "numericColumn", "minWidth": 190},
    {"headerName": "Max_SHg_Saturation (%)", "field": "max_shg_sat_pct", "type": "numericColumn", "minWidth": 200},
    {"headerName": "Grain_Density_Diff (g/mL)", "field": "grain_density_diff_g_ml", "type": "numericColumn", "minWidth": 220},
    {"headerName": "Conformance_Vol_Pct (%)", "field": "conformance_vol_pct", "type": "numericColumn", "minWidth": 210},
    {"headerName": "QC Flag", "field": "qc_flag", "minWidth": 110},
    {
        "headerName": "QC Reasons",
        "field": "qc_reasons",
        "minWidth": 520,
        "flex": 3,
        "wrapText": True,
        "autoHeight": True,
        "cellStyle": {"whiteSpace": "normal", "lineHeight": "1.25"},
    },
    {"headerName": "Exclude_from_SHM", "field": "exclude_from_shm", "minWidth": 160},

    # ---- Thomeer / Bimodal diagnostics ----
    {"headerName": "Thomeer mode", "field": "thomeer_mode", "minWidth": 130},
    {"headerName": "Thomeer Pd (psia)", "field": "thomeer_pd", "minWidth": 150},
    {"headerName": "Thomeer G", "field": "thomeer_g", "minWidth": 110},
    {"headerName": "Thomeer BV (%)", "field": "thomeer_bv", "minWidth": 140},
    {"headerName": "Pd1 (psia)", "field": "thomeer_pd1", "minWidth": 120},
    {"headerName": "G1", "field": "thomeer_g1", "minWidth": 90},
    {"headerName": "BV1 (%)", "field": "thomeer_bv1", "minWidth": 110},
    {"headerName": "Pd2 (psia)", "field": "thomeer_pd2", "minWidth": 120},
    {"headerName": "G2", "field": "thomeer_g2", "minWidth": 90},
    {"headerName": "BV2 (%)", "field": "thomeer_bv2", "minWidth": 110},
    {"headerName": "Macro frac", "field": "thomeer_macro_frac", "minWidth": 120},
    {"headerName": "Bimodal QC", "field": "thomeer_bimodal_qc", "minWidth": 120},
    {"headerName": "Δlog10(Pd)", "field": "thomeer_pd_sep_log10", "minWidth": 130},
    # ---- Winland (bimodal) ----
    {"headerName": "k Winland total (mD)", "field": "k_winland_md_total", "minWidth": 175},
    {"headerName": "k Winland macro (mD)", "field": "k_winland_md_macro", "minWidth": 175},
]

EMPTY_SAMPLE = {
    "sample_id": None,
    "filename": None,
    "well": None,
    "meta": {},
    "results": {},
    "data": _ensure_schema(pd.DataFrame({"Pressure": [], "CumVol": [], "IncVol": []})).to_dict("records"),
    "created_at": None,
}


# ---------------------------
# PNM 3D (Synthetic Network Visualization)
# ---------------------------
def _pnm_psd_distribution(df: pd.DataFrame, params: Dict[str, Any]) -> Tuple[np.ndarray, np.ndarray]:
    """Return (r_mid_um, weights) from the sample incremental intrusion as a stable PSD proxy."""
    params = params or DEFAULT_PARAMS
    df = _ensure_schema(df)
    df = recompute_derived(df, params)

    if "r_um" not in df.columns or df["r_um"].isna().all():
        df = add_pore_throat_columns(df, params)

    r = df["r_um"].to_numpy(dtype=float)
    inc = df["IncVol"].to_numpy(dtype=float)
    ok = np.isfinite(r) & np.isfinite(inc) & (r > 0) & (inc >= 0)
    r = r[ok]
    inc = inc[ok]
    if len(r) < 5:
        raise ValueError("Not enough points to build PSD distribution.")

    log_r = np.log10(r)
    nbins = int(params.get("pnm_psd_bins", 40))
    nbins = max(10, min(200, nbins))
    edges = np.linspace(float(np.nanmin(log_r)), float(np.nanmax(log_r)), nbins + 1)

    bin_idx = np.digitize(log_r, edges) - 1
    bin_idx = np.clip(bin_idx, 0, nbins - 1)

    dv = np.zeros(nbins, dtype=float)
    r_mid = np.zeros(nbins, dtype=float)
    for i in range(nbins):
        mask = bin_idx == i
        r_mid[i] = 10 ** ((edges[i] + edges[i + 1]) / 2.0)
        if np.any(mask):
            dv[i] = float(np.nansum(inc[mask]))

    dv_sum = float(np.nansum(dv))
    if dv_sum <= 0:
        raise ValueError("No incremental intrusion volume available for PSD distribution.")
    w = dv / dv_sum

    # Remove empty bins for cleaner sampling
    keep = w > 0
    r_mid = r_mid[keep]
    w = w[keep]
    w = w / float(np.sum(w))
    return r_mid, w


def _pnm_split_distribution_bimodal(
    r_mid_um: np.ndarray,
    w: np.ndarray,
    r_split_um: Optional[float] = None,
) -> Tuple[Tuple[np.ndarray, np.ndarray], Tuple[np.ndarray, np.ndarray], Optional[float]]:
    """Split a PSD sampling distribution into macro and micro parts.

    We split by a radius threshold r_split_um:
      - macro: r >= r_split_um
      - micro: r <  r_split_um

    If r_split_um is missing/invalid or if the split produces an empty side, we fall back to a
    median split in log10(r) space.

    Returns:
      (r_macro, w_macro), (r_micro, w_micro), r_split_used
    """
    r_mid_um = np.asarray(r_mid_um, dtype=float)
    w = np.asarray(w, dtype=float)
    if r_mid_um.size == 0 or w.size == 0:
        return (r_mid_um, w), (r_mid_um, w), None

    # Normalize weights defensively
    w = np.clip(w, 0, None)
    wsum = float(np.sum(w))
    if wsum <= 0:
        w = np.ones_like(w) / float(len(w))
    else:
        w = w / wsum

    r_split_used: Optional[float] = None

    def _do_split(thr: float):
        mask_macro = r_mid_um >= thr
        mask_micro = ~mask_macro
        r_ma, w_ma = r_mid_um[mask_macro], w[mask_macro]
        r_mi, w_mi = r_mid_um[mask_micro], w[mask_micro]
        if len(r_ma) >= 3 and float(np.sum(w_ma)) > 0:
            w_ma = w_ma / float(np.sum(w_ma))
        if len(r_mi) >= 3 and float(np.sum(w_mi)) > 0:
            w_mi = w_mi / float(np.sum(w_mi))
        return (r_ma, w_ma), (r_mi, w_mi)

    # 1) Try provided split
    try:
        if r_split_um is not None and np.isfinite(float(r_split_um)) and float(r_split_um) > 0:
            r_split_used = float(r_split_um)
            (r_ma, w_ma), (r_mi, w_mi) = _do_split(r_split_used)
            if len(r_ma) >= 3 and len(r_mi) >= 3:
                return (r_ma, w_ma), (r_mi, w_mi), r_split_used
    except Exception:
        r_split_used = None

    # 2) Fallback: median split in log-space
    try:
        lr = np.log10(np.clip(r_mid_um, 1e-12, None))
        med = float(np.nanmedian(lr))
        r_split_used = float(10 ** med)
        (r_ma, w_ma), (r_mi, w_mi) = _do_split(r_split_used)
        if len(r_ma) >= 3 and len(r_mi) >= 3:
            return (r_ma, w_ma), (r_mi, w_mi), r_split_used
    except Exception:
        pass

    # 3) Last resort: return full distribution for both
    return (r_mid_um, w), (r_mid_um, w), r_split_used


def _pnm3d_seed(sample_id: Optional[str]) -> int:
    sid = (sample_id or "sample").strip() or "sample"
    h = hashlib.md5(sid.encode("utf-8")).hexdigest()
    return int(h[:8], 16)  # 32-bit deterministic seed


def _pnm_generate_synthetic_network(
    r_mid_um: np.ndarray,
    w: np.ndarray,
    n_nodes: int,
    z_mean: float,
    seed: int,
) -> Tuple[np.ndarray, np.ndarray, List[Tuple[int, int]], np.ndarray]:
    """Generate a simple synthetic pore network (nodes+edges) for 3D visualization.

    This is NOT a micro-CT extracted network. It is a qualitative network that respects:
      - coordination number ~ z_mean
      - throat size distribution sampled from PSD (r_mid_um, w)

    Returns:
      pos: (n,3) node positions
      pore_r: (n,) pore radii (um) (synthetic)
      edges: list of (i,j)
      throat_r: (m,) throat radii (um) for each edge
    """
    n_nodes = int(max(30, min(2000, n_nodes)))
    z_mean = float(max(1.5, min(12.0, z_mean)))

    rng = np.random.default_rng(int(seed) & 0xFFFFFFFF)

    # Node positions in unit cube
    pos = rng.random((n_nodes, 3))

    # Target degree distribution around z_mean
    deg_target = rng.poisson(lam=z_mean, size=n_nodes).astype(int)
    deg_target = np.clip(deg_target, 1, min(12, n_nodes - 1))

    # Ensure total degree is even
    if int(deg_target.sum()) % 2 == 1:
        if deg_target[0] < (n_nodes - 1):
            deg_target[0] += 1
        else:
            deg_target[0] -= 1

    # Candidate neighbors: use KDTree if available
    k_cand = min(14, n_nodes - 1)
    neighbor_list: List[List[int]] = []
    try:
        from scipy.spatial import cKDTree  # type: ignore
        tree = cKDTree(pos)
        _, idxs = tree.query(pos, k=k_cand + 1)  # includes itself
        for row in idxs:
            neighbor_list.append([int(j) for j in row[1:]])
    except Exception:
        # Fallback: brute-force nearest for small/medium n
        for i in range(n_nodes):
            d2 = np.sum((pos - pos[i]) ** 2, axis=1)
            order = np.argsort(d2)
            neighbor_list.append([int(j) for j in order[1 : k_cand + 1]])

    deg_cur = np.zeros(n_nodes, dtype=int)
    edges: List[Tuple[int, int]] = []
    edge_set = set()

    # Greedy fill using nearest neighbors
    node_order = np.argsort(-deg_target)
    for i in node_order:
        for j in neighbor_list[int(i)]:
            if deg_cur[int(i)] >= deg_target[int(i)]:
                break
            if deg_cur[int(j)] >= deg_target[int(j)]:
                continue
            a, b = (int(i), int(j)) if int(i) < int(j) else (int(j), int(i))
            if (a, b) in edge_set:
                continue
            edge_set.add((a, b))
            edges.append((a, b))
            deg_cur[int(i)] += 1
            deg_cur[int(j)] += 1

    # Random fill for any remaining degrees
    remaining = np.where(deg_cur < deg_target)[0]
    attempts = 0
    max_attempts = n_nodes * 40
    while len(remaining) > 0 and attempts < max_attempts:
        i = int(rng.choice(remaining))
        j = int(rng.integers(0, n_nodes))
        if i == j:
            attempts += 1
            continue
        if deg_cur[i] >= deg_target[i] or deg_cur[j] >= deg_target[j]:
            attempts += 1
            remaining = np.where(deg_cur < deg_target)[0]
            continue
        a, b = (i, j) if i < j else (j, i)
        if (a, b) in edge_set:
            attempts += 1
            continue
        edge_set.add((a, b))
        edges.append((a, b))
        deg_cur[i] += 1
        deg_cur[j] += 1
        attempts += 1
        remaining = np.where(deg_cur < deg_target)[0]

    # Sample throat radii for edges from PSD
    if len(edges) == 0:
        raise ValueError("Could not generate a network (no edges). Try increasing n_nodes.")

    w = np.asarray(w, dtype=float)
    w = w / float(np.sum(w))
    r_mid_um = np.asarray(r_mid_um, dtype=float)

    idx_e = rng.choice(len(r_mid_um), size=len(edges), replace=True, p=w)
    throat_r = r_mid_um[idx_e]

    # Synthetic pore radii (slightly larger than throats)
    idx_n = rng.choice(len(r_mid_um), size=n_nodes, replace=True, p=w)
    pore_r = r_mid_um[idx_n] * 2.0

    return pos, pore_r, edges, throat_r


def fig_pnm3d_network(
    df: pd.DataFrame,
    params: Dict[str, Any],
    meta: Dict[str, Any],
    res: Dict[str, Any],
    ui: Dict[str, Any],
    sample_id: Optional[str] = None,
) -> go.Figure:
    """3D synthetic pore-network visualization derived from PSD (qualitative).

    For bimodal Thomeer fits, we overlay two synthetic subnetworks:
      - Macro (backbone) system
      - Micro (matrix) system

    This is a *statistical* visualization, not a micro-CT extracted network.
    """
    params = params or DEFAULT_PARAMS
    res = res or {}
    meta = meta or {}
    ui = ui or {}

    fig = go.Figure()

    # Network settings
    n_nodes_base = int(params.get("pnm3d_nodes", 350))
    n_nodes_base = int(np.clip(n_nodes_base, 80, 1500))
    z_mean = float(res.get("pnm_z") or params.get("pnm_coordination_z", 4.0))
    seed_param = params.get("pnm3d_seed")
    seed = int(seed_param) if (seed_param is not None and str(seed_param).strip() != "") else _pnm3d_seed(sample_id)

    # Visual scaling (Plotly marker size is in px)
    sphere_scale = float(params.get("pnm3d_sphere_scale", 2.5) or 2.5)
    sphere_scale = float(np.clip(sphere_scale, 0.5, 6.0))

    # PSD -> sampling distribution
    try:
        r_mid_um, w = _pnm_psd_distribution(df, params)
    except Exception as e:
        fig.add_annotation(
            text=f"PNM 3D unavailable: {str(e)}",
            x=0.5, y=0.5, xref="paper", yref="paper",
            showarrow=False,
            font=dict(size=14),
        )
        fig = apply_plot_theme(fig, ui.get("plot_theme", "dark"))
        return fig

    # Detect bimodal mode
    mode = ""
    try:
        mode = str((res or {}).get("thomeer_mode") or "").lower()
    except Exception:
        mode = ""

    if mode == "bimodal":
        # Try to obtain a macro/micro split from Pd2 (preferred)
        r_split_um: Optional[float] = None
        pd2 = (res or {}).get("thomeer_pd2_psia")
        try:
            pd2_f = float(pd2)
        except Exception:
            pd2_f = float("nan")

        if np.isfinite(pd2_f) and pd2_f > 0:
            sigma = float(params.get("sigma_hg_air_npm", DEFAULT_PARAMS["sigma_hg_air_npm"]))
            theta = float(params.get("theta_hg_air_deg", DEFAULT_PARAMS["theta_hg_air_deg"]))
            try:
                r_split_um = float(_compute_radius_um(np.array([pd2_f], dtype=float), sigma, theta)[0])
                if not (np.isfinite(r_split_um) and r_split_um > 0):
                    r_split_um = None
            except Exception:
                r_split_um = None

        # If Pd2 split is unavailable, fallback to PSD peak separation
        if r_split_um is None:
            try:
                pk = _detect_psd_peaks(df, params)
                peaks = pk.get("peaks") or []
                if isinstance(peaks, list) and len(peaks) >= 2:
                    r1 = float(peaks[0].get("r_um"))
                    r2 = float(peaks[1].get("r_um"))
                    if np.isfinite(r1) and np.isfinite(r2) and r1 > 0 and r2 > 0:
                        r_split_um = float(math.sqrt(r1 * r2))
            except Exception:
                r_split_um = None

        # Split sampling distributions
        (r_ma, w_ma), (r_mi, w_mi), r_split_used = _pnm_split_distribution_bimodal(r_mid_um, w, r_split_um=r_split_um)

        # Node counts proportional to macro fraction (visual only)
        f = (res or {}).get("thomeer_macro_frac")
        try:
            f = float(f)
        except Exception:
            f = float("nan")
        if not (np.isfinite(f) and 0.0 < float(f) < 1.0):
            f = 0.5

        min_nodes = int(params.get("pnm3d_bimodal_min_nodes", 160))
        min_nodes = int(np.clip(min_nodes, 60, 600))

        n_macro = max(min_nodes, int(round(n_nodes_base * float(f))))
        n_micro = max(min_nodes, int(round(n_nodes_base * (1.0 - float(f)))))

        # Build subnetworks (separate RNG seeds)
        try:
            pos_M, pore_r_M, edges_M, throat_r_M = _pnm_generate_synthetic_network(
                r_ma, w_ma, n_nodes=n_macro, z_mean=z_mean, seed=int(seed)
            )
            pos_m, pore_r_m, edges_m, throat_r_m = _pnm_generate_synthetic_network(
                r_mi, w_mi, n_nodes=n_micro, z_mean=z_mean, seed=int(seed) + 1
            )
        except Exception as e:
            fig.add_annotation(
                text=f"PNM 3D (bimodal) unavailable: {str(e)}",
                x=0.5, y=0.5, xref="paper", yref="paper",
                showarrow=False,
                font=dict(size=14),
            )
            fig = apply_plot_theme(fig, ui.get("plot_theme", "dark"))
            return fig

        def _edges_xyz(pos: np.ndarray, edges: List[Tuple[int, int]]):
            xs: List[float] = []
            ys: List[float] = []
            zs: List[float] = []
            for a, b in edges:
                xa, ya, za = pos[a]
                xb, yb, zb = pos[b]
                xs += [float(xa), float(xb), None]
                ys += [float(ya), float(yb), None]
                zs += [float(za), float(zb), None]
            return xs, ys, zs

        # --- Macro edges (backbone)
        xM, yM, zM = _edges_xyz(pos_M, edges_M)
        fig.add_trace(
            go.Scatter3d(
                x=xM, y=yM, z=zM,
                mode="lines",
                name="Macro throats (backbone)",
                line=dict(color=COLORS["accent"], width=4),
                hoverinfo="skip",
                opacity=0.75,
            )
        )

        # --- Macro nodes
        prM = np.asarray(pore_r_M, dtype=float)
        prM_norm = prM / float(np.nanmax(prM)) if np.nanmax(prM) > 0 else prM
        # Marker size is in px – scale up to improve visual proportion vs edges.
        sizeM = (3.5 + 10.0 * prM_norm) * sphere_scale
        fig.add_trace(
            go.Scatter3d(
                x=pos_M[:, 0], y=pos_M[:, 1], z=pos_M[:, 2],
                mode="markers",
                name="Macro pores",
                marker=dict(size=sizeM, color=COLORS["curve2"], opacity=0.85),
                text=[f"Macro node {i}<br>pore_r≈{prM[i]:.2g} µm" for i in range(len(prM))],
                hoverinfo="text",
            )
        )

        # --- Micro edges (matrix)
        xm, ym, zm = _edges_xyz(pos_m, edges_m)
        fig.add_trace(
            go.Scatter3d(
                x=xm, y=ym, z=zm,
                mode="lines",
                name="Micro throats (matrix)",
                line=dict(color=COLORS["curve1"], width=2),
                hoverinfo="skip",
                opacity=0.35,
            )
        )

        # --- Micro nodes
        prm = np.asarray(pore_r_m, dtype=float)
        prm_norm = prm / float(np.nanmax(prm)) if np.nanmax(prm) > 0 else prm
        sizem = (2.8 + 7.0 * prm_norm) * sphere_scale
        fig.add_trace(
            go.Scatter3d(
                x=pos_m[:, 0], y=pos_m[:, 1], z=pos_m[:, 2],
                mode="markers",
                name="Micro pores",
                marker=dict(size=sizem, color=COLORS["curve3"], opacity=0.55),
                text=[f"Micro node {i}<br>pore_r≈{prm[i]:.2g} µm" for i in range(len(prm))],
                hoverinfo="text",
            )
        )

        # Toggle buttons (Both / Macro / Micro)
        fig.update_layout(
            updatemenus=[
                dict(
                    type="buttons",
                    direction="right",
                    x=0.0,
                    y=1.14,
                    xanchor="left",
                    yanchor="top",
                    showactive=True,
                    buttons=[
                        dict(label="Both", method="update", args=[{"visible": [True, True, True, True]}]),
                        dict(label="Macro only", method="update", args=[{"visible": [True, True, False, False]}]),
                        dict(label="Micro only", method="update", args=[{"visible": [False, False, True, True]}]),
                    ],
                )
            ]
        )

        title = f"PNM 3D (Bimodal) — {sample_id or 'Sample'}"
        fig.update_layout(
            title=title,
            margin=dict(l=0, r=0, t=50, b=0),
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0.0),
            scene=dict(
                xaxis=dict(showticklabels=False, title="", showgrid=False, zeroline=False),
                yaxis=dict(showticklabels=False, title="", showgrid=False, zeroline=False),
                zaxis=dict(showticklabels=False, title="", showgrid=False, zeroline=False),
                aspectmode="cube",
            ),
        )

        # Context annotation
        try:
            pd1 = res.get("thomeer_pd1_psia") or res.get("thomeer_pd_psia")
            pd2v = res.get("thomeer_pd2_psia")
            txt = (
                "Synthetic networks (visual only).\n"
                f"macro_frac≈{float(f):.2f} | Pd1≈{pd1} psia | Pd2≈{pd2v} psia"
            )
            if r_split_used is not None:
                txt += f" | split r≈{float(r_split_used):.3g} µm"
        except Exception:
            txt = "Synthetic networks (visual only)."
        fig.add_annotation(
            text=txt,
            x=0.0, y=-0.10, xref="paper", yref="paper",
            showarrow=False,
            font=dict(size=11),
            align="left",
        )

        fig = apply_plot_theme(fig, ui.get("plot_theme", "dark"))
        return fig

    # -------------------------
    # Unimodal / generic PNM 3D
    # -------------------------
    try:
        pos, pore_r, edges, throat_r = _pnm_generate_synthetic_network(r_mid_um, w, n_nodes=n_nodes_base, z_mean=z_mean, seed=seed)
    except Exception as e:
        fig.add_annotation(
            text=f"PNM 3D unavailable: {str(e)}",
            x=0.5, y=0.5, xref="paper", yref="paper",
            showarrow=False,
            font=dict(size=14),
        )
        fig = apply_plot_theme(fig, ui.get("plot_theme", "dark"))
        return fig

    # Edge styling by throat radius bins (quantiles)
    thr = np.asarray(throat_r, dtype=float)
    q = np.quantile(thr, [0.0, 0.5, 0.8, 0.95, 1.0])
    q = np.unique(q)
    if len(q) < 3:
        q = np.array([float(np.min(thr)), float(np.median(thr)), float(np.max(thr))])

    colors = [COLORS["curve1"], COLORS["curve2"], COLORS["curve3"], COLORS["accent"]]
    n_bins = max(1, min(len(colors), len(q) - 1))
    if (len(q) - 1) != n_bins:
        qs = np.linspace(0, 1, n_bins + 1)
        q = np.quantile(thr, qs)
        q = np.unique(q)
        if len(q) - 1 < n_bins:
            n_bins = max(1, len(q) - 1)

    bin_ids = np.digitize(thr, q[1:-1], right=True)
    for b in range(n_bins):
        xs: List[float] = []
        ys: List[float] = []
        zs: List[float] = []
        count = 0
        for (a, c), bid in zip(edges, bin_ids):
            if int(bid) != b:
                continue
            xa, ya, za = pos[a]
            xb, yb, zb = pos[c]
            xs += [float(xa), float(xb), None]
            ys += [float(ya), float(yb), None]
            zs += [float(za), float(zb), None]
            count += 1
        if count == 0:
            continue
        lo = q[b]
        hi = q[b + 1]
        fig.add_trace(
            go.Scatter3d(
                x=xs, y=ys, z=zs,
                mode="lines",
                name=f"Throats {lo:.2g}–{hi:.2g} µm",
                line=dict(color=colors[b % len(colors)], width=3),
                hoverinfo="skip",
                opacity=0.65,
            )
        )

    # Nodes
    pr = np.asarray(pore_r, dtype=float)
    pr_norm = pr / float(np.nanmax(pr)) if np.nanmax(pr) > 0 else pr
    # Marker size is in px – apply user-defined scale while keeping proportions.
    marker_size = (3.0 + 10.0 * pr_norm) * sphere_scale
    fig.add_trace(
        go.Scatter3d(
            x=pos[:, 0], y=pos[:, 1], z=pos[:, 2],
            mode="markers",
            name="Pores",
            marker=dict(
                size=marker_size,
                color=pr,
                colorscale="Viridis",
                opacity=0.9,
                showscale=True,
                colorbar=dict(title="Pore radius (µm)"),
            ),
            text=[f"Node {i}<br>pore_r≈{pr[i]:.2g} µm" for i in range(len(pr))],
            hoverinfo="text",
        )
    )

    title = f"PNM 3D (Synthetic) — {sample_id or 'Sample'}"
    fig.update_layout(
        title=title,
        margin=dict(l=0, r=0, t=40, b=0),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0.0),
        scene=dict(
            xaxis=dict(showticklabels=False, title="", showgrid=False, zeroline=False),
            yaxis=dict(showticklabels=False, title="", showgrid=False, zeroline=False),
            zaxis=dict(showticklabels=False, title="", showgrid=False, zeroline=False),
            aspectmode="cube",
        ),
    )

    fig.add_annotation(
        text="Synthetic network for visualization only (not micro‑CT). Use PNM Fast for k.",
        x=0.0, y=-0.08, xref="paper", yref="paper",
        showarrow=False,
        font=dict(size=11),
        align="left",
    )

    fig = apply_plot_theme(fig, ui.get("plot_theme", "dark"))
    return fig

def make_layout() -> dbc.Container:
    # Auto-recover last session (disk autosave) if available.
    # This helps if the PC restarts and the browser state is empty.
    auto_payload = _safe_read_json(AUTOSAVE_PATH, {})
    if isinstance(auto_payload, dict) and auto_payload.get("library"):
        init_library = auto_payload.get("library") or []
        init_current_id = auto_payload.get("current_id")
        if not init_current_id and init_library:
            init_current_id = init_library[0].get("sample_id")
        init_params = auto_payload.get("params") or DEFAULT_PARAMS
        init_ui = auto_payload.get("ui") or {"plot_mode": "intrusion", "xlog": True, "overlay_inc": False}
        init_ui.setdefault("plot_theme", "dark")
        init_ui.setdefault("shf_axis", "height_m")
        init_ui.setdefault("kprof_mode", "logk")
        init_ui.setdefault("kprof_fill", False)
        init_log = auto_payload.get("log") or []
        init_petroqc = auto_payload.get("petroqc") or []
        init_status = f"Recovered autosave: {AUTOSAVE_PATH}"
    else:
        init_library = []
        init_current_id = None
        init_params = DEFAULT_PARAMS
        init_ui = {
            "plot_mode": "intrusion",
            "xlog": True,
            "overlay_inc": False,
            "plot_theme": "dark",
            "shf_axis": "height_m",
            "kprof_mode": "logk",
        }
        init_log = []
        init_petroqc = []
        init_status = "Ready."

    return dbc.Container(
        id="app-root",
        fluid=True,
        style={"padding": "10px 52px 10px 14px"},
        children=[
            # Stores
            dcc.Store(id="store-workspace", data={"dir": WORKSPACE_DIR}, storage_type="local"),
            dcc.Store(id="store-project-refresh", data=0),
            dcc.Store(id="store-library", data=init_library, storage_type="local"),
            dcc.Store(id="store-current-id", data=init_current_id, storage_type="local"),
            dcc.Store(id="store-params", data=init_params, storage_type="local"),
            dcc.Store(
                id="store-ui",
                data=init_ui,
                storage_type="local",
            ),
            dcc.Store(id="store-status", data=init_status),
            dcc.Store(id="store-log", data=init_log, storage_type="local"),
            dcc.Store(id="store-petroqc", data=init_petroqc, storage_type="local"),
                dcc.Store(id="store-logn4", data=None, storage_type="local"),

            # Top bar
            dbc.Row(
                align="center",
                className="mb-2",
                children=[
                    dbc.Col(
                        html.Div([
                            html.Div("GeoPore Analytics", style={"fontSize": "22px", "fontWeight": "800"}),
                            html.Div(f"Ver. {VERSION}", className="small-muted"),
                        ]),
                        width=6,
                    ),
                    dbc.Col(
                        html.Div(
                            className="topbar-actions",
                            children=[
                                _btn("New", "btn-new", outline=True, block=False),
                                # Open project (upload)
                                dcc.Upload(
                                    id="upload-project",
                                    children=_btn("Open", "btn-open-proxy", outline=True, block=False),
                                    multiple=False,
                                    accept=".json",
                                    style={"display": "inline-block"},
                                ),
                                _btn("Save", "btn-save", outline=False, block=False),
                                _btn("Storage", "btn-storage", outline=True, block=False, title="Project Storage / Workspace"),
                                dcc.Download(id="download-project"),
            html.Div(id="autosave-dummy", style={"display": "none"}),
                            ],
                        ),
                        width="auto",
                    ),
                ],
            ),

            # Project Storage Modal
            dbc.Modal(
                id="modal-storage",
                is_open=False,
                size="lg",
                centered=True,
                scrollable=True,
                children=[
                    dbc.ModalHeader(dbc.ModalTitle("Project Storage")),
                    dbc.ModalBody(
                        [
                            html.Div("Workspace folder", className="small-muted"),
                            dbc.Row(
                                className="g-2",
                                children=[
                                    dbc.Col(
                                        dbc.Input(
                                            id="inp-workspace-dir",
                                            value=DEFAULT_WORKSPACE_DIR,
                                            placeholder="Workspace folder (e.g., C:/geopore_workspace)",
                                        ),
                                        width=9,
                                    ),
                                    dbc.Col(
                                        _btn(
                                            "Browse",
                                            "btn-browse-workspace",
                                            outline=True,
                                            block=True,
                                            title="Choose workspace folder",
                                        ),
                                        width=3,
                                    ),
                                ],
                            ),
                            html.Div(style={"height": "8px"}),
                            html.Div("Projects", className="small-muted"),
                            dbc.Row(
                                className="g-2",
                                children=[
                                    dbc.Col(
                                        dbc.Select(
                                            id="dd-project-files",
                                            options=[],
                                            value=None,
                                            style={"width": "100%"},
                                        ),
                                        width=8,
                                    ),
                                    dbc.Col(
                                        _btn(
                                            "Load",
                                            "btn-load-project",
                                            outline=False,
                                            block=True,
                                            title="Load selected project from disk",
                                        ),
                                        width=2,
                                    ),
                                    dbc.Col(
                                        _btn(
                                            "Refresh",
                                            "btn-refresh-projects",
                                            outline=True,
                                            block=True,
                                            title="Refresh project list",
                                        ),
                                        width=2,
                                    ),
                                ],
                            ),
                            html.Div(style={"height": "6px"}),
                            html.Div(id="lbl-workspace-path", className="small-muted"),
                            html.Hr(),
                            html.Div("Import data from folder", className="small-muted"),
                            dbc.Row(
                                className="g-2",
                                children=[
                                    dbc.Col(
                                        dbc.Input(
                                            id="inp-import-dir",
                                            value="",
                                            placeholder="Folder containing Excel/CSV files…",
                                        ),
                                        width=9,
                                    ),
                                    dbc.Col(
                                        _btn(
                                            "Browse",
                                            "btn-browse-import-dir",
                                            outline=True,
                                            block=True,
                                            title="Choose folder with MICP Excel files",
                                        ),
                                        width=3,
                                    ),
                                ],
                            ),
                            html.Div(style={"height": "6px"}),
                            _btn(
                                "Import Folder",
                                "btn-import-folder",
                                outline=False,
                                block=True,
                                title="Import all .xls/.xlsx/.csv files from the selected folder",
                            ),
                            html.Div(
                                "Supported: .xls, .xlsx, .csv. (Local disk import)",
                                className="small-muted",
                                style={"marginTop": "6px"},
                            ),
                        ]
                    ),
                    dbc.ModalFooter(
                        _btn("Close", "btn-storage-close", outline=True, block=False)
                    ),
                ],
            ),

            dbc.Row(
                children=[
                    # LEFT: Dashboard
                    dbc.Col(
                        width=8,
                        children=[
                            dbc.Card(
                                className="mb-2",
                                body=True,
                                children=[
                                    dbc.Row(
                                        align="center",
                                        className="mb-2",
                                        children=[
                                            dbc.Col(
                                                html.Div(
                                                    id="dashboard-title",
                                                    style={"fontSize": "18px", "fontWeight": "800"},
                                                ),
                                                width=True,
                                            ),
                                            dbc.Col(
                                                dbc.InputGroup(
            [
                dbc.InputGroupText("Sample"),
                dbc.Select(id="sel-sample", options=[], value=None),
            ],
            size="sm",
        ),
                                                width=5,
                                            ),
                                        ],
                                    ),

                                    dbc.Tabs(
                                        id="tabs-left",
                                        active_tab="tab-plot",
                                        children=[
                                            dbc.Tab(label="Plot", tab_id="tab-plot"),
                                            dbc.Tab(label="Data Grid", tab_id="tab-grid"),
                                            dbc.Tab(label="Metrics", tab_id="tab-metrics"),
                                        dbc.Tab(label="PetroQC", tab_id="tab-petroqc"),
                                        ],
                                    ),
                                    html.Div(style={"height": "10px"}),

                                    html.Div(
                                        id="tab-content",
                                        children=[
                                            html.Div(
                                                id="wrap-plot",
                                                children=[dcc.Graph(id="graph", style={"height": f"{DASHBOARD_HEIGHT_PX}px"}, config=GRAPH_CONFIG)],
                                            ),
                                            html.Div(
                                                id="wrap-grid",
                                                style={"display": "none"},
                                                children=[
                                                    dag.AgGrid(
                                                        id="grid",
                                                        className=AG_THEME,
                                                        rowData=[],
                                                        columnDefs=COLUMN_DEFS,
                                                        defaultColDef={"resizable": True, "sortable": True, "filter": True, "floatingFilter": True},
                                                        dashGridOptions={
                                                            "rowSelection": "multiple",
                                                            "animateRows": True,
                                                            "undoRedoCellEditing": True,
                                                            "undoRedoCellEditingLimit": 50,
                                                                                                                    },
                                                        style={"height": f"{DASHBOARD_HEIGHT_PX}px", "width": "100%"},
                                                    )
                                                ],
                                            ),
                                            html.Div(
                                                id="wrap-metrics",
                                                style={"display": "none"},
                                                children=[html.Div(id="metrics-panel")],
                                            ),
                                            html.Div(
                                                id="wrap-petroqc",
                                                style={"display": "none"},
                                                children=[
                                                    dag.AgGrid(
                                                        id="grid-petroqc",
                                                        className=AG_THEME,
                                                        columnDefs=PETROQC_COLUMN_DEFS,
                                                        rowData=[],
                                                        defaultColDef={
                                                            "resizable": True,
                                                            "sortable": True,
                                                            "filter": True,
                                                            "floatingFilter": True,
                                                        },
                                                        dashGridOptions={
                                                            "rowSelection": "single",
                                                            "animateRows": True,
                                                            # Shade discarded / excluded samples in yellow
                                                            "getRowStyle": {
                                                                "styleConditions": [
                                                                    {
                                                                        # rowData uses 'exclude_from_shm' (snake_case). Using
                                                                        # a mismatched key prevents the highlight from triggering.
                                                                        "condition": "params.data && params.data.exclude_from_shm",
                                                                        "style": {"backgroundColor": "#fff3cd"},
                                                                    }
                                                                ]
                                                            },
                                                        },
                                                        style={"height": f"{DASHBOARD_HEIGHT_PX}px", "width": "100%"},
                                                    )
                                                ],
                                            ),
                                        ],
                                        style={"minHeight": f"{DASHBOARD_HEIGHT_PX + 70}px"},
                                    ),
                                ],
                            ),

                            dbc.Card(
                                body=True,
                                children=[
                                    html.Div("Alerts and comments", className="section-title"),
                                    html.Pre(id="status-text", style={"whiteSpace": "pre-wrap", "margin": 0}),
                                ],
                            ),
                        ],
                    ),

                    # RIGHT: Control Panel
                    dbc.Col(
                        width=4,
                        className="control-sidebar",
                        children=[
                            dbc.Card(
                                className="mb-2",
                                body=True,
                                children=[
                                    html.Div("Pre‑Process", className="section-title"),
                                    dbc.Row(
                                        className="g-2",
                                        children=[
                                            dbc.Col(
                                                dcc.Upload(
                                                    id="upload-data",
                                                    children=_btn("Import Data", "btn-import-proxy"),
                                                    multiple=True,
                                                    accept=".xls,.xlsx,.csv,.txt",
                                                    style={"width": "100%"},
                                                ),
                                                width=6,
                                            ),
                                            dbc.Col(_btn("Generate report", "btn-report", outline=False), width=6),
                                            dbc.Col(_btn("Run QA/QC", "btn-qaqc", outline=True), width=6),
                                            dbc.Col(_btn("Run PetroQC", "btn-petroqc", outline=True), width=6),
                                            dbc.Col(_btn("Scrub Data", "btn-scrub", outline=True), width=6),
                                            dbc.Col(_btn("Export CSV", "btn-export-csv", outline=True), width=6),
                                            dbc.Col(_btn("Closure Corr", "btn-conf", outline=True), width=6),
                                            dbc.Col(_btn("Undo Closure", "btn-conf-undo", outline=True), width=6),
                                            dbc.Col(_btn("Add row", "btn-add-row", outline=True), width=6),
                                            dbc.Col(_btn("Export PetroQC", "btn-export-petroqc", outline=True), width=6),
                                        ],
                                    ),
                                    dcc.Download(id="download-report"),
                                    dcc.Download(id="download-csv"),
                                    dcc.Download(id="download-petroqc"),

                                    html.Hr(),
                                    html.Div("Well / Session", className="section-title"),
                                    dbc.InputGroup(
                                        [
                                            dbc.InputGroupText("Well"),
                                            dbc.Input(id="inp-well", placeholder="(auto)", type="text"),
                                        ],
                                        size="sm",
                                    ),
                                    html.Div(style={"height": "8px"}),
                                    html.Div(id="loaded-files", className="small-muted"),
                                    dbc.Button("Remove sample(s)…", id="btn-remove-sample", color="danger", outline=True, size="sm", className="mt-2"),
                                    html.Div("(Removes from the current project/session only; it does NOT delete any files on disk)", className="text-muted small"),
                                ],
                            ),


                            dbc.Card(
                                className="mb-2",
                                body=True,
                                children=[
                                    html.Div("Calculations", className="section-title"),
                                    dbc.Row(
                                        className="g-2",
                                        children=[
                                            dbc.Col(_btn("Permeability (Swanson/Winland)", "btn-perm", outline=False), width=6),
                                            dbc.Col(_btn("Pc → Throat (Radius)", "btn-radius", outline=True), width=6),
                                            dbc.Col(_btn("Pc → Throat (Diameter)", "btn-diam", outline=True), width=6),
                                            dbc.Col(_btn("IFT/Angle Params", "btn-params", outline=True), width=6),
                                        ],
                                    ),
                                ],
                            ),

                            dbc.Card(
                                className="mb-2",
                                body=True,
                                children=[
                                    html.Div("Modeling & Rock Typing", className="section-title"),
                                    dbc.Row(
                                        className="g-2",
                                        children=[
                                            dbc.Col(_btn("Fit Thomeer", "btn-thomeer", outline=False), width=6),
                                            dbc.Col(_btn("J‑Function", "btn-jfunc", outline=True), width=6),
                                            dbc.Col(_btn("Cluster Samples", "btn-cluster", outline=True), width=6),
                                            dbc.Col(_btn("Export SHF", "btn-export-shf", outline=True), width=6),
                                            dbc.Col(_btn("Export Thomeer Params", "btn-export-thomeer", outline=True), width=6),
                                            dbc.Col(_btn("Fit Thomeer (Bimodal)", "btn-thomeer-bimodal", outline=True), width=6),
                                            dbc.Col(_btn("Run PNM (Fast)", "btn-pnm-fast", outline=True), width=6),
                                        ],
                                    ),
                                    dcc.Download(id="download-shf"),
                                    dcc.Download(id="download-thomeer"),

                                    dbc.Alert(
                                        id="alert-bimodal-hint",
                                        is_open=False,
                                        color="warning",
                                        className="mt-2",
                                        style={"fontSize": "12px", "padding": "8px 10px"},
                                        children=[
                                            html.Div(id="alert-bimodal-hint-body"),
                                            dbc.Button(
                                                "Run Bimodal Fit",
                                                id="btn-bimodal-hint-run",
                                                n_clicks=0,
                                                color="primary",
                                                size="sm",
                                                className="mt-2",
                                                style={"display": "none", "color": "#FFFFFF"},
                                            ),
                                        ],
                                    ),

                                    html.Hr(className="my-2"),
                                    html.Div("Thomeer controls (manual)", className="section-title", style={"fontSize": "13px"}),
                html.Div("Closure correction: drag the Closure line in the Pc plot.", className="small-muted"),
                html.Div("Bimodal Thomeer: Pd1/Pd2 are fitted from the corrected curve (not tied to closure).", className="small-muted"),
                                    html.Div(style={"height": "6px"}),
                                    html.Div("G (pore geometrical factor)", className="small-muted"),
                                    dcc.Slider(id="slider-thomeer-g", min=0.01, max=3.0, step=0.01, value=0.10, tooltip={"placement": "bottom", "always_visible": False}),
                                    html.Div(id="thomeer-g-readout", className="thomeer-readout"),
                                    html.Div(style={"height": "10px"}),
                                    html.Div("Bv∞ (% bulk)", className="small-muted"),
                                    dcc.Slider(id="slider-thomeer-bvinf", min=0.0, max=20.0, step=0.1, value=5.0, tooltip={"placement": "bottom", "always_visible": False}),
                                    html.Div(id="thomeer-bvinf-readout", className="thomeer-readout"),
                                    html.Div(id="thomeer-pd-readout", className="thomeer-readout"),

                                ],
                            ),

                            dbc.Card(
                                body=True,
                                children=[
                                    html.Div("Visualization", className="section-title"),
                                    dbc.Row(
                                        className="g-2 align-items-center mb-2",
                                        children=[
                                            dbc.Col(html.Div("Plot BG", className="text-muted", style={"fontSize": "0.90rem"}), width=4),
                                            dbc.Col(
                                                dcc.Dropdown(
                                                    id="dd-plot-theme",
                                                    options=[
                                                        {"label": "Dark", "value": "dark"},
                                                        {"label": "Light", "value": "light"},
                                                        {"label": "White", "value": "white"},
                                                    ],
                                                    value="dark",
                                                    clearable=False,
                                                ),
                                                width=8,
                                            ),
                                        ],
                                    ),
                                    dbc.Row(
                                        className="g-2 align-items-center mb-2",
                                        children=[
                                            dbc.Col(html.Div("SHF Axis", className="text-muted", style={"fontSize": "0.90rem"}), width=4),
                                            dbc.Col(
                                                dcc.Dropdown(
                                                    id="dd-shf-axis",
                                                    options=[
                                                        {"label": "Height (m)", "value": "height_m"},
                                                        {"label": "Height (ft)", "value": "height_ft"},
                                                        {"label": "Depth (m)", "value": "depth_m"},
                                                        {"label": "Depth (ft)", "value": "depth_ft"},
                                                    ],
                                                    value="height_m",
                                                    clearable=False,
                                                ),
                                                width=8,
                                            ),
                                        ],
                                    ),
                                    dbc.Row(
                                        className="g-2 align-items-center mb-2",
                                        children=[
                                            dbc.Col(html.Div("k Profile", className="text-muted", style={"fontSize": "0.90rem"}), width=4),
                                            dbc.Col(
                                                dcc.RadioItems(
                                                    id="dd-kprof-mode",
                                                    options=[
                                                    {"label": "Log(k) interpolation", "value": "logk"},
                                                    {"label": "Raw points", "value": "raw"},
                                                ],
                                                value="logk",
                                                    inline=True,
                                                    className="radio-kprof",
                                                    labelStyle={"color": "#ffffff", "marginRight": "14px"},
                                                    inputStyle={"marginRight": "6px"},
                                                    style={"color": "#ffffff"},
                                                ),
                                                width=8,
                                            ),
                                        ],
                                    ),
                                    dbc.Row(
                                        className="g-2 align-items-center mb-2",
                                        children=[
                                            dbc.Col(
                                                html.Div("Fill", className="text-muted", style={"fontSize": "0.90rem"}),
                                                width=4,
                                            ),
                                            dbc.Col(
                                                dcc.Checklist(
                                                    id="chk-kprof-fill",
                                                    options=[{"label": "Enable shaded fill", "value": "fill"}],
                                                    # Use init_ui here (make_layout scope). "ui" is only available inside callbacks.
                                                    value=["fill"] if init_ui.get("kprof_fill") else [],
                                                    labelStyle={"display": "inline-block", "marginRight": "14px", "color": "#ffffff"},
                                                    inputStyle={"marginRight": "6px"},
                                                    style={"color": "#ffffff"},
                                                ),
                                                width=8,
                                            ),
                                        ],
                                    ),
                                    dbc.Row(
                                        className="g-2",
                                        children=[
                                            dbc.Col(_btn("Log/Lin X", "btn-toggle-xlog", outline=True), width=6),
                                            dbc.Col(_btn("Overlay Inc/Cum", "btn-toggle-overlay", outline=True), width=6),
                                            dbc.Col(_btn("Intrusion", "btn-viz-intrusion", outline=False), width=6),
                                            dbc.Col(_btn("Pc vs Sw", "btn-viz-pcsw", outline=False), width=6),
                                            dbc.Col(_btn("PSD", "btn-viz-psd", outline=False), width=6),
                                            dbc.Col(_btn("Thomeer", "btn-viz-thomeer", outline=False), width=6),
                                            dbc.Col(_btn("SHF", "btn-viz-shf", outline=False), width=6),
                                            dbc.Col(_btn("Winland", "btn-viz-winland", outline=False), width=6),
                                            dbc.Col(_btn("PNM 3D", "btn-viz-pnm3d", outline=False), width=6),
                                        ],
                                    ),
                                ],
                            ),

                                                        dbc.Card(
                                className="mb-2",
                                body=True,
                                children=[
                                    html.Div("Multi-Sample Plots", className="section-title"),
                                    dbc.Row(
                                        className="g-2",
                                        children=[
                                            dbc.Col(_btn("Pc Overlay", "btn-ms-pc-overlay", outline=True), width=6),
                                            dbc.Col(_btn("PSD Compare", "btn-ms-psd-compare", outline=True), width=6),
                                            dbc.Col(_btn("Cum Intrusion", "btn-ms-cum-intrusion", outline=True), width=6),
                                            dbc.Col(_btn("Φ vs k", "btn-ms-phik", outline=True), width=6),
                                            dbc.Col(_btn("J-Function", "btn-ms-jfunc", outline=True), width=6),
                                            dbc.Col(_btn("G vs Pd", "btn-ms-gpd", outline=True), width=6),
                                            dbc.Col(_btn("Petro Logs", "btn-ms-petrolog", outline=True), width=6),
                                            dbc.Col(_btn("k Profile", "btn-ms-kprof", outline=True), width=6),
                                            dbc.Col(_btn("HFU Log", "btn-ms-hfu", outline=True), width=6),
                                            dbc.Col(_btn("SHM Curves", "btn-ms-shm", outline=True), width=6),
                                            dbc.Col(_btn("k Core vs k PNM", "btn-ms-kpnm", outline=True), width=6),
                                            dbc.Col(_btn("CI Log (PNM/Core)", "btn-ms-ci-log", outline=True), width=6),
                                            dbc.Col(_btn("Reset (Single)", "btn-ms-reset", outline=True), width=6),
                                        ],
                                    ),
                                ],
                            ),


                    dbc.Card(
                        className="mb-2",
                        body=True,
                        children=[
                            html.Div("External Core Logs", className="section-title"),
                            html.Div(
                                "Import LogN4A/LogN4_Vert logs (PorAmb, PorOB, PermAmb, PermOB, etc.) when available.",
                                className="text-muted small",
                                style={"marginBottom": "6px"},
                            ),
                            dbc.Row(
                                className="g-2",
                                children=[
                                    dbc.Col(
                                        dcc.Upload(
                                            id="upload-logn4",
                                            children=_btn("Import N4 Logs", "btn-import-logn4-proxy", outline=False),
                                            multiple=True,
                                            accept=".xlsx,.xls",
                                            style={"width": "100%"},
                                        ),
                                        width=6,
                                    ),
                                    dbc.Col(_btn("Clear", "btn-logn4-clear", outline=True), width=3),
                                    dbc.Col(_btn("Export CSV", "btn-logn4-export", outline=True), width=3),
                                ],
                            ),
                            html.Div(id="logn4-status", className="text-muted small", style={"marginTop": "6px"}),
                            dbc.Row(
                                className="g-2",
                                style={"marginTop": "6px"},
                                children=[
                                    dbc.Col(
                                        dcc.Dropdown(
                                            id="dd-logn4-sheet",
                                            options=[],
                                            value=None,
                                            clearable=False,
                                            placeholder="Select sheet...",
                                        ),
                                        width=4,
                                    ),
                                    dbc.Col(
                                        dcc.Dropdown(
                                            id="dd-logn4-logs",
                                            options=[],
                                            value=[],
                                            multi=True,
                                            placeholder="Select logs...",
                                        ),
                                        width=8,
                                    ),
                                ],
                            ),
                            dcc.Graph(
                                id="graph-logn4",
                                figure=go.Figure(),
                                config=PLOTLY_CONFIG,
                                style={"height": "320px", "marginTop": "6px"},
                            ),
                            dcc.Download(id="download-logn4-csv"),
                        ],
                    ),

dbc.Card(
    className="mb-2",
    body=True,
    children=[
        dbc.CardHeader(
            [
                html.Div("Core Validation", className="section-title"),
                html.Div(
                    "Fast QC: stress sensitivity + model mismatch + bimodal evidence",
                    className="section-subtitle",
                ),
            ]
        ),
        dbc.CardBody(
            [
                html.Div(
                    "Uses PermAmb/PermOB from the imported N4 logs (when present) and compares with Winland/PNM predictions.",
                    className="text-muted",
                    style={"fontSize": "12px", "marginBottom": "6px"},
                ),
                html.Div(id="core-validation-table"),
            ]
        ),
    ],
),

dbc.Card(
                                className="mb-2",
                                body=True,
                                children=[
                                    html.Div("Workflow Status Tracker", className="section-title"),
                                    html.Div("Step Progress Bar", className="small-muted"),
                                    dbc.Progress(id="wf-progress", value=0, label="0/5", className="mb-2"),
                                    html.Div(id="wf-steps", className="wf-steps"),

                                    html.Hr(className="my-2"),
                                    
                                    html.Div("KPIs", className="section-title"),
                                    dbc.Row(
                                        className="g-2",
                                        children=[
                                            dbc.Col(
                                                html.Div(
                                                    [
                                                        html.Div("Porosity (%)", className="wf-kpi-label"),
                                                        dbc.Input(id="kpi-phi", value="", disabled=True, className="kpi-input", size="sm"),
                                                    ]
                                                ),
                                                width=4,
                                            ),
                                            dbc.Col(
                                                html.Div(
                                                    [
                                                        html.Div("Permeability (mD)", className="wf-kpi-label"),
                                                        dbc.Input(id="kpi-k", value="", disabled=True, className="kpi-input", size="sm"),
                                                    ]
                                                ),
                                                width=4,
                                            ),
                                            dbc.Col(
                                                html.Div(
                                                    [
                                                        html.Div("r35 (µm)", className="wf-kpi-label"),
                                                        dbc.Input(id="kpi-r35", value="", disabled=True, className="kpi-input", size="sm"),
                                                    ]
                                                ),
                                                width=4,
                                            ),
                                        ],
                                    ),
                                    dbc.Row(
                                        className="g-2 mt-1",
                                        children=[
                                            dbc.Col(
                                                html.Div(
                                                    [
                                                        html.Div("k method", className="wf-kpi-label"),
                                                        dbc.Input(id="kpi-kmethod", value="", disabled=True, className="kpi-input", size="sm"),
                                                    ]
                                                ),
                                                width=6,
                                            ),
                                            dbc.Col(
                                                html.Div(
                                                    [
                                                        html.Div("RRT", className="wf-kpi-label"),
                                                        dbc.Input(id="kpi-rrt", value="", disabled=True, className="kpi-input", size="sm"),
                                                    ]
                                                ),
                                                width=6,
                                            ),
                                        ],
                                    ),

html.Hr(className="my-2"),

                                    html.Div("Petrophysical QAQC", className="section-title"),
                                    html.Div(id="petro-qc-summary", className="petro-qc-box"),

                                    html.Hr(className="my-2"),
                                    html.Div("Sample Decision", className="section-title"),
                                    dbc.Row(
                                        className="g-2 align-items-center",
                                        children=[
                                            dbc.Col(dbc.Badge("PENDING", id="badge-decision", color="secondary"), width="auto"),
                                            dbc.Col(dbc.Badge("Petro QC: —", id="badge-petro", color="dark"), width="auto"),
                                            dbc.Col(dbc.Badge("Recommend: —", id="badge-recommend", color="dark"), width="auto"),
                                        ],
                                    ),
                                    dbc.Row(
                                        className="g-2 mt-2",
                                        children=[
                                            dbc.Col(_btn("Accept Sample", "btn-accept-sample", outline=False), width=6),
                                            dbc.Col(_btn("Discard Sample", "btn-discard-sample", outline=True), width=6),
                                        ],
                                    ),

                                    html.Hr(className="my-2"),
                                    html.Div("Log", className="section-title"),
                                    dcc.Textarea(
                                        id="wf-log",
                                        value="",
                                        readOnly=True,
                                        className="log-area",
                                        style={"height": "180px"},
                                    ),
                                ],
                            ),

                        ],
                    ),
                ]
            ),

            # Modal: Parameters
            dbc.Modal(
                id="modal-params",
                is_open=False,
                size="lg",
                children=[
                    dbc.ModalHeader(dbc.ModalTitle("IFT / Angle / Overrides")),
                    dbc.ModalBody(
                        dbc.Row(
                            className="g-3",
                            children=[
                                dbc.Col(dbc.InputGroup([dbc.InputGroupText("σ Hg-air (N/m)"), dbc.Input(id="p-sigma-hg", type="number")]), width=6),
                                dbc.Col(dbc.InputGroup([dbc.InputGroupText("θ Hg-air (deg)"), dbc.Input(id="p-theta-hg", type="number")]), width=6),
                                dbc.Col(dbc.InputGroup([dbc.InputGroupText("σ res (N/m)"), dbc.Input(id="p-sigma-res", type="number")]), width=6),
                                dbc.Col(dbc.InputGroup([dbc.InputGroupText("θ res (deg)"), dbc.Input(id="p-theta-res", type="number")]), width=6),
                                dbc.Col(dbc.InputGroup([dbc.InputGroupText("ρw (kg/m³)"), dbc.Input(id="p-rho-w", type="number")]), width=6),
                                dbc.Col(dbc.InputGroup([dbc.InputGroupText("ρhc (kg/m³)"), dbc.Input(id="p-rho-hc", type="number")]), width=6),
                                dbc.Col(dbc.InputGroup([dbc.InputGroupText("FWL depth (m)"), dbc.Input(id="p-fwl-depth", type="number", placeholder="optional")]), width=6),
                                dbc.Col(dbc.InputGroup([dbc.InputGroupText("φ override (%)"), dbc.Input(id="p-phi-ovr", type="number", placeholder="optional")]), width=6),
                                dbc.Col(dbc.InputGroup([dbc.InputGroupText("k override (mD)"), dbc.Input(id="p-k-ovr", type="number", placeholder="optional")]), width=6),
                                dbc.Col(dbc.InputGroup([dbc.InputGroupText("Swanson a"), dbc.Input(id="p-swa", type="number")]), width=6),
                                dbc.Col(dbc.InputGroup([dbc.InputGroupText("Swanson b"), dbc.Input(id="p-swb", type="number")]), width=6),
                                dbc.Col(html.Div("Rock Type bins/labels are configurable in code (DEFAULT_PARAMS).", className="small-muted"), width=12),
                            ],
                        )
                    ),
                    dbc.ModalFooter(
                        dbc.ButtonGroup(
                            [
                                dbc.Button("Close", id="btn-params-close", color="secondary"),
                                dbc.Button("Apply", id="btn-params-apply", color="primary"),
                            ]
                        )
                    ),
                ],
            ),


            # Modal: Discard reason (required for DISCARDED decision)
            dbc.Modal(
                id="modal-discard",
                is_open=False,
                size="lg",
                children=[
                    dbc.ModalHeader(dbc.ModalTitle("Discard Sample — Provide Reason")),
                    dbc.ModalBody(
                        [
                            html.Div(
                                "Please write the reasons to discard this sample. This will be saved in the session and included in the PDF report.",
                                className="small-muted",
                                style={"marginBottom": "0.5rem"},
                            ),
                            dbc.Textarea(
                                id="txt-discard-reason",
                                value="",
                                placeholder="E.g., apparent/skeletal density out of range; stem volume used outside [0,100]%; porosity > 100%; non-physical tortuosity; etc.",
                                style={"minHeight": "140px"},
                            ),
                            html.Div(id="discard-error", className="text-danger small", style={"marginTop": "0.5rem"}),
                        ]
                    ),
                    dbc.ModalFooter(
                        dbc.ButtonGroup(
                            [
                                dbc.Button("Cancel", id="btn-discard-cancel", color="secondary"),
                                dbc.Button("Confirm discard", id="btn-discard-confirm", color="danger"),
                            ]
                        )
                    ),
                ],
            ),

            # Modal: Remove samples (multi-select)
            dbc.Modal(
                id="modal-remove-samples",
                is_open=False,
                size="lg",
                centered=True,
                children=[
                    dbc.ModalHeader(dbc.ModalTitle("Remove samples from project")),
                    dbc.ModalBody(
                        [
                            html.Div(
                                "Select one or more samples to remove from the current project list. "
                                "This will NOT delete the original Excel files from disk.",
                                style={"marginBottom": "10px"},
                            ),
                            dcc.Dropdown(
                                id="dd-remove-samples",
                                options=[],
                                value=[],
                                multi=True,
                                placeholder="Select samples...",
                                style={"color": "#000"},
                            ),
                            html.Div(
                                "Tip: You can select multiple entries and remove them in one action.",
                                style={"fontSize": "12px", "opacity": 0.85, "marginTop": "8px"},
                            ),
                        ]
                    ),
                    dbc.ModalFooter(
                        [
                            dbc.Button("Cancel", id="btn-remove-cancel", color="secondary", n_clicks=0, className="me-2"),
                            dbc.Button("Remove selected", id="btn-remove-confirm", color="danger", n_clicks=0),
                        ]
                    ),
                ],
            ),


        ],
    )

app.layout = make_layout()

# ---------------------------
# Tabs visibility (Plot / Grid / Metrics)
# ---------------------------
@app.callback(
    Output("wrap-plot", "style"),
    Output("wrap-grid", "style"),
    Output("wrap-metrics", "style"),
    Output("wrap-petroqc", "style"),
    Input("tabs-left", "active_tab"),
)
def toggle_tabs(active_tab):
    show = {"display": "block"}
    hide = {"display": "none"}

    if active_tab == "tab-grid":
        return hide, show, hide, hide
    if active_tab == "tab-metrics":
        return hide, hide, show, hide
    if active_tab == "tab-petroqc":
        return hide, hide, hide, show

    # Default
    return show, hide, hide, hide


# ---------------------------
# Helper: library operations
# ---------------------------
def _now_iso() -> str:
    return _dt.datetime.now().isoformat(timespec="seconds")

def _make_unique_sample_id(existing_ids: set, base: str) -> str:
    base = base.strip() or "sample"
    if base not in existing_ids:
        return base
    k = 2
    while f"{base} ({k})" in existing_ids:
        k += 1
    return f"{base} ({k})"

def _lib_get(library: List[Dict[str, Any]], sample_id: Optional[str]) -> Optional[Dict[str, Any]]:
    if not sample_id:
        return None
    for s in library:
        if s.get("sample_id") == sample_id:
            return s
    return None

def _lib_set(library: List[Dict[str, Any]], sample: Dict[str, Any]) -> List[Dict[str, Any]]:
    sid = sample.get("sample_id")
    out = []
    replaced = False
    for s in library:
        if s.get("sample_id") == sid:
            out.append(sample)
            replaced = True
        else:
            out.append(s)
    if not replaced:
        out.append(sample)
    return out

# ---------------------------
# Top bar: New / Save / Open
# ---------------------------
@app.callback(
    Output("store-library", "data", allow_duplicate=True),
    Output("store-current-id", "data", allow_duplicate=True),
    Output("store-params", "data", allow_duplicate=True),
    Output("store-ui", "data", allow_duplicate=True),
    Output("store-log", "data", allow_duplicate=True),
    Output("store-petroqc", "data", allow_duplicate=True),
    Output("store-status", "data", allow_duplicate=True),
    Input("btn-new", "n_clicks"),
    State("store-workspace", "data"),
    prevent_initial_call=True,
)
def on_new(_, ws_data):
    # Clear disk autosave (best-effort)
    try:
        ws_dir = (ws_data or {}).get("dir") or WORKSPACE_DIR
        ws_paths = ensure_workspace(ws_dir)
        auto_path = os.path.join(ws_paths["projects"], "autosave_project.json")
        if os.path.exists(auto_path):
            os.remove(auto_path)
    except Exception:
        pass

    return (
        [],
        None,
        DEFAULT_PARAMS,
        {"plot_mode": "intrusion", "xlog": True, "overlay_inc": False, "plot_theme": "dark"},
        [],
        [],
        "New session started.",
    )

@app.callback(
    Output("download-project", "data"),
    Output("store-status", "data", allow_duplicate=True),
    Output("store-project-refresh", "data", allow_duplicate=True),
    Input("btn-save", "n_clicks"),
    State("store-library", "data"),
    State("store-current-id", "data"),
    State("store-params", "data"),
    State("store-ui", "data"),
    State("store-log", "data"),
    State("store-petroqc", "data"),
    State("store-workspace", "data"),
    prevent_initial_call=True,
)
def on_save_project(_, library, current_id, params, ui, log, petroqc, ws_data):
    ui = ui or {"plot_mode": "intrusion", "xlog": True, "overlay_inc": False}
    ui.setdefault("plot_theme", "dark")

    payload = {
        "version": VERSION,
        "saved_at": _now_iso(),
        "library": library or [],
        "current_id": current_id,
        "params": params or DEFAULT_PARAMS,
        "ui": ui,
        "log": log or [],
        "petroqc": petroqc or [],
    }

    ws_dir = (ws_data or {}).get("dir") or WORKSPACE_DIR
    ws_paths = ensure_workspace(ws_dir)
    projects_dir = ws_paths["projects"]

    # Human-friendly filename (uses current sample if available)
    base = "geopore_project"
    try:
        if current_id:
            s = next((x for x in (library or []) if x.get("sample_id") == current_id), None)
            if s:
                base = s.get("sample_name") or s.get("sample_id") or base
    except Exception:
        pass
    base = re.sub(r"[^A-Za-z0-9_.-]+", "_", str(base)).strip("_") or "geopore_project"
    ts = _dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    disk_name = f"{base}_{ts}.json"
    disk_path = os.path.join(projects_dir, disk_name)

    # Save to disk (best-effort)
    _safe_write_json(disk_path, payload)

    # Update autosave too
    auto_path = os.path.join(projects_dir, "autosave_project.json")
    _safe_write_json(auto_path, payload)

    status = f"Saved project to disk: {disk_path}"
    refresh_token = _dt.datetime.now().timestamp()

    # Keep browser-download export too (optional but handy)
    download = dcc.send_bytes(
        lambda b: b.write(json.dumps(payload, indent=2).encode("utf-8")),
        disk_name,
    )
    return download, status, refresh_token

@app.callback(
    Output("store-library", "data", allow_duplicate=True),
    Output("store-current-id", "data", allow_duplicate=True),
    Output("store-params", "data", allow_duplicate=True),
    Output("store-ui", "data", allow_duplicate=True),
    Output("store-log", "data", allow_duplicate=True),
    Output("store-petroqc", "data", allow_duplicate=True),
    Output("store-status", "data", allow_duplicate=True),
    Input("upload-project", "contents"),
    State("upload-project", "filename"),
    prevent_initial_call=True,
)
def on_open_project(contents, filename):
    if not contents:
        return no_update, no_update, no_update, no_update, no_update, no_update, no_update
    try:
        content_type, content_string = contents.split(",", 1)
        decoded = base64.b64decode(content_string)
        payload = json.loads(decoded.decode("utf-8"))

        library = payload.get("library") or []
        current_id = payload.get("current_id")
        if not current_id and library:
            current_id = library[0].get("sample_id")

        params = payload.get("params") or DEFAULT_PARAMS
        ui = payload.get("ui") or {"plot_mode": "intrusion", "xlog": True, "overlay_inc": False}
        ui.setdefault("plot_theme", "dark")

        log = payload.get("log") or []
        petroqc = payload.get("petroqc") or []

        return library, current_id, params, ui, log, petroqc, f"Loaded project: {filename}"
    except Exception as e:
        return no_update, no_update, no_update, no_update, no_update, no_update, f"Failed to open project: {e}"


# ---------------------------
# Workspace / Project Storage (disk persistence)
# ---------------------------

@app.callback(
    # Keep the input in sync with the current workspace on page load.
    # NOTE: do NOT use allow_duplicate here; the "Browse" callback below uses
    # allow_duplicate with prevent_initial_call=True to safely share this output.
    Output("inp-workspace-dir", "value", allow_duplicate=True),
    Input("store-workspace", "data"),
    prevent_initial_call='initial_duplicate',
)
def _sync_workspace_input(ws_data):
    return (ws_data or {}).get("dir") or WORKSPACE_DIR


@app.callback(
    Output("store-workspace", "data", allow_duplicate=True),
    Output("store-status", "data", allow_duplicate=True),
    Input("inp-workspace-dir", "value"),
    State("store-workspace", "data"),
    prevent_initial_call=True,
)
def _set_workspace_dir(inp_value, ws_data):
    new_dir = (inp_value or "").strip()
    if not new_dir:
        return no_update, no_update

    current_dir = (ws_data or {}).get("dir")
    try:
        if current_dir and os.path.abspath(current_dir) == os.path.abspath(new_dir):
            return no_update, no_update
    except Exception:
        pass
    try:
        set_workspace_dir(new_dir)
        ws_paths = ensure_workspace(new_dir)
        return {"dir": ws_paths["workspace"]}, f"Workspace set to: {ws_paths['workspace']}"
    except Exception as e:
        return no_update, f"Workspace error: {e}"


@app.callback(
    Output("dd-project-files", "options"),
    Output("dd-project-files", "value"),
    # The layout provides `lbl-workspace-path` to display the active workspace/projects paths.
    # (Earlier versions referenced `workspace-hint`, which does not exist in the layout and
    # triggers a front-end "ID not found in layout" error.)
    Output("lbl-workspace-path", "children"),
    Input("store-workspace", "data"),
    Input("btn-refresh-projects", "n_clicks"),
    Input("store-project-refresh", "data"),
    State("dd-project-files", "value"),
)
def _refresh_project_dropdown(ws_data, _n_refresh, _token, current_value):
    ws_dir = (ws_data or {}).get("dir") or WORKSPACE_DIR
    ws_paths = ensure_workspace(ws_dir)
    projects_dir = ws_paths["projects"]
    auto_path = os.path.join(projects_dir, "autosave_project.json")

    options = list_project_files(projects_dir)
    values = [o.get("value") for o in options if isinstance(o, dict)]

    if current_value not in values:
        if auto_path in values:
            current_value = auto_path
        else:
            current_value = values[0] if values else None

    hint = f"Workspace: {ws_dir} | Projects: {projects_dir}"
    return options, current_value, hint


@app.callback(
    Output("store-library", "data", allow_duplicate=True),
    Output("store-current-id", "data", allow_duplicate=True),
    Output("store-params", "data", allow_duplicate=True),
    Output("store-ui", "data", allow_duplicate=True),
    Output("store-log", "data", allow_duplicate=True),
    Output("store-petroqc", "data", allow_duplicate=True),
    Output("store-status", "data", allow_duplicate=True),
    Input("btn-load-project", "n_clicks"),
    State("dd-project-files", "value"),
    prevent_initial_call=True,
)
def _load_project_from_disk(_n, filepath):
    if not filepath:
        return (
            no_update,
            no_update,
            no_update,
            no_update,
            no_update,
            no_update,
            "Select a project file first.",
        )
    if not os.path.exists(filepath):
        return (
            no_update,
            no_update,
            no_update,
            no_update,
            no_update,
            no_update,
            f"File not found: {filepath}",
        )

    payload = _safe_read_json(filepath, {})
    if not isinstance(payload, dict):
        return (
            no_update,
            no_update,
            no_update,
            no_update,
            no_update,
            no_update,
            f"Invalid project file: {filepath}",
        )

    library = payload.get("library") or []
    current_id = payload.get("current_id")
    if not current_id and library:
        current_id = library[0].get("sample_id")

    params = payload.get("params") or DEFAULT_PARAMS
    ui = payload.get("ui") or {"plot_mode": "intrusion", "xlog": True, "overlay_inc": False}
    ui.setdefault("plot_theme", "dark")

    log = payload.get("log") or []
    petroqc = payload.get("petroqc") or []

    return (
        library,
        current_id,
        params,
        ui,
        log,
        petroqc,
        f"Loaded project from disk: {os.path.basename(filepath)}",
    )


@app.callback(
    Output("autosave-dummy", "children"),
    Input("store-library", "data"),
    Input("store-current-id", "data"),
    Input("store-params", "data"),
    Input("store-ui", "data"),
    Input("store-log", "data"),
    Input("store-petroqc", "data"),
    State("store-workspace", "data"),
    prevent_initial_call=True,
)
def _autosave_to_disk(library, current_id, params, ui, log, petroqc, ws_data):
    """Auto-save session to disk so a PC restart doesn't wipe progress."""
    ws_dir = (ws_data or {}).get("dir") or WORKSPACE_DIR
    ws_paths = ensure_workspace(ws_dir)
    auto_path = os.path.join(ws_paths["projects"], "autosave_project.json")

    ui = ui or {"plot_mode": "intrusion", "xlog": True, "overlay_inc": False}
    ui.setdefault("plot_theme", "dark")

    payload = {
        "version": VERSION,
        "saved_at": _now_iso(),
        "library": library or [],
        "current_id": current_id,
        "params": params or DEFAULT_PARAMS,
        "ui": ui,
        "log": log or [],
        "petroqc": petroqc or [],
        "autosave": True,
    }

    _safe_write_json(auto_path, payload)
    return _now_iso()


# ---------------------------
# Plot theme selector (dark / light / white)
# ---------------------------

@app.callback(
    Output("store-ui", "data", allow_duplicate=True),
    Input("dd-plot-theme", "value"),
    Input("dd-shf-axis", "value"),
    Input("dd-kprof-mode", "value"),
    State("store-ui", "data"),
    prevent_initial_call=True,
)
def _set_plot_theme(theme_value, shf_axis_value, kprof_mode_value, ui):
    ui = ui or {
        "plot_mode": "intrusion",
        "xlog": True,
        "overlay_inc": False,
        "plot_theme": "dark",
        "shf_axis": "height_m",
        "kprof_mode": "logk",
    }
    if theme_value is not None:
        ui["plot_theme"] = (theme_value or "dark").strip().lower()
    if shf_axis_value is not None:
        ui["shf_axis"] = (shf_axis_value or "height_m").strip().lower()
    if kprof_mode_value is not None:
        ui["kprof_mode"] = (kprof_mode_value or "logk").strip().lower()
    return ui


@app.callback(
    Output("dd-plot-theme", "value"),
    Input("store-ui", "data"),
)
def _sync_plot_theme_dropdown(ui):
    return (ui or {}).get("plot_theme", "dark")


@app.callback(Output("dd-shf-axis", "value"), Input("store-ui", "data"))
def _sync_shf_axis_dropdown(ui):
    ui = ui or {}
    return ui.get("shf_axis", "height_m")


@app.callback(Output("dd-kprof-mode", "value"), Input("store-ui", "data"))
def _sync_kprof_mode_dropdown(ui):
    ui = ui or {}
    return ui.get("kprof_mode", "logk")


@app.callback(
    Output("store-ui", "data", allow_duplicate=True),
    Input("chk-kprof-fill", "value"),
    State("store-ui", "data"),
    prevent_initial_call=True,
)
def _sync_kprof_fill(values, ui):
    ui = dict(ui or {})
    ui["kprof_fill"] = bool(values) and ("fill" in (values or []))
    return ui


@app.callback(
    Output("store-current-id", "data", allow_duplicate=True),
    Input("sel-sample", "value"),
    prevent_initial_call=True,
)
def on_select_sample(sample_id):
    return sample_id



# ---------------------------
# Remove selected sample(s) from the current project library (multi-select)
@app.callback(
    Output("modal-remove-samples", "is_open"),
    Output("dd-remove-samples", "options"),
    Output("dd-remove-samples", "value"),
    Output("store-library", "data", allow_duplicate=True),
    Output("store-current-id", "data", allow_duplicate=True),
    Output("store-log", "data", allow_duplicate=True),
    Output("store-status", "data", allow_duplicate=True),
    Input("btn-remove-sample", "n_clicks"),
    Input("btn-remove-confirm", "n_clicks"),
    Input("btn-remove-cancel", "n_clicks"),
    State("modal-remove-samples", "is_open"),
    State("dd-remove-samples", "value"),
    State("sel-sample", "value"),
    State("store-library", "data"),
    State("store-current-id", "data"),
    State("store-log", "data"),
    prevent_initial_call=True,
)
def on_remove_samples(n_open, n_confirm, n_cancel, is_open, remove_values, selected_id, library, current_id, logs):
    trig = callback_context.triggered_id

    library = _coerce_library_list(library)

    # Build dropdown options (robustly handle either "id" or "sample_id")
    def _sid(s):
        return (s or {}).get("id") or (s or {}).get("sample_id")

    dd_options = [{"label": _sample_label(s), "value": _sid(s)} for s in library if _sid(s)]

    # Open modal / refresh selection
    if trig in ("btn-remove-sample", "dd-remove-samples"):
        if remove_values is None:
            remove_values = [selected_id] if selected_id else []
        elif not isinstance(remove_values, (list, tuple)):
            remove_values = [remove_values]
        return True, dd_options, (remove_values or []), no_update, no_update, no_update, no_update

    # Cancel
    if trig == "btn-remove-cancel":
        return False, dd_options, (remove_values or []), no_update, no_update, no_update, no_update

    # Confirm remove
    if trig == "btn-remove-confirm":
        if not remove_values:
            return False, dd_options, [], no_update, no_update, no_update, "No samples selected."

        if not isinstance(remove_values, (list, tuple)):
            remove_values = [remove_values]

        to_remove = {str(x) for x in remove_values if x is not None and str(x) != ""}
        if not to_remove:
            return False, dd_options, [], no_update, no_update, no_update, "No samples selected."

        def sid_str(s):
            return str((s or {}).get("id") or (s or {}).get("sample_id") or "")

        library2 = [s for s in library if sid_str(s) not in to_remove]

        current_id2 = current_id
        if current_id and str(current_id) in to_remove:
            current_id2 = (_sid(library2[0]) if library2 else None)

        status = f"Removed {len(to_remove)} sample(s) from the session."

        # NOTE: We don't manually mutate store-log here. The existing
        # append_status_to_log() callback will stamp + append this status.
        return False, dd_options, [], library2, current_id2, no_update, status

    return is_open, dd_options, (remove_values or []), no_update, no_update, no_update, no_update

# ---------------------------
# Push current sample data into grid / metrics / graph
# ---------------------------
@app.callback(
    Output("grid", "rowData"),
    Output("metrics-panel", "children"),
    Output("graph", "figure"),
    Output("status-text", "children"),
    Input("store-library", "data"),
    Input("store-current-id", "data"),
    Input("store-ui", "data"),
    Input("store-status", "data"),
    Input("store-logn4", "data"),
    State("store-params", "data"),
)
def update_views(library, current_id, ui, status, logn4_store, params):
    library = library or []
    ui = ui or {"plot_mode": "intrusion", "xlog": True, "overlay_inc": False}
    params = params or DEFAULT_PARAMS
    mode = (ui or {}).get("plot_mode", "intrusion")

    sample = _lib_get(library, current_id) if current_id else None
    if not sample:
        empty_df = _ensure_schema(pd.DataFrame({"Pressure": [], "CumVol": [], "IncVol": []}))
        fig = go.Figure()
        fig = apply_plot_theme(fig, (ui or {}).get("plot_theme", "dark"))
        metrics = html.Div("No sample loaded.", className="small-muted")
        return empty_df.to_dict("records"), metrics, fig, status or ""

    df = pd.DataFrame(sample.get("data", []))
    df = _ensure_schema(df)
    df = recompute_derived(df, params)

    # Update back derived (so plots/grid consistent even after edits)
    sample = {**sample, "data": df.to_dict("records")}
    # We do NOT write back to store-library here to avoid loops; writebacks happen in action callbacks.

    res = sample.get("results", {}) or {}
    meta = sample.get("meta", {}) or {}

    # Metrics panel (quick view)
    # Backbone / fractal proxy pressure (computed if missing)
    bb_psia = res.get("backbone_pressure_psia")
    if bb_psia is None and df is not None:
        try:
            bb_psia, _bb_info = compute_backbone_pressure_psia(df, params)
        except Exception:
            bb_psia = None

    metrics_rows = [
        ("Sample", sample.get("sample_id")),
        ("File", sample.get("filename")),
        ("Well", sample.get("well")),
        ("Porosity (%)", meta.get("porosity_pct")),
        ("Permeability (mD)", meta.get("permeability_md")),
        ("Threshold Pressure (psia)", res.get("threshold_pressure_psia_used", meta.get("threshold_pressure_psia"))),
        ("Backbone / Fractal P (psia)", bb_psia),
        ("Winland mode", res.get("winland_mode")),
        ("r35 used (µm)", res.get("r35_um")),
        ("r35 total (µm)", res.get("r35_um_total")),
        ("r35 macro-norm (µm)", res.get("r35_um_macro")),
        ("Conformance Pc (psia)", res.get("conf_pknee_psia")),
        ("Conformance Vol (mL/g)", res.get("conf_vol_ml_g")),
        ("Rock Type", res.get("rock_type")),
        ("k_swanson (mD)", res.get("k_swanson_md")),
        ("k_winland used (mD)", res.get("k_winland_md")),
        ("k_winland total (mD)", res.get("k_winland_md_total")),
        ("k_winland macro (mD)", res.get("k_winland_md_macro")),
        ("k_pnm (mD)", res.get("k_pnm_md")),
        ("PNM CI (k_pnm/k_core)", res.get("pnm_ci")),
        ("Thomeer mode", res.get("thomeer_mode")),
        ("Thomeer Pd (psia)", res.get("thomeer_pd_psia")),
        ("Thomeer G", res.get("thomeer_G")),
        ("Thomeer Vb∞", res.get("thomeer_vb_inf")),
        ("Thomeer Pd1 (psia)", res.get("thomeer_pd1_psia") if res.get("thomeer_mode") == "bimodal" else None),
        ("Thomeer G1", res.get("thomeer_G1") if res.get("thomeer_mode") == "bimodal" else None),
        ("Thomeer Vb∞1", res.get("thomeer_vb_inf1") if res.get("thomeer_mode") == "bimodal" else None),
        ("Thomeer Pd2 (psia)", res.get("thomeer_pd2_psia") if res.get("thomeer_mode") == "bimodal" else None),
        ("Thomeer G2", res.get("thomeer_G2") if res.get("thomeer_mode") == "bimodal" else None),
        ("Thomeer Vb∞2", res.get("thomeer_vb_inf2") if res.get("thomeer_mode") == "bimodal" else None),
        ("Thomeer macro frac", res.get("thomeer_macro_frac") if res.get("thomeer_mode") == "bimodal" else None),
        ("Thomeer R²", res.get("thomeer_r2")),
        ("Cluster", res.get("cluster")),
    ]
    metrics_table = dbc.Table(
        [html.Thead(html.Tr([html.Th("Metric"), html.Th("Value")]))] +
        [html.Tbody([html.Tr([html.Td(k), html.Td("" if v is None else str(v))]) for k, v in metrics_rows])],
        bordered=True,
        hover=True,
        size="sm",
        style={"color": COLORS["text"]},
    )


    # Multi-sample default filter: hide discarded / hard-fail samples
    if mode == "winland" or str(mode).startswith("ms_"):
        try:
            library_ms = [
                s for s in _coerce_library_list(library)
                if not _is_excluded_sample_for_multisample(s, params)
            ]
        except Exception:
            library_ms = _coerce_library_list(library)
    else:
        library_ms = _coerce_library_list(library)

    # Figure selection
    if mode == "pcsw":
        fig = fig_pc_sw(df, ui)
    elif mode == "psd":
        fig = fig_psd(df, ui)
    elif mode == "thomeer":
        fig = fig_thomeer(df, res, params, meta, ui)
    elif mode == "shf":
        fig = fig_shf(df, params, ui)
    elif mode == "winland":
        fig = fig_winland_crossplot(library_ms)
    elif mode == "pnm3d":
        fig = fig_pnm3d_network(df, params, meta, res, ui, sample_id=sample.get("sample_id"))
    elif mode == "ms_pc_overlay":
        fig = fig_ms_pc_overlay(library_ms, ui, highlight_id=current_id)
    elif mode == "ms_psd_compare":
        fig = fig_ms_psd_compare(library_ms, ui, highlight_id=current_id)
    elif mode == "ms_cum_intrusion":
        fig = fig_ms_cum_intrusion(library_ms, ui, highlight_id=current_id)
    elif mode == "ms_phi_k":
        fig = fig_ms_phi_k_crossplot(library_ms, params, highlight_id=current_id)
    elif mode == "ms_jfunc":
        fig = fig_ms_j_function(library_ms, params, highlight_id=current_id)
    elif mode == "ms_g_pd":
        fig = fig_ms_g_vs_pd(library_ms, highlight_id=current_id)
    elif mode == "ms_petro_logs":
        fig = fig_ms_petro_logs(library_ms, params)
    elif mode == "ms_k_profile":
        # fig_ms_k_profile(project, mode=...) only needs the library/project dict.
        # Passing "params" here causes a server-side callback exception and makes the
        # "k Profile" button appear to do nothing.
        fig = fig_ms_k_profile(library_ms, mode=ui.get("kprof_mode", "logk"), fill_models=bool(ui.get("kprof_fill", False)), logn4_store=logn4_store)
    elif mode == "ms_hfu":
        fig = fig_ms_hfu_log(library_ms, params)
    elif mode == "ms_shm":
        fig = fig_ms_shm_curves(library_ms, params, ui)
    elif mode == "ms_k_pnm":
        fig = fig_ms_k_pnm_crossplot(library_ms, params, highlight_id=current_id)
    elif mode == "ms_ci_log":
        fig = fig_ms_ci_log(library_ms, params, highlight_id=current_id)
    else:
        fig = fig_intrusion(df, ui)

    # Closure/Conformance marker (draggable vertical line)
    # NOTE: Draggable behavior comes from dcc.Graph(config=GRAPH_CONFIG) and the shape below.
    conf_pk = res.get("conf_pknee_psia")
    conf_applied = bool(res.get("conf_applied"))
    
    # --- Reference vertical lines (Conformance, Threshold, Pd, Backbone) ---
    if mode in ("intrusion", "pcsw", "thomeer"):
        try:
            shapes = []
            annots = []

            # Conformance (data cleaning) line
            conf_line = conf_pk
            conf_editable = bool(res and res.get("conf_applied"))
            if (conf_line is None) and (df is not None):
                try:
                    conf_line, _vtmp, _m = detect_conformance_knee(df, params)
                except Exception:
                    conf_line = None
                conf_editable = False

            if conf_line is not None and np.isfinite(conf_line) and conf_line > 0:
                shapes.append(
                    dict(
                        type="line",
                        xref="x",
                        yref="paper",
                        x0=conf_line,
                        x1=conf_line,
                        y0=0,
                        y1=1,
                        line=dict(color="orange", width=2, dash="dash"),
                        editable=conf_editable,
                    )
                )
                annots.append(
                    dict(
                        x=conf_line,
                        y=1.0,
                        yref="paper",
                        text="Conformance",
                        showarrow=False,
                        font=dict(color="orange", size=12),
                        yshift=12,
                    )
                )

            # Threshold (Pth) line (entry into connected pore network)
            pth_line = None
            if res:
                pth_line = res.get("threshold_pressure_psia_used")
            if pth_line is None:
                pth_line = meta.get("threshold_pressure_psia") if meta else None
            if (pth_line is None) and (df is not None):
                pth_line, _pth_info = compute_threshold_pressure_psia(df, params)
            if pth_line is not None and np.isfinite(pth_line) and pth_line > 0:
                shapes.append(
                    dict(
                        type="line",
                        xref="x",
                        yref="paper",
                        x0=pth_line,
                        x1=pth_line,
                        y0=0,
                        y1=1,
                        line=dict(color="gold", width=2, dash="dot"),
                        editable=False,
                    )
                )
                annots.append(
                    dict(
                        x=pth_line,
                        y=1.0,
                        yref="paper",
                        text="Pth",
                        showarrow=False,
                        font=dict(color="gold", size=12),
                        yshift=26,
                    )
                )
            # Pd (Thomeer displacement pressures used in current model)
            th_mode = (res.get("thomeer_mode") or "").strip().lower() if isinstance(res, dict) else ""
            if th_mode in ("bimodal", "unimodal") and plot_mode in ("intrusion", "pcsw", "thomeer"):
                if th_mode == "bimodal":
                    pd1_line = _num(res.get("thomeer_pd1_psia"))
                    pd2_line = _num(res.get("thomeer_pd2_psia"))
                    if pd1_line is not None and pd1_line > 0:
                        shapes.append(dict(type="line", xref="x", yref="paper", x0=pd1_line, x1=pd1_line, y0=0, y1=1, line=dict(color="deepskyblue", width=2, dash="dot"), editable=False))
                        annots.append(dict(x=pd1_line, y=1.0, yref="paper", text="Pd1", showarrow=False, font=dict(color="deepskyblue", size=12), yshift=10))
                    if pd2_line is not None and pd2_line > 0:
                        shapes.append(dict(type="line", xref="x", yref="paper", x0=pd2_line, x1=pd2_line, y0=0, y1=1, line=dict(color="violet", width=2, dash="dot"), editable=False))
                        annots.append(dict(x=pd2_line, y=1.0, yref="paper", text="Pd2", showarrow=False, font=dict(color="violet", size=12), yshift=22))
                else:
                    pd_line = _num(res.get("thomeer_pd_psia"))
                    if pd_line is not None and pd_line > 0:
                        shapes.append(dict(type="line", xref="x", yref="paper", x0=pd_line, x1=pd_line, y0=0, y1=1, line=dict(color="deepskyblue", width=2, dash="dot"), editable=False))
                        annots.append(dict(x=pd_line, y=1.0, yref="paper", text="Pd", showarrow=False, font=dict(color="deepskyblue", size=12), yshift=10))
            # Backbone / Fractal proxy line
            if bb_psia is not None and np.isfinite(bb_psia) and bb_psia > 0:
                shapes.append(
                    dict(
                        type="line",
                        xref="x",
                        yref="paper",
                        x0=bb_psia,
                        x1=bb_psia,
                        y0=0,
                        y1=1,
                        line=dict(color="magenta", width=2, dash="solid"),
                        editable=False,
                    )
                )
                annots.append(
                    dict(
                        x=bb_psia,
                        y=1.0,
                        yref="paper",
                        text="Backbone",
                        showarrow=False,
                        font=dict(color="magenta", size=12),
                        yshift=54,
                    )
                )
            if shapes:
                existing_shapes = list(fig.layout.shapes) if getattr(fig.layout, "shapes", None) else []
                fig.update_layout(shapes=existing_shapes + shapes)
                for a in annots:
                    fig.add_annotation(**a)
        except Exception:
            pass
    # If correction applied, optionally overlay the original (raw) cumulative curve for reference (intrusion plot only)
    if mode == "intrusion" and conf_applied and sample.get("data_raw"):
        try:
            raw_df = pd.DataFrame(sample.get("data_raw", []))
            raw_df = _ensure_schema(raw_df)
            raw_df = raw_df.sort_values("Pressure", ascending=True)
            raw_trace = go.Scatter(
                x=raw_df["Pressure"],
                y=raw_df["CumVol"],
                mode="lines",
                name="Cumulative (raw)",
                line=dict(color="rgba(255,255,255,0.35)", dash="dot"),
            )
            fig.add_trace(raw_trace)
        except Exception:
            pass

    # Apply plot background theme (dark / light / white)
    theme = (ui or {}).get("plot_theme", "dark")
    fig = apply_plot_theme(fig, theme)
    xlog_scale = False
    try:
        xaxis = getattr(getattr(fig, "layout", None), "xaxis", None)
        xlog_scale = str(getattr(xaxis, "type", "")).lower() == "log"
    except Exception:
        xlog_scale = False
    fig = apply_robust_autoscale_x(fig, xlog=xlog_scale)
    fig = apply_smart_legend(fig)

    status_txt = status or ""
    return df.to_dict("records"), metrics_table, fig, status_txt


# ---------------------------
# Thomeer manual controls (sliders)
# - User adjusts G and Bv∞
# - Pd is auto-locked to the Closure/Conformance line when applied.
# ---------------------------
@app.callback(
    Output("slider-thomeer-g", "value"),
    Output("slider-thomeer-bvinf", "value"),
    Output("slider-thomeer-bvinf", "max"),
    Output("slider-thomeer-g", "disabled"),
    Output("slider-thomeer-bvinf", "disabled"),
    Output("thomeer-g-readout", "children"),
    Output("thomeer-bvinf-readout", "children"),
    Output("thomeer-pd-readout", "children"),
    Input("store-current-id", "data"),
    Input("store-library", "data"),
    State("store-params", "data"),
)
def update_thomeer_controls(current_id, library, params):
    params = params or DEFAULT_PARAMS
    library = library or []
    if not current_id:
        return 0.10, 5.0, 20.0, True, True, "G = -", "Bv∞ = -", "Pd = - (load a sample)"

    s = _lib_get(library, current_id)
    if not s:
        return 0.10, 5.0, 20.0, True, True, "G = -", "Bv∞ = -", "Pd = -"

    res = s.get("results", {}) or {}
    meta = s.get("meta", {}) or {}

    # If a bi-modal Thomeer fit is active, disable the unimodal sliders to avoid overwriting results.
    if res.get("thomeer_mode") == "bimodal":
        try:
            pd1 = float(res.get("thomeer_pd1_psia") or res.get("thomeer_pd_psia") or 1.0)
        except Exception:
            pd1 = 1.0
        try:
            pd2 = float(res.get("thomeer_pd2_psia") or 0.0)
        except Exception:
            pd2 = 0.0
        try:
            g1 = float(res.get("thomeer_G1") or res.get("thomeer_G") or 0.10)
        except Exception:
            g1 = 0.10
        try:
            g2 = float(res.get("thomeer_G2") or 0.50)
        except Exception:
            g2 = 0.50

        # Use total Bv∞ as the slider position (sliders are disabled anyway)
        try:
            bvinf_pct = float(res.get("thomeer_bvinf_pct") or (float(res.get("thomeer_vb_inf", 0.0)) * 100.0))
        except Exception:
            bvinf_pct = 5.0
        if not np.isfinite(bvinf_pct):
            bvinf_pct = 5.0

        max_bvinf = float(np.clip(max(bvinf_pct * 1.25, 20.0), 5.0, 100.0))
        g_val = float(np.clip(g1, 0.01, 3.0))
        bvinf_pct = float(np.clip(bvinf_pct, 0.0, max_bvinf))

        vb1_pct = res.get("thomeer_bvinf1_pct")
        vb2_pct = res.get("thomeer_bvinf2_pct")
        try:
            vb1_pct = float(vb1_pct) if vb1_pct is not None else float(res.get("thomeer_vb_inf1", 0.0)) * 100.0
        except Exception:
            vb1_pct = 0.0
        try:
            vb2_pct = float(vb2_pct) if vb2_pct is not None else float(res.get("thomeer_vb_inf2", 0.0)) * 100.0
        except Exception:
            vb2_pct = 0.0

        g_txt = f"Bimodal fit: G1={g1:.3f}, G2={g2:.3f}"
        b_txt = f"Bv∞1={vb1_pct:.2f}%, Bv∞2={vb2_pct:.2f}% (total={bvinf_pct:.2f}%)"
        closure = float(res.get("conf_pknee_psia", np.nan))
        sep = float(res.get("thomeer_pd_sep_log10", np.nan))
        qc = str(res.get("thomeer_bimodal_qc", "")).strip()
        parts = []
        if np.isfinite(closure):
            parts.append(f"Closure={closure:.3f} psia")
        parts.append(f"Pd1={pd1:.3f} psia")
        if pd2 and np.isfinite(pd2):
            parts.append(f"Pd2={pd2:.3f} psia")
        else:
            parts.append("Pd2= -")
        if np.isfinite(sep) and qc:
            parts.append(f"Δlog10(Pd)={sep:.2f} ({qc})")
        elif np.isfinite(sep):
            parts.append(f"Δlog10(Pd)={sep:.2f}")
        pd_txt = " | ".join(parts)
        return g_val, bvinf_pct, max_bvinf, True, True, g_txt, b_txt, pd_txt

    # Pd: linked to closure/conformance if applied
    pd_locked = False
    pd_psia = None
    if res.get("conf_applied") and res.get("conf_pknee_psia") is not None:
        pd_locked = True
        pd_psia = float(res.get("conf_pknee_psia"))
    elif res.get("thomeer_pd_psia") is not None:
        pd_psia = float(res.get("thomeer_pd_psia"))
    elif res.get("threshold_pressure_psia_used") is not None:
        pd_psia = float(res.get("threshold_pressure_psia_used"))
    elif meta.get("threshold_pressure_psia") is not None:
        pd_psia = float(meta.get("threshold_pressure_psia"))

    # fallback: min valid pressure
    if pd_psia is None:
        try:
            df = _ensure_schema(pd.DataFrame(s.get("data", [])))
            p = df["Pressure"].to_numpy(dtype=float)
            p = p[np.isfinite(p) & (p > 0)]
            pd_psia = float(np.nanmin(p)) if len(p) else 1.0
        except Exception:
            pd_psia = 1.0

    # G value
    try:
        g_val = float(res.get("thomeer_G", 0.10))
    except Exception:
        g_val = 0.10
    g_val = float(np.clip(g_val, 0.01, 3.0))

    # Bv∞ slider: use porosity if available, else proxy on HgSat
    phi_pct = params.get("phi_override_pct") or meta.get("porosity_pct")
    has_phi = False
    try:
        phi_frac = float(phi_pct) / 100.0
        has_phi = bool(np.isfinite(phi_frac) and phi_frac > 0)
    except Exception:
        has_phi = False

    # Pick current Bv∞ from stored results or from data max
    bvinf_pct = None
    if res.get("thomeer_bvinf_pct") is not None:
        try:
            bvinf_pct = float(res.get("thomeer_bvinf_pct"))
        except Exception:
            bvinf_pct = None
    if bvinf_pct is None and res.get("thomeer_vb_inf") is not None:
        try:
            bvinf_pct = float(res.get("thomeer_vb_inf")) * 100.0
        except Exception:
            bvinf_pct = None

    if bvinf_pct is None:
        # derive from current data
        try:
            df = _ensure_schema(pd.DataFrame(s.get("data", [])))
            df = recompute_derived(df, params)
            if has_phi:
                vb = df["HgSat"].to_numpy(dtype=float) * phi_frac
                bvinf_pct = float(np.nanmax(vb) * 100.0)
            else:
                vb = df["HgSat"].to_numpy(dtype=float)
                bvinf_pct = float(np.nanmax(vb) * 100.0)
        except Exception:
            bvinf_pct = 5.0

    if not np.isfinite(bvinf_pct):
        bvinf_pct = 5.0

    # Dynamic max for Bv∞
    if has_phi:
        max_bvinf = float(np.clip(max(phi_frac * 100.0 * 1.25, bvinf_pct * 1.25, 5.0), 5.0, 60.0))
        label = "Bv∞"
    else:
        max_bvinf = 100.0
        label = "HgSat∞ (proxy)"

    bvinf_pct = float(np.clip(bvinf_pct, 0.0, max_bvinf))

    g_txt = f"G = {g_val:.3f}"
    b_txt = f"{label} = {bvinf_pct:.2f}%"
    # Pd informational text (Pd is seeded from Threshold Pth; conformance is only for data cleaning)
    pth_note = res.get("threshold_pressure_psia_used") or meta.get("threshold_pressure_psia")
    if pth_note is not None:
        try:
            pth_note = float(pth_note)
        except Exception:
            pth_note = None

    if pth_note is not None and np.isfinite(pth_note) and abs(pd_psia - float(pth_note)) < 1e-9:
        pd_txt = f"Pd = {pd_psia:.3f} psia (from Pth)"
    else:
        pd_txt = f"Pd = {pd_psia:.3f} psia"

    return g_val, bvinf_pct, max_bvinf, False, False, g_txt, b_txt, pd_txt


@app.callback(
    Output("store-library", "data", allow_duplicate=True),
    Output("store-status", "data", allow_duplicate=True),
    Input("slider-thomeer-g", "value"),
    Input("slider-thomeer-bvinf", "value"),
    State("store-library", "data"),
    State("store-current-id", "data"),
    State("store-params", "data"),
    prevent_initial_call=True,
)
def on_thomeer_slider_change(g_val, bvinf_pct, library, current_id, params):
    if not current_id:
        return no_update, no_update

    params = params or DEFAULT_PARAMS
    library = library or []
    s = _lib_get(library, current_id)
    if not s:
        return no_update, no_update

    res = s.get("results", {}) or {}
    meta = s.get("meta", {}) or {}

    # Prevent the unimodal sliders from overwriting a bi-modal fit
    if res.get("thomeer_mode") == "bimodal":
        return no_update, no_update

    # Pd: use stored Pd; otherwise seed from Threshold (Pth). Conformance is ONLY for data cleaning.
    pd_psia = float(
        res.get("thomeer_pd_psia")
        or res.get("threshold_pressure_psia_used")
        or meta.get("threshold_pressure_psia")
        or 1.0
    )
    pd_locked = False

    try:
        G = float(g_val)
    except Exception:
        G = float(res.get("thomeer_G", 0.10) or 0.10)
    G = float(np.clip(G, 0.01, 3.0))

    try:
        bvinf_pct_f = float(bvinf_pct)
    except Exception:
        bvinf_pct_f = float(res.get("thomeer_bvinf_pct", 5.0) or 5.0)
    bvinf_pct_f = float(np.clip(bvinf_pct_f, 0.0, 100.0))

    # If nothing changed materially, avoid churn (important because sliders are programmatically updated on sample switch)
    if (
        res.get("thomeer_G") is not None and abs(float(res.get("thomeer_G")) - G) < 1e-9 and
        res.get("thomeer_bvinf_pct") is not None and abs(float(res.get("thomeer_bvinf_pct")) - bvinf_pct_f) < 1e-9 and
        res.get("thomeer_pd_psia") is not None and abs(float(res.get("thomeer_pd_psia")) - pd_psia) < 1e-9
    ):
        return no_update, no_update

    df = _ensure_schema(pd.DataFrame(s.get("data", [])))
    df = recompute_derived(df, params)

    pc = df["Pressure"].to_numpy(dtype=float)

    # vb series (bulk fraction if phi exists; else HgSat proxy)
    phi_pct = params.get("phi_override_pct") or meta.get("porosity_pct")
    vb_label = "Bv (bulk fraction)"
    try:
        phi_frac = float(phi_pct) / 100.0
        if not np.isfinite(phi_frac) or phi_frac <= 0:
            raise ValueError
        vb = df["HgSat"].to_numpy(dtype=float) * phi_frac
    except Exception:
        vb = df["HgSat"].to_numpy(dtype=float)
        vb_label = "HgSat (proxy for Bv)"

    vb_inf = bvinf_pct_f / 100.0

    # R² for current slider setting (use only valid numeric points)
    m = np.isfinite(pc) & (pc > 0) & np.isfinite(vb) & (vb >= 0)
    pc_m = pc[m]
    vb_m = vb[m]
    r2 = np.nan
    if len(pc_m) >= 5:
        vb_pred = thomeer_vb(pc_m, vb_inf, pd_psia, G)
        ss_res = float(np.nansum((vb_m - vb_pred) ** 2))
        ss_tot = float(np.nansum((vb_m - np.nanmean(vb_m)) ** 2))
        r2 = 1.0 - ss_res / ss_tot if ss_tot > 0 else np.nan

    res2 = {**res}
    res2.update({
        "thomeer_mode": "manual_sliders",
        "thomeer_vb_label": vb_label,
        "thomeer_vb_inf": float(vb_inf),
        "thomeer_bvinf_pct": float(bvinf_pct_f),
        "thomeer_pd_psia": float(pd_psia),
        "thomeer_G": float(G),
        "thomeer_r2": float(r2) if np.isfinite(r2) else np.nan,
        "thomeer_pd_locked": bool(pd_locked),
        "thomeer_done": True,
        "thomeer_at": _now_iso(),
    })

    s2 = {**s, "results": res2}
    library2 = _lib_set(library, s2)

    lock_txt = "locked" if pd_locked else "unlocked"
    msg = f"Thomeer updated ({lock_txt} Pd): Pd={pd_psia:.3f} psia, G={G:.3f}, Bv∞={bvinf_pct_f:.2f}% (R²={res2.get('thomeer_r2')})."
    return library2, msg


# ---------------------------
# Persist grid edits into the library (per-sample)
# NOTE: Dash AG Grid editing triggers callbacks via cellValueChanged.
# ---------------------------
@app.callback(
    Output("store-library", "data", allow_duplicate=True),
    Output("store-status", "data", allow_duplicate=True),
    Input("grid", "cellValueChanged"),
    State("grid", "rowData"),
    State("store-library", "data"),
    State("store-current-id", "data"),
    State("store-params", "data"),
    prevent_initial_call=True,
)
def persist_grid_edits(_evt, rowdata, library, current_id, params):
    if not current_id:
        return no_update, no_update
    library = library or []
    params = params or DEFAULT_PARAMS

    sample = _lib_get(library, current_id)
    if not sample:
        return no_update, no_update

    df = pd.DataFrame(rowdata or [])
    df = _ensure_schema(df)
    df = recompute_derived(df, params)

    sample2 = {**sample, "data": df.to_dict("records")}
    library2 = _lib_set(library, sample2)
    return library2, "Grid updated (edits saved in session)."

# ---------------------------
# Actions: QAQC / Scrub / Permeability / Radius/Diameter / Thomeer / J / Cluster / Exports
# ---------------------------
def _update_current_sample(library, current_id, updater):
    library = library or []
    if not current_id:
        return library, None
    s = _lib_get(library, current_id)
    if not s:
        return library, None
    s2 = updater(s)
    return _lib_set(library, s2), s2

@app.callback(
    Output("store-library", "data", allow_duplicate=True),
    Output("store-status", "data", allow_duplicate=True),
    Input("btn-qaqc", "n_clicks"),
    State("store-library", "data"),
    State("store-current-id", "data"),
    State("store-params", "data"),
    prevent_initial_call=True,
)
def on_qaqc(_, library, current_id, params):
    def upd(s):
        df = pd.DataFrame(s.get("data", []))
        df = _ensure_schema(df)
        dfq = apply_qaqc_flags(df)
        dfq = recompute_derived(dfq, params or DEFAULT_PARAMS)
        res = {**(s.get("results", {}) or {})}
        res.update({"qaqc_done": True, "qaqc_at": _now_iso()})

        # Also run petrophysical QAQC (direct property plausibility)
        meta = s.get("meta", {}) or {}
        try:
            res.update(petrophysical_qaqc(meta, params or DEFAULT_PARAMS, df=df, res=res))
        except Exception:
            pass
        return {**s, "data": dfq.to_dict("records"), "results": res}

    library2, s2 = _update_current_sample(library, current_id, upd)
    if not s2:
        return no_update, "No sample selected."
    dfq = pd.DataFrame(s2["data"])
    n_p = int((dfq["Flag_Pressure_Down"] == "Y").sum())
    n_c = int((dfq["Flag_Cum_Down"] == "Y").sum())
    n_i = int((dfq["Flag_Inc_Neg_Fail"] == "Y").sum())
    msg = f"QA/QC executed.\nFlag_Pressure_Down: {n_p}\nFlag_Cum_Down: {n_c}\nFlag_Inc_Neg_Fail: {n_i}"

    # Add petrophysical QAQC summary (if available)
    res_now = s2.get("results", {}) or {}
    if res_now.get("petro_qc_done"):
        pg = res_now.get("petro_qc_grade", "")
        pr = res_now.get("petro_qc_recommendation", "")
        n_pet = len(res_now.get("petro_qc_issues") or [])
        msg += f"\nPetroQC: {pg} ({pr}) — Issues: {n_pet}"
    elif (res_now.get("petro_qc_grade") or res_now.get("petro_qc_recommendation")):
        msg += f"\nPetroQC: {res_now.get('petro_qc_grade','')} ({res_now.get('petro_qc_recommendation','')})"

    return library2, msg



# ---------------------------
# PetroQC (multi-sample table)
# ---------------------------
@app.callback(
    Output("store-library", "data", allow_duplicate=True),
    Output("store-status", "data", allow_duplicate=True),
    Output("store-petroqc", "data", allow_duplicate=True),
    Input("btn-petroqc", "n_clicks"),
    State("store-library", "data"),
    State("store-params", "data"),
    prevent_initial_call=True,
)
def on_petroqc(n_clicks, library, params):
    if not n_clicks:
        raise PreventUpdate

    library = library or []
    if not library:
        return no_update, "PetroQC: no samples loaded.", no_update

    params = params or DEFAULT_PARAMS
    updated, rows, counts = run_petroqc_on_library(library, params)

    msg = (
        f"PetroQC complete — PASS: {counts.get('PASS', 0)}, "
        f"WARN: {counts.get('WARN', 0)}, FAIL: {counts.get('FAIL', 0)}."
    )
    return updated, msg, rows


@app.callback(
    Output("grid-petroqc", "rowData"),
    Input("store-petroqc", "data"),
)
def update_petroqc_grid(rows):
    return rows or []


@app.callback(
    Output("store-current-id", "data", allow_duplicate=True),
    Input("grid-petroqc", "selectedRows"),
    prevent_initial_call=True,
)
def select_sample_from_petroqc(selected_rows):
    if not selected_rows:
        raise PreventUpdate
    rid = selected_rows[0].get("id")
    if not rid:
        raise PreventUpdate
    return rid

@app.callback(
    Output("store-library", "data", allow_duplicate=True),
    Output("store-status", "data", allow_duplicate=True),
    Input("btn-scrub", "n_clicks"),
    State("store-library", "data"),
    State("store-current-id", "data"),
    State("store-params", "data"),
    prevent_initial_call=True,
)
def on_scrub(_, library, current_id, params):
    def upd(s):
        df = pd.DataFrame(s.get("data", []))
        df = scrub_data(df)
        df = recompute_derived(df, params or DEFAULT_PARAMS)
        return {**s, "data": df.to_dict("records")}

    library2, s2 = _update_current_sample(library, current_id, upd)
    if not s2:
        return no_update, "No sample selected."
    return library2, "Scrub complete: sorted, CumVol monotonic, IncVol recomputed."


@app.callback(
    Output("store-library", "data", allow_duplicate=True),
    Output("store-status", "data", allow_duplicate=True),
    Input("btn-conf", "n_clicks"),
    State("store-library", "data"),
    State("store-current-id", "data"),
    State("store-params", "data"),
    prevent_initial_call=True,
)
def on_conformance(_, library, current_id, params):
    params = params or DEFAULT_PARAMS

    def upd(s):
        df0 = pd.DataFrame(s.get("data", []))
        df0 = _ensure_schema(df0)

        pk, vv, method = detect_conformance_knee(df0, params)

        # Always record attempt (even if not found)
        res = {**(s.get("results", {}) or {})}
        res.update({"conf_method": method})

        if pk is None or vv is None:
            res.update({"conf_applied": False})
            return {**s, "results": res}

        s2 = dict(s)

        # Keep a snapshot of the original imported data (first time only)
        if not s2.get("data_raw"):
            s2["data_raw"] = s2.get("data", [])

        dfc = apply_conformance_correction(df0, pknee_psia=pk, v_conf_ml_g=vv)
        dfc = recompute_derived(dfc, params)

        # Threshold (Pth) from the raw curve (per acceleration algorithm)
        pth_psia, pth_info = compute_threshold_pressure_psia(df0, params)

        # Backbone / Fractal proxy from the conformance-corrected curve (effective saturation)
        bb_psia, bb_info = compute_backbone_pressure_psia(dfc, params)

        res.update({
            "conf_applied": True,
            "conf_method": method,
            "conf_pknee_psia": round(float(pk), 6),
            "conf_vol_ml_g": round(float(vv), 6),
        })

        # Store Pth (entry pressure) – best default candidate for Pd (NOT conformance)
        if pth_psia is not None and np.isfinite(pth_psia) and pth_psia > 0:
            res["threshold_pressure_psia_used"] = round(float(pth_psia), 6)
            if isinstance(pth_info, dict):
                res["threshold_pressure_method"] = pth_info.get("method")
                res["threshold_pressure_detail"] = pth_info.get("detail")
            res["thomeer_pd_psia"] = round(float(pth_psia), 6)

        # Store Backbone / Fractal proxy pressure
        if bb_psia is not None and np.isfinite(bb_psia) and bb_psia > 0:
            res["backbone_pressure_psia"] = round(float(bb_psia), 6)
            if isinstance(bb_info, dict):
                res["backbone_sat_eff_frac"] = bb_info.get("sat_eff_frac")

        return {**s2, "data": dfc.to_dict("records"), "results": res}

    library2, s2 = _update_current_sample(library, current_id, upd)
    if not s2:
        return no_update, "No sample selected."

    res = s2.get("results", {}) or {}
    if not res.get("conf_applied"):
        return library2, f"Conformance correction: knee not found ({res.get('conf_method')})."

    return library2, (
        "Conformance correction applied "
        f"({res.get('conf_method')}): Pc_knee={res.get('conf_pknee_psia')} psia, "
        f"V_conf={res.get('conf_vol_ml_g')} mL/g. "
        "Tip: visually verify the vertical line on the plot."
    )


@app.callback(
    Output("store-library", "data", allow_duplicate=True),
    Output("store-status", "data", allow_duplicate=True),
    Input("btn-conf-undo", "n_clicks"),
    State("store-library", "data"),
    State("store-current-id", "data"),
    State("store-params", "data"),
    prevent_initial_call=True,
)
def on_conformance_undo(_, library, current_id, params):
    params = params or DEFAULT_PARAMS

    def upd(s):
        if not s.get("data_raw"):
            return s

        df0 = pd.DataFrame(s.get("data_raw", []))
        df0 = _ensure_schema(df0)
        df0 = recompute_derived(df0, params)

        s2 = dict(s)
        s2["data"] = df0.to_dict("records")

        res = {**(s2.get("results", {}) or {})}
        for k in [
            "conf_applied",
            "conf_pknee_psia",
            "conf_vol_ml_g",
            "conf_method",
            "backbone_pressure_psia",
            "backbone_sat_eff_frac",
        ]:
            res.pop(k, None)

        # Recompute Threshold and Backbone on the restored raw curve
        pth_psia, pth_info = compute_threshold_pressure_psia(df0, params)
        bb_psia, bb_info = compute_backbone_pressure_psia(df0, params)

        if pth_psia is not None and np.isfinite(pth_psia) and pth_psia > 0:
            res["threshold_pressure_psia_used"] = round(float(pth_psia), 6)
            if isinstance(pth_info, dict):
                res["threshold_pressure_method"] = pth_info.get("method")
                res["threshold_pressure_detail"] = pth_info.get("detail")
            res["thomeer_pd_psia"] = round(float(pth_psia), 6)

        if bb_psia is not None and np.isfinite(bb_psia) and bb_psia > 0:
            res["backbone_pressure_psia"] = round(float(bb_psia), 6)
            if isinstance(bb_info, dict):
                res["backbone_sat_eff_frac"] = bb_info.get("sat_eff_frac")

        s2["results"] = res
        return s2

    library2, s2 = _update_current_sample(library, current_id, upd)
    if not s2:
        return no_update, "No sample selected."
    if not s2.get("data_raw"):
        return library2, "Undo conformance: no raw snapshot stored (apply correction once first)."
    return library2, "Conformance correction undone (restored original data)."


# ---------------------------
# Draggable closure line handler
# When the user drags the vertical line (shape[0]) on the main graph,
# we recompute the conformance/closure correction from the stored baseline (data_raw)
# and tie Thomeer Pd to that same pressure.
# ---------------------------
@app.callback(
    Output("store-library", "data", allow_duplicate=True),
    Output("store-status", "data", allow_duplicate=True),
    Input("graph", "relayoutData"),
    State("store-library", "data"),
    State("store-current-id", "data"),
    State("store-params", "data"),
    prevent_initial_call=True,
)
def on_drag_closure_line(relayout, library, current_id, params):
    if not current_id or not relayout or not isinstance(relayout, dict):
        return no_update, no_update

    # We only care about edits to the first shape (our closure line)
    x_new = None
    for k in ("shapes[0].x0", "shapes[0].x1"):
        if k in relayout:
            x_new = relayout.get(k)
            break

    # Some plotly versions may send the full shapes list
    if x_new is None and "shapes" in relayout and isinstance(relayout.get("shapes"), list) and relayout["shapes"]:
        s0 = relayout["shapes"][0]
        if isinstance(s0, dict) and "x0" in s0:
            x_new = s0.get("x0")

    if x_new is None:
        return no_update, no_update

    try:
        pk = float(x_new)
    except Exception:
        return no_update, no_update

    if not np.isfinite(pk) or pk <= 0:
        return no_update, no_update

    params = params or DEFAULT_PARAMS
    library = library or []
    sample = _lib_get(library, current_id)
    if not sample:
        return no_update, no_update

    res = sample.get("results", {}) or {}

    # Only apply if closure/conformance mode has been activated at least once.
    # (We need a baseline snapshot to avoid cumulative shifts stacking.)
    base_rows = sample.get("data_raw")
    if not base_rows:
        return no_update, no_update

    df_base = _ensure_schema(pd.DataFrame(base_rows))
    v_conf = cumvol_at_pressure(df_base, pk)
    if v_conf is None:
        return no_update, no_update

    df_corr = apply_conformance_correction(df_base, pknee_psia=pk, v_conf_ml_g=v_conf)
    df_corr = recompute_derived(df_corr, params)

    # Recompute Threshold (Pth) and Backbone proxy on corrected curve
    pth_psia, pth_info = compute_threshold_pressure_psia(df_corr, params)
    bb_psia, bb_info = compute_backbone_pressure_psia(df_corr, params)

    res2 = {**res}
    res2.update({
        "conf_applied": True,
        "conf_method": "manual_drag",
        "conf_pknee_psia": round(float(pk), 6),
        "conf_vol_ml_g": round(float(v_conf), 6),
    })

    # Store Pth (entry pressure) – best default candidate for Pd (NOT conformance)
    if pth_psia is not None and np.isfinite(pth_psia) and pth_psia > 0:
        res2["threshold_pressure_psia_used"] = round(float(pth_psia), 6)
        if isinstance(pth_info, dict):
            res2["threshold_pressure_method"] = pth_info.get("method")
            res2["threshold_pressure_detail"] = pth_info.get("detail")
        # Use Pth as Pd seed for Thomeer calculations
        res2["thomeer_pd_psia"] = round(float(pth_psia), 6)

    # Store Backbone / Fractal proxy pressure
    if bb_psia is not None and np.isfinite(bb_psia) and bb_psia > 0:
        res2["backbone_pressure_psia"] = round(float(bb_psia), 6)
        if isinstance(bb_info, dict):
            res2["backbone_sat_eff_frac"] = bb_info.get("sat_eff_frac")

    sample2 = dict(sample)
    sample2["data"] = df_corr.to_dict("records")
    sample2["results"] = res2

    library2 = _lib_set(library, sample2)
    msg = (
        f"Closure line moved → Pc={res2['conf_pknee_psia']} psia | "
        f"Vconf={res2['conf_vol_ml_g']} mL/g. Pd updated for Thomeer."
    )
    return library2, msg


@app.callback(
    Output("store-library", "data", allow_duplicate=True),
    Output("store-status", "data", allow_duplicate=True),
    Input("btn-radius", "n_clicks"),
    Input("btn-diam", "n_clicks"),
    State("store-library", "data"),
    State("store-current-id", "data"),
    State("store-params", "data"),
    prevent_initial_call=True,
)
def on_radius_or_diam(_, __, library, current_id, params):
    # recompute_derived already populates both r_um and d_um
    def upd(s):
        df = pd.DataFrame(s.get("data", []))
        df = _ensure_schema(df)
        df = recompute_derived(df, params or DEFAULT_PARAMS)
        return {**s, "data": df.to_dict("records")}

    library2, s2 = _update_current_sample(library, current_id, upd)
    if not s2:
        return no_update, "No sample selected."
    return library2, "Pore throat radius/diameter computed (r_um, d_um)."


@app.callback(
    Output("store-library", "data", allow_duplicate=True),
    Output("store-status", "data", allow_duplicate=True),
    Input("btn-add-row", "n_clicks"),
    State("store-library", "data"),
    State("store-current-id", "data"),
    State("store-params", "data"),
    prevent_initial_call=True,
)
def on_add_row(_, library, current_id, params):
    params = params or DEFAULT_PARAMS

    def upd(s):
        df = pd.DataFrame(s.get("data", []))
        df = _ensure_schema(df)

        # Build an empty row aligned to schema
        row = {col: ("" if col in FLAG_COLS else np.nan) for col in df.columns}
        df2 = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
        df2 = recompute_derived(df2, params)
        return {**s, "data": df2.to_dict("records")}

    library2, s2 = _update_current_sample(library, current_id, upd)
    if not s2:
        return no_update, "No sample selected."
    return library2, "Row added."

@app.callback(
    Output("download-csv", "data"),
    Input("btn-export-csv", "n_clicks"),
    State("store-library", "data"),
    State("store-current-id", "data"),
    prevent_initial_call=True,
)
def export_csv(_, library, current_id):
    sample = _lib_get(library or [], current_id)
    if not sample:
        return no_update
    df = pd.DataFrame(sample.get("data", []))
    df = _ensure_schema(df)

    out = df[REQUIRED_COLS + FLAG_COLS].copy()
    fn = f"MICP_{sample.get('sample_id','sample')}.csv".replace(" ", "_")
    return dcc.send_data_frame(out.to_csv, fn, index=False)


@app.callback(
    Output("download-petroqc", "data"),
    Input("btn-export-petroqc", "n_clicks"),
    State("store-petroqc", "data"),
    State("store-library", "data"),
    State("store-params", "data"),
    prevent_initial_call=True,
)
def export_petroqc(n_clicks, petroqc_rows, library, params):
    if not n_clicks:
        raise PreventUpdate

    rows = petroqc_rows or []
    if not rows and (library or []):
        try:
            _, rows, _ = run_petroqc_on_library(library or [], params or DEFAULT_PARAMS)
        except Exception:
            rows = []

    if not rows:
        return no_update

    df = pd.DataFrame(rows)

    # Order + rename to match the PetroQC header spec
    col_order = [
        "well",
        "core",
        "sample_id",
        "depth_m",
        "porosity_pct",
        "permeability_md",
        "k_pnm_md",
        "pnm_ci",
        "threshold_pressure_psia",
        "tortuosity_factor",
        "bulk_density_g_ml",
        "skeletal_density_g_ml",
        "stem_volume_used_pct",
        "max_shg_sat_pct",
        "grain_density_diff_g_ml",
        "conformance_vol_pct",
        "qc_flag",
        "qc_reasons",
        "exclude_from_shm",
    ]
    keep = [c for c in col_order if c in df.columns]
    df = df[keep].copy()

    rename = {
        "well": "WELL",
        "core": "CORE",
        "sample_id": "Sample_ID",
        "depth_m": "DEPTH_m",
        "porosity_pct": "Porosity_pct",
        "permeability_md": "Permeability_mD",
        "k_pnm_md": "PNM_k_mD",
        "pnm_ci": "PNM_CI",
        "threshold_pressure_psia": "Threshold_Pressure_psia",
        "tortuosity_factor": "Tortuosity_factor",
        "bulk_density_g_ml": "Bulk_Density_g_ml",
        "skeletal_density_g_ml": "Apparent_Skeletal_Density_g_ml",
        "stem_volume_used_pct": "Stem_Volume_Used_pct",
        "max_shg_sat_pct": "Max_SHg_Saturation_pct",
        "grain_density_diff_g_ml": "Grain_Density_Diff_g_ml",
        "conformance_vol_pct": "Conformance_Vol_pct",
        "qc_flag": "QC_Flag",
        "qc_reasons": "QC_Reasons",
        "exclude_from_shm": "Exclude_from_SHM",
    }
    df = df.rename(columns=rename)

    # Filename
    well = "PetroQC"
    try:
        w0 = str(rows[0].get("well") or "").strip()
        if w0:
            well = w0
    except Exception:
        pass

    stamp = _dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    fn = f"PetroQC_{well}_{stamp}.csv".replace(" ", "_")

    return dcc.send_data_frame(df.to_csv, fn, index=False)


@app.callback(
    Output("store-library", "data", allow_duplicate=True),
    Output("store-status", "data", allow_duplicate=True),
    Input("btn-perm", "n_clicks"),
    State("store-library", "data"),
    State("store-current-id", "data"),
    State("store-params", "data"),
    prevent_initial_call=True,
)
def on_perm(_, library, current_id, params):
    params = params or DEFAULT_PARAMS

    def upd(s):
        df = pd.DataFrame(s.get("data", []))
        df = _ensure_schema(df)
        df = recompute_derived(df, params)
        meta = s.get("meta", {}) or {}
        res = s.get("results", {}) or {}

        res_sw = compute_swanson(df, params, meta)
        res_wl = compute_winland_k(df, params, meta, res)

        # Merge results
        res2 = {**res, **res_sw, **res_wl}

        # Rock type classification
        r35 = res2.get("r35_um")
        res2["rock_type"] = rock_type_from_r35(r35, params)

        return {**s, "data": df.to_dict("records"), "results": res2}

    library2, s2 = _update_current_sample(library, current_id, upd)
    if not s2:
        return no_update, "No sample selected."
    res = s2.get("results", {})
    msg = "Permeability calc done."
    if "swanson_error" in res:
        msg += f"\nSwanson: {res['swanson_error']}"
    if "winland_error" in res:
        msg += f"\nWinland: {res['winland_error']}"
    if "winland_warning" in res:
        msg += f"\nWinland: {res['winland_warning']}"
    if "r35_um" in res:
        msg += f"\nr35_um: {res['r35_um']}"
    if "rock_type" in res:
        msg += f"\nRockType: {res['rock_type']}"
    return library2, msg

@app.callback(
    Output("store-library", "data", allow_duplicate=True),
    Output("store-status", "data", allow_duplicate=True),
    Input("btn-thomeer", "n_clicks"),
    State("store-library", "data"),
    State("store-current-id", "data"),
    State("store-params", "data"),
    prevent_initial_call=True,
)
def on_thomeer(_, library, current_id, params):
    params = params or DEFAULT_PARAMS

    def upd(s):
        df = pd.DataFrame(s.get("data", []))
        df = _ensure_schema(df)
        df = recompute_derived(df, params)
        meta = s.get("meta", {}) or {}
        res = s.get("results", {}) or {}        # Pd seed: prefer stored Pd; else Threshold (Pth); else a simple data-based guess
        pd_fixed = None

        # 1) Previously stored Pd (manual/previous fit)
        try:
            if res.get("thomeer_pd_psia") is not None:
                pd_fixed = float(res.get("thomeer_pd_psia"))
        except Exception:
            pd_fixed = None

        # 2) Threshold (Pth) from results (preferred)
        if pd_fixed is None:
            try:
                if res.get("threshold_pressure_psia_used") is not None:
                    pd_fixed = float(res.get("threshold_pressure_psia_used"))
            except Exception:
                pd_fixed = None

        # 3) Threshold from parsed Micromeritics report (if present)
        if pd_fixed is None:
            try:
                if meta.get("threshold_pressure_psia") is not None:
                    pd_fixed = float(meta.get("threshold_pressure_psia"))
            except Exception:
                pd_fixed = None

        # 4) Guess Pd: first Pc where HgSat reaches 1% of its max (fallback)
        if pd_fixed is None:
            try:
                hg = df["HgSat"].to_numpy(dtype=float)
                pc_arr = df["Pressure"].to_numpy(dtype=float)
                m = np.isfinite(hg) & np.isfinite(pc_arr) & (pc_arr > 0)
                hg = hg[m]
                pc_arr = pc_arr[m]
                if len(hg) >= 3:
                    hmax = float(np.nanmax(hg))
                    idx = int(np.argmax(hg >= 0.01 * hmax))
                    pd_fixed = float(pc_arr[idx])
            except Exception:
                pd_fixed = None

        if pd_fixed is None:
            pd_fixed = float(np.nanmin(df["Pressure"].to_numpy(dtype=float)[df["Pressure"].to_numpy(dtype=float) > 0]))



        fit = fit_thomeer_fixed_pd(df, pd_fixed, params, meta)
        res2 = {**res, **fit}
        # Auto hint: possible bi-modal pore system (see compute_bimodal_flags)
        try:
            res2.update(compute_bimodal_flags(df, params, res2))
        except Exception:
            pass

        # Update Winland (macro-normalized for bi-modal) + rock type
        try:
            res2.update(compute_winland_k(df, params, meta, res2))
            res2["rock_type"] = rock_type_from_r35(res2.get("r35_um"), params)
        except Exception:
            pass


        if "thomeer_error" not in res2:
            res2.update({"thomeer_done": True, "thomeer_at": _now_iso()})
        return {**s, "data": df.to_dict("records"), "results": res2}

    library2, s2 = _update_current_sample(library, current_id, upd)
    if not s2:
        return no_update, "No sample selected."
    res = s2.get("results", {})
    if "thomeer_error" in res:
        return library2, f"Thomeer: {res['thomeer_error']}"
    msg = (
        "Thomeer fit complete.\n"
        f"Pd(psia): {res.get('thomeer_pd_psia')}\n"
        f"G: {res.get('thomeer_G')}\n"
        f"Vb∞: {res.get('thomeer_vb_inf')}\n"
        f"R²: {res.get('thomeer_r2')}"
    )
    if res.get("bimodal_hint"):
        msg += "\nPossible bi-modal pore system → try Fit Thomeer (Bimodal)."
    return library2, msg



@app.callback(
    Output("store-library", "data", allow_duplicate=True),
    Output("store-status", "data", allow_duplicate=True),
    Input("btn-thomeer-bimodal", "n_clicks"),
    Input("btn-bimodal-hint-run", "n_clicks"),
    State("store-library", "data"),
    State("store-current-id", "data"),
    State("store-params", "data"),
    prevent_initial_call=True,
)
def on_thomeer_bimodal(n_clicks_main, n_clicks_hint, library, current_id, params):
    """Fit Thomeer in *true bimodal* mode (macro + micro), after closure correction.

    Design intent:
      1) Keep **closure/conformance** as a *prior correction* (volume correction on the curve).
      2) Fit **two pore systems** with **Pd1 and Pd2 both physical** (NOT Pd1=closure).
      3) Apply a QC rule based on the **expected pressure separation** between pore systems.
         If separation is too small, the bimodal split becomes degenerate → keep unimodal.

    The QC threshold is controlled by:
        params['bimodal_pd_sep_min_log10']   (default ~0.8 decades; ratio ~6.3)
    """
    if not (n_clicks_main or n_clicks_hint):
        return dash.no_update, dash.no_update

    def upd(s):
        df = pd.DataFrame(s.get("data", []))
        if df.empty:
            s.setdefault("results", {})["thomeer_error"] = "No data"
            return s

        df = _ensure_schema(df)
        df = recompute_derived(df, params)

        meta = s.setdefault("meta", {})
        res_store = s.setdefault("results", {})

        # --- 1) Ensure closure/conformance correction is applied first ---
        if not bool(meta.get("conf_applied", False)):
            pknee_psia, vol_at, _ = detect_conformance_knee(df, params)
            if (pknee_psia is not None) and (vol_at is not None) and np.isfinite(pknee_psia) and np.isfinite(vol_at):
                df_corr = apply_conformance_correction(df, pknee_psia, vol_at)

                # Preserve a copy of the original (raw) data once.
                s["data_raw"] = s.get("data_raw") or s.get("data") or []
                s["data"] = df_corr.to_dict("records")

                # Update meta
                meta["conf_applied"] = True
                meta["conf_pknee_psia"] = float(pknee_psia)
                meta["conf_vol_ml_g"] = float(vol_at)

                # Work on corrected curve downstream
                df = df_corr
            else:
                # We can still proceed, but warn in results
                res_store["conf_warning"] = "Closure knee not detected; bimodal fit proceeded without conformance correction."

        closure_psia = float(meta.get("conf_pknee_psia", np.nan))

        # --- 2) Build vb(pc) in bulk-volume fraction space ---
        poro_pct = float(meta.get("porosity_pct", np.nan))
        phi_frac = poro_pct / 100.0 if (np.isfinite(poro_pct) and poro_pct > 0) else 1.0

        pc = df["Pressure"].values  # Pc in psia (MICP)
        vb = df["HgSat"].values * phi_frac

        # --- 3) Fit true bimodal (Pd1 and Pd2 both free) ---
        res = fit_thomeer_bimodal(pc, vb, vb_upper=phi_frac if phi_frac > 0 else None)
        if res.get("thomeer_error"):
            res_store["thomeer_error"] = res.get("thomeer_error")
            return s

        # --- 4) QC: pressure separation between Pd1 and Pd2 ---
        pd1 = float(res.get("thomeer_pd1_psia", np.nan))
        pd2 = float(res.get("thomeer_pd2_psia", np.nan))
        sep_log10 = np.log10(pd2 / pd1) if (np.isfinite(pd1) and np.isfinite(pd2) and pd1 > 0 and pd2 > 0) else np.nan

        min_sep = float(params.get("bimodal_pd_sep_min_log10", 0.8))
        qc_pass = bool(np.isfinite(sep_log10) and sep_log10 >= min_sep)

        res["thomeer_pd_sep_log10"] = float(sep_log10) if np.isfinite(sep_log10) else np.nan
        res["thomeer_bimodal_qc"] = "PASS" if qc_pass else "FAIL"

        # keep closure visible alongside Pd1/Pd2 in UI
        if np.isfinite(closure_psia):
            res["conf_pknee_psia"] = closure_psia

        if not qc_pass:
            # Store candidate (scalar-only) for traceability, but keep unimodal as the active model.
            res_store["thomeer_bimodal_candidate"] = {
                "pd1_psia": pd1,
                "pd2_psia": pd2,
                "G1": float(res.get("thomeer_G1", np.nan)),
                "G2": float(res.get("thomeer_G2", np.nan)),
                "vb_inf1": float(res.get("thomeer_vb_inf1", np.nan)),
                "vb_inf2": float(res.get("thomeer_vb_inf2", np.nan)),
                "macro_frac": float(res.get("thomeer_macro_frac", np.nan)),
                "r2": float(res.get("thomeer_r2", np.nan)),
                "sep_log10": float(res.get("thomeer_pd_sep_log10", np.nan)),
                "min_sep_log10": min_sep,
            }
            res_store["thomeer_bimodal_qc"] = "FAIL"
            res_store["thomeer_pd_sep_log10"] = float(sep_log10) if np.isfinite(sep_log10) else np.nan

            # Refit unimodal on the same (closure-corrected) curve.
            # NOTE: In GeoPore Analytics v1.8.x the unimodal helper is
            # `fit_thomeer(df, params, meta)`. An older call signature
            # (`fit_thomeer(pc, vb, vb_upper=...)`) slipped in during a refactor
            # and triggers:
            #   TypeError: fit_thomeer() got an unexpected keyword argument 'vb_upper'
            # We already have the corrected DataFrame (`df`) plus `params` and `meta`,
            # so use the supported signature.
            uni = fit_thomeer(df, params, meta)
            if uni and (not uni.get("thomeer_error")):
                res_store.update(uni)
                # Preserve the QC info (so the UI can show why bimodal wasn't adopted)
                res_store["thomeer_bimodal_qc"] = "FAIL"
                res_store["thomeer_pd_sep_log10"] = float(sep_log10) if np.isfinite(sep_log10) else np.nan
                # Keep hints updated
                res_store.update(_bimodal_hints(df, params, res_store))
            else:
                # If unimodal fails for some reason, keep the bimodal result but flagged.
                res_store.update(res)
                res_store.update(_bimodal_hints(df, params, res_store))
            return s

        # QC passed → adopt bimodal
        res_store.update(res)
        res_store.update(_bimodal_hints(df, params, res_store))
        return s

    library2, sample2 = _update_current_sample(library, current_id, upd)

    # Build a friendly status message
    r = (sample2 or {}).get("results", {}) if isinstance(sample2, dict) else {}
    closure_psia = float(r.get("conf_pknee_psia", np.nan))
    if r.get("thomeer_mode") == "bimodal":
        pd1 = float(r.get("thomeer_pd1_psia", np.nan))
        pd2 = float(r.get("thomeer_pd2_psia", np.nan))
        sep = float(r.get("thomeer_pd_sep_log10", np.nan))
        qc = r.get("thomeer_bimodal_qc", "PASS")
        msg = (
            "Thomeer bimodal fit complete (closure-corrected).\n"
            f"Closure Pc: {closure_psia:.3f} psia\n"
            f"Pd1 (macro): {pd1:.3f} psia\n"
            f"Pd2 (micro): {pd2:.3f} psia\n"
            f"Δlog10(Pd): {sep:.2f} ({qc})\n"
            f"Macro frac: {float(r.get('thomeer_macro_frac', np.nan)):.2f}\n"
            f"R²: {float(r.get('thomeer_r2', np.nan)):.3f}"
        )
    else:
        cand = r.get("thomeer_bimodal_candidate", {}) if isinstance(r.get("thomeer_bimodal_candidate", {}), dict) else {}
        sep = cand.get("sep_log10", np.nan)
        min_sep = cand.get("min_sep_log10", params.get("bimodal_pd_sep_min_log10", 0.8))
        msg = (
            "Bimodal Thomeer candidate failed QC → keeping unimodal.\n"
            f"Reason: Δlog10(Pd)={float(sep):.2f} < {float(min_sep):.2f} decades\n"
            f"Candidate Pd1={float(cand.get('pd1_psia', np.nan)):.3f} psia, Pd2={float(cand.get('pd2_psia', np.nan)):.3f} psia\n"
            f"Unimodal Pd={float(r.get('thomeer_pd_psia', np.nan)):.3f} psia, R²={float(r.get('thomeer_r2', np.nan)):.3f}"
        )

    return library2, msg
def on_pnm_fast(_, library, current_id, params):
    """Run lightweight PNM (fast) for the current sample."""
    params = params or DEFAULT_PARAMS

    def upd(s):
        df = pd.DataFrame(s.get("data", []))
        df = _ensure_schema(df)
        df = recompute_derived(df, params)
        meta = s.get("meta", {}) or {}
        res = s.get("results", {}) or {}

        try:
            pnm = compute_pnm_fast(df, params, meta, res)
        except Exception as e:
            pnm = {"pnm_error": str(e)}

        res2 = {**res, **pnm}
        if "pnm_error" not in res2:
            res2.update({"pnm_done": True, "pnm_at": _now_iso()})
        return {**s, "data": df.to_dict("records"), "results": res2}

    library2, s2 = _update_current_sample(library, current_id, upd)
    if not s2:
        return no_update, "No sample selected."
    res = s2.get("results", {}) or {}
    if res.get("pnm_error"):
        return library2, f"PNM: {res.get('pnm_error')}"
    msg = (
        "PNM (fast) complete.\n"
        f"k_PNM (mD): {res.get('k_pnm_md')}\n"
        f"CI (k_PNM/k_core): {res.get('pnm_ci', '')}\n"
        f"z: {res.get('pnm_z')}, tau_eff: {res.get('pnm_tau_eff')}"
    )
    return library2, msg




@app.callback(
    Output("store-library", "data", allow_duplicate=True),
    Output("store-status", "data", allow_duplicate=True),
    Input("btn-jfunc", "n_clicks"),
    State("store-library", "data"),
    State("store-current-id", "data"),
    State("store-params", "data"),
    prevent_initial_call=True,
)
def on_jfunc(_, library, current_id, params):
    params = params or DEFAULT_PARAMS

    def upd(s):
        df = pd.DataFrame(s.get("data", []))
        df = _ensure_schema(df)
        df = recompute_derived(df, params)
        meta = s.get("meta", {}) or {}
        res = s.get("results", {}) or {}

        out = compute_j_function(df, params, meta, res)
        if "J" in out:
            df["J"] = out["J"]
        res2 = {**res, **{k: v for k, v in out.items() if k != "J"}}
        return {**s, "data": df.to_dict("records"), "results": res2}

    library2, s2 = _update_current_sample(library, current_id, upd)
    if not s2:
        return no_update, "No sample selected."
    res = s2.get("results", {})
    if "j_error" in res:
        return library2, f"J-function: {res['j_error']}"
    return library2, "J-function computed (column 'J' added)."

@app.callback(
    Output("store-library", "data", allow_duplicate=True),
    Output("store-status", "data", allow_duplicate=True),
    Input("btn-cluster", "n_clicks"),
    State("store-library", "data"),
    State("store-params", "data"),
    prevent_initial_call=True,
)
def on_cluster(_, library, params):
    params = params or DEFAULT_PARAMS
    library2 = cluster_library(library or [], params)
    return library2, "Clustering complete (results.cluster assigned where possible)."

# ---- Exports
@app.callback(
    Output("download-shf", "data"),
    Input("btn-export-shf", "n_clicks"),
    State("store-library", "data"),
    State("store-current-id", "data"),
    State("store-params", "data"),
    prevent_initial_call=True,
)
def export_shf(_, library, current_id, params):
    sample = _lib_get(library or [], current_id)
    if not sample:
        return no_update
    df = pd.DataFrame(sample.get("data", []))
    df = _ensure_schema(df)
    df = recompute_derived(df, params or DEFAULT_PARAMS)

    out = df[["Pressure", "Sw", "Pc_res_pa", "Height_m"]].copy()
    out.columns = ["Pc_psia", "Sw_frac", "Pc_res_Pa", "Height_m"]

    fn = f"SHF_{sample.get('sample_id','sample')}.csv".replace(" ", "_")
    return dcc.send_data_frame(out.to_csv, fn, index=False)

@app.callback(
    Output("download-thomeer", "data"),
    Input("btn-export-thomeer", "n_clicks"),
    State("store-library", "data"),
    prevent_initial_call=True,
)
def export_thomeer_params(_, library):
    library = library or []
    rows = []
    for s in library:
        res = s.get("results", {}) or {}
        mode = res.get("thomeer_mode")
        is_bi = (mode == "bimodal")

        rows.append({
            "sample_id": s.get("sample_id"),
            "filename": s.get("filename"),
            "well": s.get("well"),

            # Always-present (legacy/unimodal) keys
            "thomeer_mode": mode,
            "thomeer_pd_psia": res.get("thomeer_pd_psia"),
            "thomeer_G": res.get("thomeer_G"),
            "thomeer_vb_inf": res.get("thomeer_vb_inf"),
            "thomeer_r2": res.get("thomeer_r2"),

            # Bi-modal expansion (macro + micro). Filled only when mode == 'bimodal'
            "thomeer_pd1_psia": res.get("thomeer_pd1_psia") if is_bi else None,
            "thomeer_G1": res.get("thomeer_G1") if is_bi else None,
            "thomeer_vb_inf1": res.get("thomeer_vb_inf1") if is_bi else None,
            "thomeer_pd2_psia": res.get("thomeer_pd2_psia") if is_bi else None,
            "thomeer_G2": res.get("thomeer_G2") if is_bi else None,
            "thomeer_vb_inf2": res.get("thomeer_vb_inf2") if is_bi else None,
            "thomeer_vb_inf_total": res.get("thomeer_vb_inf_total") if is_bi else None,
            "thomeer_macro_frac": res.get("thomeer_macro_frac") if is_bi else None,
        })
    df = pd.DataFrame(rows)
    fn = f"thomeer_params_{_dt.datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
    return dcc.send_data_frame(df.to_csv, fn, index=False)

@app.callback(
    Output("download-report", "data"),
    Output("store-status", "data", allow_duplicate=True),
    Input("btn-report", "n_clicks"),
    State("store-library", "data"),
    State("store-current-id", "data"),
    State("store-params", "data"),
    prevent_initial_call=True,
)
def export_report(_, library, current_id, params):
    library = library or []
    sample = _lib_get(library, current_id)
    if not sample:
        return no_update, "No sample selected."

    # Ensure some key results exist
    df = pd.DataFrame(sample.get("data", []))
    df = _ensure_schema(df)
    df = recompute_derived(df, params or DEFAULT_PARAMS)
    meta = sample.get("meta", {}) or {}
    res = sample.get("results", {}) or {}

    # Ensure decision fields exist
    res.setdefault("sample_decision", "PENDING")
    res.setdefault("discard_reason", "")
    res.setdefault("decision_at", None)

    # Ensure petrophysical QAQC exists (for older projects)
    if (not res.get("petro_qc_done")) and (not res.get("petro_qc_grade")):
        try:
            res.update(petrophysical_qaqc(meta, params or DEFAULT_PARAMS, df=df, res=res))
        except Exception:
            pass

    # Compute r35/rock type if missing
    if "r35_um" not in res:
        res.update(compute_winland_k(df, params or DEFAULT_PARAMS, meta, res))
        res["rock_type"] = rock_type_from_r35(res.get("r35_um"), params or DEFAULT_PARAMS)
    # Compute thomeer if missing (Pd locked to Closure if available)
    if "thomeer_pd_psia" not in res and "thomeer_error" not in res:
        pd_fixed = None
        try:
            if res.get("conf_applied") and res.get("conf_pknee_psia") is not None:
                pd_fixed = float(res.get("conf_pknee_psia"))
        except Exception:
            pd_fixed = None
        if pd_fixed is None:
            try:
                if res.get("threshold_pressure_psia_used") is not None:

                    pd_fixed = float(res.get("threshold_pressure_psia_used"))

                elif meta.get("threshold_pressure_psia") is not None:

                    pd_fixed = float(meta.get("threshold_pressure_psia"))
            except Exception:
                pd_fixed = None
        if pd_fixed is None:
            # fallback: smallest positive Pc
            try:
                p = df["Pressure"].to_numpy(dtype=float)
                p = p[np.isfinite(p) & (p > 0)]
                pd_fixed = float(np.nanmin(p)) if len(p) else 1.0
            except Exception:
                pd_fixed = 1.0

        res.update(fit_thomeer_fixed_pd(df, pd_fixed, params or DEFAULT_PARAMS, meta))

    sample2 = {**sample, "data": df.to_dict("records"), "results": res}
    library2 = _lib_set(library, sample2)

    pdf_bytes = build_pdf_report_bytes(sample2, params or DEFAULT_PARAMS, library2)
    fn = f"MICP_Report_{sample2.get('sample_id','sample')}.pdf".replace(" ", "_")
    return dcc.send_bytes(lambda b: b.write(pdf_bytes), fn), f"Report generated: {fn}"

# ---------------------------
# Visualization buttons -> set UI state
# ---------------------------
@app.callback(
    Output("store-ui", "data", allow_duplicate=True),
    Input("btn-toggle-xlog", "n_clicks"),
    State("store-ui", "data"),
    prevent_initial_call=True,
)
def toggle_xlog(_, ui):
    ui = ui or {"plot_mode": "intrusion", "xlog": True, "overlay_inc": False}
    ui["xlog"] = not bool(ui.get("xlog", True))
    return ui

@app.callback(
    Output("store-ui", "data", allow_duplicate=True),
    Input("btn-toggle-overlay", "n_clicks"),
    State("store-ui", "data"),
    prevent_initial_call=True,
)
def toggle_overlay(_, ui):
    ui = ui or {"plot_mode": "intrusion", "xlog": True, "overlay_inc": False}
    ui["overlay_inc"] = not bool(ui.get("overlay_inc", False))
    return ui

def _set_mode(ui, mode):
    ui = ui or {"plot_mode": "intrusion", "xlog": True, "overlay_inc": False}
    ui["plot_mode"] = mode
    return ui

@app.callback(
    Output("store-ui", "data", allow_duplicate=True),
    Input("btn-viz-intrusion", "n_clicks"),
    Input("btn-viz-pcsw", "n_clicks"),
    Input("btn-viz-psd", "n_clicks"),
    Input("btn-viz-thomeer", "n_clicks"),
    Input("btn-viz-shf", "n_clicks"),
    Input("btn-viz-winland", "n_clicks"),
    Input("btn-viz-pnm3d", "n_clicks"),
        Input("btn-ms-pc-overlay", "n_clicks"),
        Input("btn-ms-psd-compare", "n_clicks"),
        Input("btn-ms-cum-intrusion", "n_clicks"),
        Input("btn-ms-phik", "n_clicks"),
        Input("btn-ms-jfunc", "n_clicks"),
        Input("btn-ms-gpd", "n_clicks"),
        Input("btn-ms-petrolog", "n_clicks"),
        Input("btn-ms-kprof", "n_clicks"),
        Input("btn-ms-hfu", "n_clicks"),
        Input("btn-ms-shm", "n_clicks"),
        Input("btn-ms-kpnm", "n_clicks"),
        Input("btn-ms-ci-log", "n_clicks"),
        Input("btn-ms-reset", "n_clicks"),

    State("store-ui", "data"),
    prevent_initial_call=True,
)
def set_viz(*args):
    ui = args[-1] if args else None
    ctx = callback_context
    if ctx and ctx.triggered:
        trig = ctx.triggered[0]["prop_id"].split(".")[0]
    else:
        trig = ""
    mapping = {
        "btn-viz-intrusion": "intrusion",
        "btn-viz-pcsw": "pcsw",
        "btn-viz-psd": "psd",
        "btn-viz-thomeer": "thomeer",
        "btn-viz-shf": "shf",
        "btn-viz-winland": "winland",
        "btn-viz-pnm3d": "pnm3d",
        "btn-ms-pc-overlay": "ms_pc_overlay",
        "btn-ms-psd-compare": "ms_psd_compare",
        "btn-ms-cum-intrusion": "ms_cum_intrusion",
        "btn-ms-phik": "ms_phi_k",
        "btn-ms-jfunc": "ms_jfunc",
        "btn-ms-gpd": "ms_g_pd",
        "btn-ms-petrolog": "ms_petro_logs",
        "btn-ms-kprof": "ms_k_profile",
        "btn-ms-hfu": "ms_hfu",
        "btn-ms-shm": "ms_shm",
        "btn-ms-kpnm": "ms_k_pnm",
        "btn-ms-ci-log": "ms_ci_log",
        "btn-ms-reset": "intrusion",
    }
    mode = mapping.get(trig, "intrusion")
    return _set_mode(ui, mode)

# ---------------------------
# Parameters modal
# ---------------------------
@app.callback(
    Output("modal-params", "is_open", allow_duplicate=True),
    Output("p-sigma-hg", "value"),
    Output("p-theta-hg", "value"),
    Output("p-sigma-res", "value"),
    Output("p-theta-res", "value"),
    Output("p-rho-w", "value"),
    Output("p-rho-hc", "value"),
    Output("p-fwl-depth", "value"),
    Output("p-phi-ovr", "value"),
    Output("p-k-ovr", "value"),
    Output("p-swa", "value"),
    Output("p-swb", "value"),
    Input("btn-params", "n_clicks"),
    Input("btn-params-close", "n_clicks"),
    State("modal-params", "is_open"),
    State("store-params", "data"),
    prevent_initial_call=True,
)
def open_close_params(n_open, n_close, is_open, params):
    ctx = callback_context
    trig = ctx.triggered[0]["prop_id"].split(".")[0] if ctx and ctx.triggered else ""
    params = params or DEFAULT_PARAMS

    if trig == "btn-params":
        return (
            True,
            params.get("sigma_hg_air_npm"),
            params.get("theta_hg_air_deg"),
            params.get("sigma_res_npm"),
            params.get("theta_res_deg"),
            params.get("rho_w_kgm3"),
            params.get("rho_hc_kgm3"),
            params.get("fwl_depth_m"),
            params.get("phi_override_pct"),
            params.get("k_override_md"),
            params.get("swanson_a"),
            params.get("swanson_b"),
        )
    if trig == "btn-params-close":
        return (
            False,
            no_update, no_update, no_update, no_update, no_update, no_update, no_update, no_update, no_update, no_update, no_update
        )
    return is_open, no_update, no_update, no_update, no_update, no_update, no_update, no_update, no_update, no_update, no_update, no_update

@app.callback(
    Output("store-params", "data", allow_duplicate=True),
    Output("modal-params", "is_open", allow_duplicate=True),
    Output("store-status", "data", allow_duplicate=True),
    Input("btn-params-apply", "n_clicks"),
    State("store-params", "data"),
    State("p-sigma-hg", "value"),
    State("p-theta-hg", "value"),
    State("p-sigma-res", "value"),
    State("p-theta-res", "value"),
    State("p-rho-w", "value"),
    State("p-rho-hc", "value"),
    State("p-fwl-depth", "value"),
    State("p-phi-ovr", "value"),
    State("p-k-ovr", "value"),
    State("p-swa", "value"),
    State("p-swb", "value"),
    prevent_initial_call=True,
)
def apply_params(_, params, sigma_hg, theta_hg, sigma_res, theta_res, rho_w, rho_hc, fwl_depth, phi_ovr, k_ovr, swa, swb):
    params = params or DEFAULT_PARAMS
    params2 = dict(params)

    def set_if_not_none(key, val):
        if val is None or (isinstance(val, str) and val.strip() == ""):
            return
        params2[key] = val

    set_if_not_none("sigma_hg_air_npm", sigma_hg)
    set_if_not_none("theta_hg_air_deg", theta_hg)
    set_if_not_none("sigma_res_npm", sigma_res)
    set_if_not_none("theta_res_deg", theta_res)
    set_if_not_none("rho_w_kgm3", rho_w)
    set_if_not_none("rho_hc_kgm3", rho_hc)
    set_if_not_none("fwl_depth_m", fwl_depth)

    # Overrides can be cleared by entering blank -> treat as None
    params2["phi_override_pct"] = phi_ovr if phi_ovr is not None else None
    params2["k_override_md"] = k_ovr if k_ovr is not None else None

    set_if_not_none("swanson_a", swa)
    set_if_not_none("swanson_b", swb)

    return params2, False, "Parameters applied."



# ---------------------------
# Workflow tracker (progress + KPIs + log)
# ---------------------------
@app.callback(
    Output("store-log", "data", allow_duplicate=True),
    Input("store-status", "data"),
    State("store-log", "data"),
    prevent_initial_call=True,
)
def append_status_to_log(status, log):
    # Append the latest status message to a session log (keeps last ~200 entries).
    if status is None:
        return no_update
    msg = str(status).strip()
    if not msg:
        return no_update

    log = log or []
    ts = _dt.datetime.now().strftime("%H:%M:%S")
    entry = f"[{ts}] {msg}"

    # Avoid consecutive duplicates (common when multiple callbacks fire)
    if log and str(log[-1]).endswith(msg):
        return no_update

    log.append(entry)
    if len(log) > 200:
        log = log[-200:]
    return log


@app.callback(
    Output("wf-progress", "value"),
    Output("wf-progress", "label"),
    Output("wf-steps", "children"),
    Output("kpi-phi", "value"),
    Output("kpi-k", "value"),
    Output("kpi-r35", "value"),
    Output("kpi-kmethod", "value"),
    Output("kpi-rrt", "value"),
    Output("wf-log", "value"),
    Input("store-library", "data"),
    Input("store-current-id", "data"),
    Input("store-params", "data"),
    Input("store-log", "data"),
)
def update_workflow_panel(library, current_id, params, log):
    library = library or []
    params = params or DEFAULT_PARAMS
    log = log or []

    sample = _lib_get(library, current_id) if current_id else None
    if not sample:
        steps = [
            html.Div([html.Span(className="wf-box wf-off"), html.Span("Data Loaded")], className="wf-step"),
            html.Div([html.Span(className="wf-box wf-off"), html.Span("QC Check")], className="wf-step"),
            html.Div([html.Span(className="wf-box wf-off"), html.Span("Closure Corrected")], className="wf-step"),
            html.Div([html.Span(className="wf-box wf-off"), html.Span("Thomeer Fit")], className="wf-step"),
            html.Div([html.Span(className="wf-box wf-off"), html.Span("Petro QC")], className="wf-step"),
        ]
        return 0, "0/5", steps, "", "", "", "", "", "\n".join(log[-60:])

    res = sample.get("results", {}) or {}
    meta = sample.get("meta", {}) or {}

    # Step flags (current sample)
    data_loaded = bool(sample.get("data")) and len(sample.get("data") or []) > 0
    qc_done = bool(res.get("qaqc_done"))
    closure_done = bool(res.get("conf_applied"))
    thomeer_done = bool(res.get("thomeer_done")) or bool(res.get("thomeer_mode"))
    petro_done = bool(res.get("petro_qc_done"))

    done_count = int(data_loaded) + int(qc_done) + int(closure_done) + int(thomeer_done) + int(petro_done)
    progress = int(round(done_count / 5.0 * 100.0))

    def _step(label: str, ok: bool):
        return html.Div(
            [html.Span(className=f"wf-box {'wf-on' if ok else 'wf-off'}"), html.Span(label)],
            className="wf-step",
        )

    steps = [
        _step("Data Loaded", data_loaded),
        _step("QC Check", qc_done),
        _step("Closure Corrected", closure_done),
        _step("Thomeer Fit", thomeer_done),
        _step("Petro QC", petro_done),
    ]

    # KPIs (prefer overrides)
    phi = params.get("phi_override_pct")
    if phi is None:
        phi = meta.get("porosity_pct")

    # r35 (µm): computed from MICP at HgSat=0.35 (if range allows)
    r35 = res.get("r35_um")

    # Permeability + method (chosen)
    k_override = params.get("k_override_md")
    if k_override is not None:
        k = k_override
        kmethod = "Override"
    else:
        k_core = meta.get("permeability_md")
        if k_core is not None:
            k = k_core
            kmethod = "Core"
        else:
            k_swa = res.get("k_swanson_md")
            k_win = res.get("k_winland_md")
            if k_swa is not None:
                k = k_swa
                kmethod = "Swanson"
            elif k_win is not None:
                k = k_win
                kmethod = "Winland (Macro)" if (res.get("winland_mode") == "macro-normalized") else "Winland"
            else:
                k = None
                kmethod = ""

    rrt = res.get("rock_type") or ""


    def _fmt_num(x, nd=2):
        try:
            xf = float(x)
            if not np.isfinite(xf):
                return ""
            return f"{xf:.{nd}f}"
        except Exception:
            return ""

    phi_txt = _fmt_num(phi, 2)
    k_txt = _fmt_num(k, 3)
    r35_txt = _fmt_num(r35, 4)
    kmethod_txt = str(kmethod) if kmethod is not None else ""
    rrt_txt = str(rrt) if rrt is not None else ""

    log_txt = "\n".join(log[-60:])

    return progress, f"{done_count}/5", steps, phi_txt, k_txt, r35_txt, kmethod_txt, rrt_txt, log_txt

# ---------------------------
# Well name update
# ---------------------------
@app.callback(
    Output("store-library", "data", allow_duplicate=True),
    Output("store-status", "data", allow_duplicate=True),
    Input("inp-well", "value"),
    State("store-library", "data"),
    prevent_initial_call=True,
)
def set_well_name(well, library):
    if not library:
        return no_update, no_update
    if not well or not str(well).strip():
        return no_update, no_update
    well = str(well).strip()
    library2 = []
    for s in library:
        s2 = dict(s)
        s2["well"] = well
        library2.append(s2)
    return library2, f"Well set to: {well}"



# ---------------------------
# Petrophysical QAQC UI + Sample decision (Accept / Discard)
# ---------------------------

@app.callback(
    Output("petro-qc-summary", "children"),
    Output("badge-decision", "children"),
    Output("badge-decision", "color"),
    Output("badge-petro", "children"),
    Output("badge-petro", "color"),
    Output("badge-recommend", "children"),
    Output("badge-recommend", "color"),
    Input("store-library", "data"),
    Input("store-current-id", "data"),
    State("store-params", "data"),
)
def update_petro_qc_panel(library, current_id, params):
    library = library or []
    sample = _lib_get(library, current_id)
    if not sample:
        return (
            html.Div("No sample selected.", className="small-muted"),
            "PENDING",
            "secondary",
            "Petro QC: —",
            "dark",
            "Recommend: —",
            "dark",
        )

    params = params or DEFAULT_PARAMS
    meta = sample.get("meta", {}) or {}
    res = sample.get("results", {}) or {}

    decision = (res.get("sample_decision") or "PENDING").upper()
    dec_color = {"ACCEPTED": "success", "DISCARDED": "danger", "PENDING": "secondary"}.get(decision, "secondary")

    # Use stored Petro QC if present; otherwise compute on the fly (for older projects)
    qc_done = bool(res.get("petro_qc_done"))
    qc_grade = (res.get("petro_qc_grade") or "").upper()
    qc_rec = (res.get("petro_qc_recommendation") or "").upper()
    qc_issues = res.get("petro_qc_issues") or []

    if (not qc_done) and (not qc_grade):
        try:
            qc = petrophysical_qaqc(meta, params, df=_ensure_schema(pd.DataFrame(sample.get('data', []))), res=res)
            qc_done = bool(qc.get("petro_qc_done"))
            qc_grade = (qc.get("petro_qc_grade") or "").upper()
            qc_rec = (qc.get("petro_qc_recommendation") or "").upper()
            qc_issues = qc.get("petro_qc_issues") or []
        except Exception:
            qc_done = False
            qc_grade = ""
            qc_rec = ""
            qc_issues = []

    petro_text = f"Petro QC: {qc_grade}" if qc_grade else "Petro QC: —"
    petro_color = {"PASS": "success", "WARN": "warning", "FAIL": "danger"}.get(qc_grade, "dark")

    rec_text = f"Recommend: {qc_rec}" if qc_rec else "Recommend: —"
    rec_color = {"ACCEPT": "success", "REVIEW": "warning", "DISCARD": "danger"}.get(qc_rec, "dark")

    # Summary content
    children = []

    # Quick meta line (helps validate the "why")
    kv = []

    def _fmt(v, unit=""):
        if v is None:
            return None
        try:
            if isinstance(v, (int, float)) and not isinstance(v, bool) and not pd.isna(v):
                return f"{v:g}{unit}"
        except Exception:
            pass
        s = str(v).strip()
        return f"{s}{unit}" if s else None

    phi_used = params.get("phi_override_pct")
    if phi_used is None:
        phi_used = meta.get("porosity_pct")

    for label, val, unit in [
        ("Porosity", phi_used, "%"),
        ("Bulk ρ", meta.get("bulk_density_g_ml"), " g/mL"),
        ("Skeletal ρ", meta.get("skeletal_density_g_ml"), " g/mL"),
        ("Stem used", meta.get("stem_volume_used_pct"), "%"),
        ("Tortuosity", meta.get("tortuosity"), ""),
    ]:
        t = _fmt(val, unit)
        if t:
            kv.append(html.Span(f"{label}: {t}", className="small-muted", style={"marginRight": "0.75rem"}))

    if kv:
        children.append(html.Div(kv, style={"marginBottom": "0.35rem"}))

    if (not qc_done) and (not qc_grade) and (not qc_issues):
        children.append(html.Div("Petrophysical QAQC not available (no metadata fields found).", className="small-muted"))
    elif not qc_issues:
        children.append(html.Div("No petrophysical issues detected.", className="small-muted"))
    else:
        for iss in qc_issues:
            lvl = (iss.get("level") or "").upper()
            code = iss.get("code") or ""
            msg = iss.get("message") or ""
            lvl_color = {"FAIL": "danger", "WARN": "warning"}.get(lvl, "secondary")
            children.append(
                html.Div(
                    [
                        dbc.Badge(lvl, color=lvl_color, className="me-2"),
                        html.Span(code, className="petro-qc-code"),
                        html.Span(msg, className="petro-qc-msg"),
                    ],
                    className="petro-qc-item",
                )
            )

    return children, decision, dec_color, petro_text, petro_color, rec_text, rec_color




@app.callback(
    Output("alert-bimodal-hint-body", "children"),
    Output("alert-bimodal-hint", "is_open"),
    Output("btn-bimodal-hint-run", "style"),
    Input("store-library", "data"),
    Input("store-current-id", "data"),
    State("store-params", "data"),
)
def update_bimodal_hint_alert(library, current_id, params):
    """UI hint under the Modeling card suggesting a bi-modal Thomeer fit."""
    library = library or []
    params = params or DEFAULT_PARAMS

    s = _lib_get(library, current_id) if current_id else None
    if not s:
        return "", False, {"display": "none", "color": "#FFFFFF"}

    res = s.get("results", {}) or {}
    df = pd.DataFrame(s.get("data", []))
    df = _ensure_schema(df)

    try:
        flags = compute_bimodal_flags(df, params, res)
    except Exception:
        return "", False, {"display": "none", "color": "#FFFFFF"}

    if flags.get("bimodal_hint"):
        conf = (flags.get("bimodal_confidence") or "MEDIUM").upper()
        badge_color = "danger" if conf == "HIGH" else ("warning" if conf == "MEDIUM" else "secondary")

        reasons = flags.get("bimodal_reasons") or []
        r2_val = flags.get("bimodal_r2_value")
        r2_thr = flags.get("bimodal_r2_threshold")
        peak_sep = flags.get("bimodal_peak_sep_log10r")

        details = []
        if r2_val is not None and r2_thr is not None:
            try:
                details.append(f"R²={float(r2_val):.3f} (thr={float(r2_thr):.3f})")
            except Exception:
                pass
        if peak_sep is not None:
            try:
                details.append(f"Δlog10(r)≈{float(peak_sep):.2f}")
            except Exception:
                pass

        body = html.Div(
            [
                html.Div(
                    [
                        html.Span("Possible bi-modal pore system", style={"fontWeight": "600"}),
                        dbc.Badge(conf, color=badge_color, className="ms-2"),
                    ],
                    style={"display": "flex", "alignItems": "center", "gap": "6px"},
                ),
                html.Ul([html.Li(r) for r in reasons], style={"margin": "6px 0 0 18px"}) if reasons else html.Div(),
                html.Div(" • ".join(details), className="small-muted", style={"marginTop": "4px"}) if details else html.Div(),
                html.Div(
                    "Suggestion: Run Fit Thomeer (Bimodal) and compare R² + macro_frac. If unimodal is OK, keep unimodal.",
                    className="small-muted",
                    style={"marginTop": "4px"},
                ),
            ]
        )

        return body, True, {"display": "inline-block", "color": "#FFFFFF"}

    return "", False, {"display": "none", "color": "#FFFFFF"}


@app.callback(
    Output("modal-discard", "is_open"),
    Output("discard-error", "children"),
    Output("txt-discard-reason", "value"),
    Output("store-library", "data", allow_duplicate=True),
    Output("store-status", "data", allow_duplicate=True),
    Input("btn-discard-sample", "n_clicks"),
    Input("btn-discard-cancel", "n_clicks"),
    Input("btn-discard-confirm", "n_clicks"),
    Input("btn-accept-sample", "n_clicks"),
    State("modal-discard", "is_open"),
    State("txt-discard-reason", "value"),
    State("store-library", "data"),
    State("store-current-id", "data"),
    prevent_initial_call=True,
)
def handle_sample_decision(
    n_discard, n_cancel, n_confirm, n_accept, is_open, reason, library, current_id
):
    library = library or []
    ctx = callback_context
    if not ctx.triggered:
        raise PreventUpdate

    trig = ctx.triggered[0]["prop_id"].split(".")[0]

    if trig == "btn-discard-sample":
        if not _lib_get(library, current_id):
            return True, "No sample selected.", reason or "", no_update, no_update
        return True, "", reason or "", no_update, "Discard sample: please provide a reason."

    if trig == "btn-discard-cancel":
        return False, "", "", no_update, "Discard cancelled."

    if trig == "btn-accept-sample":
        def upd(s):
            res = dict(s.get("results", {}) or {})
            res["sample_decision"] = "ACCEPTED"
            res["discard_reason"] = ""
            res["decision_at"] = _now_iso()
            return {**s, "results": res}

        library2, s2 = _update_current_sample(library, current_id, upd)
        if not s2:
            return False, "", "", no_update, "No sample selected."
        return False, "", "", library2, f"Sample accepted: {s2.get('sample_id', current_id)}"

    if trig == "btn-discard-confirm":
        if not _lib_get(library, current_id):
            return True, "No sample selected.", reason or "", no_update, no_update

        if not (reason and str(reason).strip()):
            return True, "Please write a discard reason (required).", reason or "", no_update, no_update

        def upd(s):
            res = dict(s.get("results", {}) or {})
            res["sample_decision"] = "DISCARDED"
            res["discard_reason"] = str(reason).strip()
            res["decision_at"] = _now_iso()
            return {**s, "results": res}

        library2, s2 = _update_current_sample(library, current_id, upd)
        sid = s2.get("sample_id", current_id) if s2 else (current_id or "")
        return False, "", "", library2, f"Sample discarded: {sid}"

    return is_open, "", reason or "", no_update, no_update

# -----------------------------------------------------------------------------
# UI FIXES (v1.6.4): restore sample loading + selector sync, and Project Storage modal
# -----------------------------------------------------------------------------

def _unique_sample_id(base: str, existing: set[str]) -> str:
    base = (base or "sample").strip()
    if base not in existing:
        return base
    i = 2
    while f"{base}_{i}" in existing:
        i += 1
    return f"{base}_{i}"


def _ask_directory_dialog(initial: str | None = None, title: str = "Select folder") -> str | None:
    """Open a native folder picker (local desktop only). Returns selected path or None."""
    try:
        import tkinter as tk
        from tkinter import filedialog

        root = tk.Tk()
        root.withdraw()
        try:
            root.wm_attributes("-topmost", 1)
        except Exception:
            pass
        folder = filedialog.askdirectory(initialdir=initial or os.getcwd(), title=title)
        try:
            root.destroy()
        except Exception:
            pass
        folder = folder.strip() if isinstance(folder, str) else ""
        return folder or None
    except Exception:
        return None


@app.callback(
    Output("modal-storage", "is_open"),
    Input("btn-storage", "n_clicks"),
    Input("btn-storage-close", "n_clicks"),
    State("modal-storage", "is_open"),
    prevent_initial_call=True,
)
def toggle_project_storage(n_open, n_close, is_open):
    trig = callback_context.triggered_id
    if trig == "btn-storage":
        return True
    if trig == "btn-storage-close":
        return False
    return is_open


@app.callback(
    # allow_duplicate because inp-workspace-dir.value is also updated by the
    # store-workspace sync callback (needed on page load).
    Output("inp-workspace-dir", "value", allow_duplicate=True),
    Output("store-status", "data", allow_duplicate=True),
    Input("btn-browse-workspace", "n_clicks"),
    State("inp-workspace-dir", "value"),
    prevent_initial_call=True,
)
def browse_workspace_folder(n, current_val):
    if not n:
        raise PreventUpdate
    folder = _ask_directory_dialog(current_val, title="Select Workspace Folder")
    if not folder:
        return no_update, no_update
    return folder, f"Workspace folder set to: {folder}"


@app.callback(
    Output("inp-import-dir", "value"),
    Output("store-status", "data", allow_duplicate=True),
    Input("btn-browse-import-dir", "n_clicks"),
    State("inp-import-dir", "value"),
    prevent_initial_call=True,
)
def browse_import_folder(n, current_val):
    if not n:
        raise PreventUpdate
    folder = _ask_directory_dialog(current_val, title="Select Folder to Import")
    if not folder:
        return no_update, no_update
    return folder, f"Import folder selected: {folder}"


@app.callback(
    Output("store-library", "data", allow_duplicate=True),
    Output("store-current-id", "data", allow_duplicate=True),
    Output("inp-well", "value", allow_duplicate=True),
    Output("store-log", "data", allow_duplicate=True),
    Output("store-status", "data", allow_duplicate=True),
    Input("btn-import-folder", "n_clicks"),
    State("inp-import-dir", "value"),
    State("store-library", "data"),
    State("store-current-id", "data"),
    State("inp-well", "value"),
    State("store-log", "data"),
    prevent_initial_call=True,
)
def import_folder(n, folder, library, current_id, well_value, log):
    if not n:
        raise PreventUpdate
    folder = (folder or "").strip()
    if not folder or not os.path.isdir(folder):
        return no_update, no_update, no_update, no_update, f"Invalid folder: {folder}"

    library = library or []
    existing_ids = {s.get("id") for s in library if isinstance(s, dict)}
    log = log or []

    # Gather files
    exts = ("*.xlsx", "*.xls", "*.csv", "*.txt")
    paths: list[str] = []
    for pat in exts:
        paths.extend(sorted(glob.glob(os.path.join(folder, pat))))
    if not paths:
        return no_update, no_update, no_update, no_update, f"No supported files found in: {folder}"

    imported = 0
    failed = 0
    new_well = well_value

    for p in paths:
        fn = os.path.basename(p)
        try:
            ext = os.path.splitext(fn)[1].lower()
            if ext in (".csv", ".txt"):
                df_raw = pd.read_csv(p)
                fmt_tag = "csv"
            else:
                with open(p, "rb") as f:
                    b = f.read()
                fmt_tag = ext
                df_raw, fmt_tag = _read_excel_bytes(b, ext)

            df, _parse_tag = parse_to_three_cols(df_raw, fmt_tag)
            meta = extract_meta_from_raw(df_raw, filename=fn)

            # Determine well
            if not new_well:
                new_well = meta.get("well") or meta.get("well_guess") or well_value

            # Determine sample id
            base_id = meta.get("sample_id") or meta.get("well_guess") or os.path.splitext(fn)[0]
            sid = _unique_sample_id(str(base_id), existing_ids)
            existing_ids.add(sid)

            sample = {
                "id": sid,
                "sample_id": sid,
                "filename": fn,
                "source_path": p,
                "well": new_well,
                "meta": meta,
                "data": df.to_dict("records"),
                "data_raw": df.to_dict("records"),
                "results": {},
            }
            library.append(sample)
            imported += 1
            log.append(f"[IMPORT] {fn} -> {sid} ({len(df)} rows)")
        except Exception as e:
            failed += 1
            log.append(f"[IMPORT][FAIL] {fn}: {e}")

    # Set current sample if needed
    if current_id is None and library:
        current_id = library[-1].get("id")

    status = f"Imported {imported} file(s) from folder. Failures: {failed}."
    return library, current_id, new_well, log, status


@app.callback(
    Output("store-library", "data", allow_duplicate=True),
    Output("store-current-id", "data", allow_duplicate=True),
    Output("inp-well", "value", allow_duplicate=True),
    Output("store-status", "data", allow_duplicate=True),
    Input("upload-data", "contents"),
    State("upload-data", "filename"),
    State("store-library", "data"),
    State("store-current-id", "data"),
    State("inp-well", "value"),
    prevent_initial_call=True,
)
def on_import_upload(list_of_contents, list_of_names, library, current_id, well_value):
    if not list_of_contents:
        raise PreventUpdate

    library = library or []
    existing_ids = {s.get("id") for s in library if isinstance(s, dict)}

    imported = 0
    failed = 0
    new_well = well_value

    for contents, fn in zip(list_of_contents, list_of_names or []):
        try:
            ext = os.path.splitext(fn)[1].lower().lstrip(".")
            df_raw, fmt_tag = _read_uploaded(contents, fn)
            df, _parse_tag = parse_to_three_cols(df_raw, fmt_tag)
            meta = extract_meta_from_raw(df_raw, filename=fn)

            if not new_well:
                new_well = meta.get("well") or meta.get("well_guess") or well_value

            base_id = meta.get("sample_id") or meta.get("well_guess") or os.path.splitext(fn)[0]
            sid = _unique_sample_id(str(base_id), existing_ids)
            existing_ids.add(sid)

            sample = {
                "id": sid,
                "sample_id": sid,
                "filename": fn,
                "well": new_well,
                "meta": meta,
                "data": df.to_dict("records"),
                "data_raw": df.to_dict("records"),
                "results": {},
            }
            library.append(sample)
            imported += 1
        except Exception:
            failed += 1

    # Default current sample
    if current_id is None and library:
        current_id = library[-1].get("id")

    msg = f"Imported {imported} file(s). Failures: {failed}."
    return library, current_id, new_well, msg




# -----------------------------------------------------------------------------
# External N4 logs module callbacks
# -----------------------------------------------------------------------------

@app.callback(
    Output("store-logn4", "data"),
    Output("logn4-status", "children"),
    Output("dd-logn4-sheet", "options"),
    Output("dd-logn4-sheet", "value"),
    Input("upload-logn4", "contents"),
    Input("btn-logn4-clear", "n_clicks"),
    State("upload-logn4", "filename"),
    State("store-logn4", "data"),
    prevent_initial_call=True,
)
def on_logn4_import(contents_list, n_clear, filenames, store):
    trig = callback_context.triggered_id

    if trig == "btn-logn4-clear":
        return None, "N4 log data cleared.", [], None

    if not contents_list:
        raise PreventUpdate

    # Dash returns a single string when multiple=False; normalize to lists.
    if not isinstance(contents_list, list):
        contents_list = [contents_list]
    if not isinstance(filenames, list):
        filenames = [filenames] if filenames else []

    # Keep lengths aligned
    if len(filenames) != len(contents_list):
        if len(filenames) == 1 and len(contents_list) > 1:
            filenames = filenames * len(contents_list)
        else:
            filenames = [(filenames[i] if i < len(filenames) else f"upload_{i+1}.xlsx") for i in range(len(contents_list))]

    store_out = store or {}

    for contents, fn in zip(contents_list, filenames):
        if not contents or "," not in contents:
            continue
        try:
            decoded = base64.b64decode(contents.split(",")[1])
        except Exception:
            continue

        sheets = _read_logn4_from_excel_bytes(decoded, filename=fn or "")
        store_out = _merge_logn4_store(store_out, sheets, filename=fn or "")

    status = _summarize_logn4_store(store_out)
    sheet_keys = list((store_out.get("sheets") or {}).keys())
    options = [{"label": k, "value": k} for k in sheet_keys]
    value = sheet_keys[0] if sheet_keys else None
    return store_out, status, options, value


@app.callback(
    Output("dd-logn4-logs", "options"),
    Output("dd-logn4-logs", "value"),
    Input("dd-logn4-sheet", "value"),
    Input("store-logn4", "data"),
)
def on_logn4_sheet_selected(sheet, store):
    if not store or not sheet:
        return [], []

    payload = (store.get("sheets") or {}).get(sheet)
    if not payload:
        return [], []

    cols = payload.get("columns", []) or []
    records = payload.get("records", []) or []

    if records:
        df = pd.DataFrame(records)
    else:
        df = pd.DataFrame(columns=cols)

    log_cols = [c for c in cols if c != "Depth"]

    # If we have data, keep only logs with at least one non-null value.
    if len(df) > 0:
        log_cols = [c for c in log_cols if c in df.columns and df[c].notna().any()]

    options = [{"label": c, "value": c} for c in log_cols]

    preferred = [c for c in ["PorAmb", "PorOB", "PermAmb", "PermOB"] if c in log_cols]
    if not preferred:
        preferred = log_cols[:2] if len(log_cols) >= 2 else log_cols

    return options, preferred


@app.callback(
    Output("graph-logn4", "figure"),
    Input("dd-logn4-sheet", "value"),
    Input("dd-logn4-logs", "value"),
    Input("store-logn4", "data"),
)
def update_logn4_graph(sheet, selected_logs, store):
    fig = go.Figure()
    apply_plot_theme(fig, theme="dark")

    if not store or not sheet:
        fig.add_annotation(
            text="No external core log data loaded.",
            x=0.5,
            y=0.5,
            xref="paper",
            yref="paper",
            showarrow=False,
        )
        fig.update_layout(height=320)
        return fig

    payload = (store.get("sheets") or {}).get(sheet)
    if not payload:
        fig.add_annotation(
            text="Selected sheet not available.",
            x=0.5,
            y=0.5,
            xref="paper",
            yref="paper",
            showarrow=False,
        )
        fig.update_layout(height=320)
        return fig

    cols = payload.get("columns", []) or []
    records = payload.get("records", []) or []

    df = pd.DataFrame(records) if records else pd.DataFrame(columns=cols)

    if "Depth" not in df.columns or len(df) == 0:
        fig.add_annotation(
            text="No log records found in this sheet.",
            x=0.5,
            y=0.5,
            xref="paper",
            yref="paper",
            showarrow=False,
        )
        fig.update_layout(height=320)
        return fig

    if not selected_logs:
        fig.add_annotation(
            text="Select one or more logs to plot.",
            x=0.5,
            y=0.5,
            xref="paper",
            yref="paper",
            showarrow=False,
        )
        fig.update_layout(height=320)
        return fig

    for col in selected_logs:
        if col not in df.columns:
            continue
        fig.add_trace(go.Scatter(x=df[col], y=df["Depth"], mode="lines+markers", name=col))

    fig.update_yaxes(autorange="reversed", title="Depth")
    fig.update_xaxes(title="Value")
    fig.update_layout(height=320, margin=dict(l=40, r=10, t=20, b=40), legend=dict(orientation="h"))

    return fig


@app.callback(
    Output("download-logn4-csv", "data"),
    Input("btn-logn4-export", "n_clicks"),
    State("dd-logn4-sheet", "value"),
    State("store-logn4", "data"),
    prevent_initial_call=True,
)
def export_logn4_csv(n_clicks, sheet, store):
    if not n_clicks or not store or not sheet:
        raise PreventUpdate

    payload = (store.get("sheets") or {}).get(sheet)
    if not payload:
        raise PreventUpdate

    cols = payload.get("columns", []) or []
    records = payload.get("records", []) or []
    df = pd.DataFrame(records) if records else pd.DataFrame(columns=cols)

    safe_sheet = re.sub(r"[^A-Za-z0-9_\-]+", "_", str(sheet))
    filename = f"{safe_sheet}_N4_logs.csv"
    return dcc.send_data_frame(df.to_csv, filename, index=False)


@app.callback(
    Output("sel-sample", "options"),
    Output("sel-sample", "value"),
    Output("dashboard-title", "children"),
    Output("loaded-files", "children"),
    Input("store-library", "data"),
    Input("store-current-id", "data"),
)

def on_library_or_current_id(library, current_id):
    """Keep the Sample dropdown + dashboard header in sync with the loaded library."""
    library = _coerce_library_list(library)

    # Build stable list of sample IDs
    ids = []
    for s in library:
        if not isinstance(s, dict):
            continue
        sid = s.get("sample_id") or (s.get("meta", {}) or {}).get("sample_id") or s.get("id") or s.get("filename")
        sid = str(sid).strip() if sid is not None else ""
        if sid:
            ids.append(sid)

    # Deduplicate while preserving order
    seen = set()
    ids_unique = []
    for sid in ids:
        if sid in seen:
            continue
        seen.add(sid)
        ids_unique.append(sid)

    options = [{"label": sid, "value": sid} for sid in ids_unique]

    # Validate/choose current value
    value = current_id if (current_id in ids_unique) else (ids_unique[0] if ids_unique else None)

    # Title and loaded files summary
    title = f"Dashboard — {value}" if value else "Dashboard"
    if ids_unique:
        preview = ", ".join(ids_unique[:12])
        if len(ids_unique) > 12:
            preview += ", …"
        loaded = f"Loaded {len(ids_unique)} sample(s): {preview}"
    else:
        loaded = "No samples loaded."

    return options, value, title, loaded



# === Core Validation panel ===
@app.callback(
    Output("core-validation-table", "children"),
    Input("store-library", "data"),
    Input("store-current-id", "data"),
    Input("store-logn4", "data"),
    Input("dd-logn4-sheet", "value"),
)
def on_core_validation_table(library, current_id, logn4_store, selected_sheet):
    library = _coerce_library_list(library)
    current_id = current_id or (library[0].get("id") if library else None)

    if not current_id:
        return html.Div("No sample selected.", className="text-muted")

    sample = next((s for s in library if s.get("id") == current_id), None)
    if not sample:
        return html.Div("No sample selected.", className="text-muted")

    res = sample.get("results") or {}

    # Predicted permeability (mD)
    k_pnm = _safe_float(res.get("k_pnm_md"))
    k_win = _safe_float(res.get("k_winland_md_total") or res.get("k_winland_md"))

    # Thomeer bimodal separation (Pd2/Pd1)
    pd1 = _safe_float(res.get("thomeer_pd1_psia"))
    pd2 = _safe_float(res.get("thomeer_pd2_psia"))
    pd_ratio = _safe_div(pd2, pd1)

    # External N4 values (ambient vs overburden)
    depth_target = _extract_depth_from_sample_id(current_id)

    por_amb = por_ob = None
    perm_amb = perm_ob = None
    depth_used = depth_delta = None
    sheet_used = None

    if isinstance(logn4_store, dict) and (logn4_store.get("sheets") or {}):
        sheets = list((logn4_store.get("sheets") or {}).keys())
        try_order = []
        if selected_sheet and selected_sheet in sheets:
            try_order.append(selected_sheet)
        for sn in sheets:
            if sn not in try_order:
                try_order.append(sn)

        for sn in try_order:
            rec, d_used, d_delta = _find_nearest_logn4_record(logn4_store, sn, depth_target)
            if not rec:
                continue

            # Column name fuzzing: the importer may prepend group labels.
            pa = _pick_record_value(rec, ["poramb", "phiamb"])
            po = _pick_record_value(rec, ["porob", "phiob"])
            ka = _pick_record_value(rec, ["permamb", "kamb"])
            ko = _pick_record_value(rec, ["permob", "kob"])

            if any(v is not None for v in (pa, po, ka, ko)):
                por_amb, por_ob = pa, po
                perm_amb, perm_ob = ka, ko
                depth_used, depth_delta = d_used, d_delta
                sheet_used = sn
                break

    stress_ratio = _safe_div(perm_ob, perm_amb)
    mismatch_pnm = _safe_div(k_pnm, perm_ob)
    mismatch_win = _safe_div(k_win, perm_ob)

    # Optional clay% (not yet imported in v1.8.21) — placeholder for future XRD module.
    clay_pct = None

    supported, score, reasons = _bimodal_supported_flag(pd_ratio=pd_ratio, stress_ratio=stress_ratio, clay_pct=clay_pct)

    badge = dbc.Badge(
        "YES" if supported else "NO",
        color="success" if supported else "secondary",
        className="ms-1",
    )

    rows = [
        ("PermAmb (mD)", _fmt_num(perm_amb)),
        ("PermOB (mD)", _fmt_num(perm_ob)),
        ("PermOB/PermAmb", _fmt_num(stress_ratio)),
        ("k_PNM (mD)", _fmt_num(k_pnm)),
        ("k_Winland (mD)", _fmt_num(k_win)),
        ("k_PNM/PermOB", _fmt_num(mismatch_pnm)),
        ("k_Winland/PermOB", _fmt_num(mismatch_win)),
        ("Pd2/Pd1", _fmt_num(pd_ratio)),
        ("Clay% (XRD)", _fmt_num(clay_pct)),
        ("Bimodal supported?", html.Span([" ", badge])),
    ]

    table = dbc.Table(
        [
            html.Thead(html.Tr([html.Th("Metric"), html.Th("Value")]), style={"borderBottom": "1px solid #444"}),
            html.Tbody([html.Tr([html.Td(k), html.Td(v)]) for k, v in rows]),
        ],
        bordered=True,
        hover=True,
        responsive=True,
        size="sm",
        style={"marginBottom": "6px"},
    )

    notes = []
    if sheet_used is None:
        notes.append("Import External Core Logs to compute PermAmb/PermOB (stress sensitivity).")
    else:
        if depth_used is not None:
            if depth_delta is None:
                notes.append(f"Using {sheet_used} at depth {depth_used:.2f}.")
            else:
                notes.append(f"Using {sheet_used} near depth {depth_used:.2f} (Δ={depth_delta:.2f}).")

    if not reasons:
        reasons_txt = "No supporting signals met the heuristic thresholds."
    else:
        reasons_txt = "Signals: " + ", ".join(reasons)

    notes.append(reasons_txt + f" (score={score}).")

    alert = dbc.Alert(
        " ".join(notes),
        color="info" if sheet_used else "secondary",
        style={"fontSize": "12px", "padding": "6px 10px", "marginBottom": 0},
    )

    return html.Div([table, alert])
def sync_sample_selector(library, current_id):
    library = library or []
    opts = []
    for s in library:
        sid = (s or {}).get("id") or (s or {}).get("sample_id")
        if sid:
            opts.append({"label": sid, "value": sid})
    # Choose a sensible value
    value = current_id
    if value is None and opts:
        value = opts[-1]["value"]
    title = f"Dashboard — {value}" if value else "Dashboard"
    files = [s.get("filename") for s in library if isinstance(s, dict) and s.get("filename")]
    loaded = "Loaded: " + ", ".join(files) if files else "Loaded: —"
    return opts, value, title, loaded

if __name__ == "__main__":
    # Disable reloader/debug by default (more stable for local file dialogs).
    debug = os.environ.get("DASH_DEBUG", "0").strip() in {"1", "true", "True"}
    port = int(os.environ.get("PORT", "8050"))
    app.run(debug=debug, host="127.0.0.1", port=port)