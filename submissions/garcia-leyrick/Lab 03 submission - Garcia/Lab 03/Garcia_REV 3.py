"""
REV3.py  (Rev. 3 -- interactive solver, catalog-driven)

Excel -> Full Catalog -> Interactive Web Solver
--------------------------------------------------
Earlier revisions resolved ONE unit system / material / member size at
Python start-up (from --units/--material/--member-size) and baked that
single choice into the output. This revision instead reads the FULL
contents of every Excel workbook -- every unit system, every material,
every W-shape, in both unit systems -- into one catalog, embeds that
catalog in the page, and lets the browser itself switch between them live:
changing the Units / Material / Member Size dropdowns re-derives the model
and re-solves immediately, no re-running Python. Excel is still the only
source of this data; the catalog is just "all of it" instead of "one row
of it".

The 3D viewer no longer depends on any CDN (three.js) -- earlier testing
showed that failing unpredictably on real machines (corporate networks,
file:// security restrictions, etc.). It's now a small dependency-free
Canvas2D wireframe renderer with mouse-orbit and wheel-zoom, written once
in this file, that only needs a browser -- no network access, ever.

Excel is the source of truth for:
  units/*.xlsx          -> unit systems and conversion factors
  materials/*.xlsx       -> RISA material library (both unit systems), incl. G
  member_size/*.xlsx     -> AISC v16.0 shapes database (both unit systems),
                            incl. Iy, J, Sx, Sy for every W-shape

Modelling assumptions, stated once, honestly
-----------------------------------------------
  * Euler-Bernoulli (no shear deformation) 3D frame elements -- the standard
    12x12 space-frame stiffness matrix, axial + torsion + biaxial bending.
  * Bending about a member's local y-axis uses the section's *weak* axis
    property (Iy / Sy); about local z uses the *strong* axis property
    (Ix / Sx). A stated modelling convention, not an AISC design check.
  * All 12 members are rigidly (moment) connected at both ends.
  * Linear-elastic, first-order analysis; the utilisation ratio shown is a
    simple elastic-superposition stress estimate against Fy, not a full
    AISC interaction-equation capacity check.
"""

import argparse
import base64
import glob
import json
import os
import re
import sys
import webbrowser
import zlib

import numpy as np

try:
    import openpyxl
except ImportError:           # standalone mode doesn't need it
    openpyxl = None

HERE = os.path.dirname(os.path.abspath(__file__))


# =============================================================================
# EMBEDDED CATALOG  -  a gzip+base64 snapshot of everything build_catalog()
# reads out of the Excel workbooks (both unit systems, every material, every
# W-shape). Used only when the units/ materials/ member_size/ folders are NOT
# beside this script, so REV3.py can run completely standalone. Regenerate it
# after editing the workbooks with:  python REV3.py --embed-catalog
# =============================================================================
EMBEDDED_CATALOG_B64 = (
    "eNrUvVuvXceVHfxXBD51AGun7he+GUa6LcFuBK0GLCAdELR07BBNkf2RVNpCkP/+7VVVc8xL1T7sOLCEGH5YOrW411i1qmrO"
    "Oebtf734+NPHT08/fHzx8n+9+OHp04c3311X717/8PTi5YtvPr1+9/3rD99/8fs58qsXP7z+9PThzeu34/5fx/LFP3y4xXL9"
    "x9vXf3x6e/83/Mdfvfjufvef33/46f7n377/9MU/vX/79un7+9//y4uXvvee2s396sU/vHhZS3fpuv77+63h/udwv/zxxcvY"
    "+63/6sX3T+8+vvn003XfrddfvXh69eO7N5/uP/r7//r6/mt/1v/5p5/Mf/+o/3v9Gv3xX//xP//wLz8698f44n//6sWvcw0X"
    "/OzUO/Ff/8qXiind6i/2Ur0H9TrXf/8/+SLZues7JP06/Ne/dsm1fiu/+EuV40v91fso+vqL7iP3Kkazicaf1Ov85v3b77/4"
    "+/cffqD3CS7Gnul9WggNHynUW17vk9t1+XO+T3H11W/8q5zVK/Ff/9q3irXTV0otXKvwZ3yr37x/9110zv3jH+Rbib+at3r3"
    "3YenT0/rlXwtbr1RT/NyfCZH58O7H9++Fa8Twq22v/3r5OPr5M++zv2WtF7Hu+B52aWb/wXfJx0/T/r858k+8Pv04PE+lY66"
    "X+b73JH/7rjcfvfs+/gU7/thvk8JOX12uV2n38+z3H53XG6feZ2c7x9ivU7N7rOr7ed5nXT8OunzX6cUF+l4C75/drH9PK9z"
    "Afzi968/vbVvxAPqpX7/+uP7d/er8U49OpKrsYZIr+TdLR5faf7n3/KN3r7+aX8b/PHxm9zR9/Umd02n/PJvct/u8cvf+iRf"
    "BH9T7/Hrtz/+8Obdjz/MFym9RNo1oXiPU8Dzu/g4jjo+1urN/21XWnHFf/nPSnujP/017xKSp3cJJfz87xK//Ods3mX86f/2"
    "u2R/ndQ/+7uU/V3+uu/ia7il9V1cubWf9V2yy+HL30a1Y/C3Z96muhgevE3HKovperGff8d88Xf//vT2+6fv/9Nh6/DYX/Wt"
    "7tJofStfsrIV/vZv9803Xr7Q9Z/qHb759PrNu7dPHz+SBRe9I026phjyLbVE2s19oaX1nfJdaHon99BdPrn+N3yX+8t8fPru"
    "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"
    "jv5GmO9XkZ4RgS8Wehq/R+z0HinRaMJonivzflUyPaPOU+B+1fDclulvfRh8f0jXyuh0NY/f62pu7uuq0+haI+OK/u1aLdcV"
    "fiWURFe14qrRVcdop9EYEq7wt0hPi3NOr6uKq85X9C8S/kV2hD53uirzGEiJNJjrav3y/arT39YJcl3N/Xe/SnOPLcWRtFlS"
    "rskYIluRzDQyZMm4hmmwLBaiAIiZILaEGByilYjqIvqNOEHiKIk4JS6XuOVJdv/3X7344fWnpw9vXr8Vem10Ln7523GwZZfD"
    "l78dwrC44r/858JXX/zdvz+9/f7p+/80/xS//OeMq6HQx/LFP3y4zTjhuyZy/cfM9Kf/mCPVvZo+vlzD9ecZ6lZcffUb/2qW"
    "k+jji/7mro9/8fvXn95e1+/ffXc1J/rdH8R//CP+I8uRLEaS/DdJ/psPT5+e6Ne/+cZfc/P26d2fP/2PVz++e/Ppbhv86dML"
    "tgLW3968u/726cPTx4/0p3/9+Ob62/u3//Ppwyv9C+Pup+///LT+PrN/2n1F9rtEcrXVi0979/77p1ef3r9SPzEIcferF8ss"
    "efWn9x++e8Ij3/zbCx764f0PT+8+ibF/+dG5P9bx8PnP6M6P371++zSM1Ps6mP/qNPT2/evv1wPfvPu3Hz/tg+sf76PfP737"
    "+ObTT/w6w052Lt/f1UX+P6/Dj/cbf7w/6g79n7765tevfo+//+7NHz+8/vDTq69++Lfxl9tf3n78ywtpmj39fz+++Z+XwfgH"
    "f1c67/pWHSYd79F59VKM/2pd59TWvbyzR7duMU73pivt154HlxNCjOPey19uTpHr6qUYp3vj/Ls+e64C/WIc9yb8Lp9YkX83"
    "LpX3Ptd3VbKMe/mcm1cvxTjd21pc9/LpeF29FON0b11zps5Ux/fWNWfXdUlh3csn8XX1UozTvbn7dS+f39fVSzGOe6/YFXv+"
    "18D3kqV8XcdO97LUKOJeMriv74J5kPIn4d4k5iElupelVhB4k/gWydM8SPkn7vU8D7HSPAj5mTrujZXnIQbcy/LX8zeOge8N"
    "ndYDy2/f+XdD5/WQG+a3S/n6Uozj3TruZfksvnHq4t5WML9VyPeXYhz3xob5TUJTeCnGeX4z5rcJ/eKlGMecdXw31koK/27s"
    "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"
)
# --- end embedded catalog ---


# =============================================================================
# INLINE ASSETS  -  the page template, the FE solver core, and the
# dependency-free Canvas2D viewer, embedded so REV3.py is a single
# self-contained script. Only the Excel data folders live outside it,
# because those are what this exercise keeps externally editable.
# =============================================================================
PAGE_TEMPLATE = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>REV3 Structural Solver</title>
<style>
  :root {
    --ink:#1b1f24; --sub:#5b6672; --line:#dfe3e8; --panel:#f7f9fb;
    --accent:#1f6feb; --accent-dark:#123a75; --good:#2d6a4f; --warn:#b45309; --bad:#b42318;
  }
  * { box-sizing:border-box; }
  html, body { height:100%; margin:0; }
  body {
    font-family:-apple-system,"Segoe UI",Helvetica,Arial,sans-serif; color:var(--ink);
    background:#fff; display:flex; flex-direction:column; font-size:13px; overflow:hidden;
  }

  /* ---- ribbon ---- */
  .ribbon { border-bottom:1px solid var(--line); background:#fff; flex-shrink:0; }
  .ribbon-top {
    display:flex; align-items:center; gap:10px; padding:8px 16px 6px 16px; flex-wrap:wrap;
  }
  .ribbon-top h1 { font-size:1.02rem; margin:0; white-space:nowrap; }
  .ribbon-top .sep { width:1px; height:22px; background:var(--line); }
  .ctl { display:flex; align-items:center; gap:6px; }
  .ctl label { font-size:0.72rem; color:var(--sub); text-transform:uppercase; letter-spacing:.03em; }
  .ctl select {
    padding:5px 7px; border:1px solid var(--line); border-radius:6px; font-size:0.82rem;
    background:#fff; max-width:230px;
  }
  .ribbon-status {
    display:flex; align-items:center; gap:14px; padding:5px 16px; background:var(--panel);
    border-top:1px solid var(--line); font-size:0.76rem; color:var(--sub); flex-wrap:wrap;
  }
  .ribbon-status b { color:var(--ink); }
  .pill {
    padding:2px 9px; border-radius:99px; font-size:0.72rem; font-weight:700;
  }
  .pill-ok { background:#e7f5ec; color:var(--good); }
  .pill-warn { background:#fdf3e4; color:var(--warn); }
  .pill-bad { background:#fdecea; color:var(--bad); }
  .pill-idle { background:#eceff2; color:var(--sub); }

  /* ---- layout ---- */
  .layout { flex:1; display:flex; min-height:0; }
  .viewer-col { flex:1 1 60%; position:relative; border-right:1px solid var(--line); min-width:0; background:#fbfcfd; }
  #view3d { width:100%; height:100%; display:block; cursor:grab; }
  #view3d:active { cursor:grabbing; }
  .view-controls { position:absolute; top:10px; left:10px; display:flex; gap:6px; flex-wrap:wrap; }
  .view-controls button {
    font-size:0.76rem; padding:5px 10px; border:1px solid var(--line); background:#fff;
    border-radius:6px; cursor:pointer;
  }
  .view-controls button.active { background:var(--accent); color:#fff; border-color:var(--accent); }
  .view-hint {
    position:absolute; bottom:10px; left:10px; font-size:0.72rem; color:var(--sub);
    background:rgba(255,255,255,.9); border:1px solid var(--line); border-radius:7px; padding:6px 10px;
  }
  .selection-box {
    position:absolute; top:10px; right:10px; background:rgba(255,255,255,.95);
    border:1px solid var(--line); border-radius:8px; padding:8px 11px; font-size:0.76rem;
    max-width:260px; display:none;
  }
  .selection-box.show { display:block; }
  .selection-box h4 { margin:0 0 5px 0; font-size:0.8rem; }

  /* ---- side panel ---- */
  .side-col { flex:1 1 40%; display:flex; flex-direction:column; min-width:350px; max-width:580px; }
  .tabs { display:flex; border-bottom:1px solid var(--line); background:var(--panel); flex-shrink:0; }
  .tabs button {
    flex:1; padding:9px 6px; border:none; background:none; cursor:pointer; font-size:0.84rem;
    font-weight:600; color:var(--sub); border-bottom:2px solid transparent;
  }
  .tabs button.active { color:var(--accent); border-bottom-color:var(--accent); background:#fff; }
  .panel { flex:1; overflow-y:auto; padding:13px 15px; display:none; }
  .panel.active { display:block; }
  .card { border:1px solid var(--line); border-radius:8px; padding:11px 13px; margin-bottom:11px; }
  .card h3 { margin:0 0 8px 0; font-size:0.85rem; }
  .kv { display:grid; grid-template-columns:1fr auto; gap:3px 10px; font-size:0.78rem; }
  .kv .k { color:var(--sub); }
  .kv .v { font-weight:600; text-align:right; }
  .grid3 { display:grid; grid-template-columns:repeat(3,1fr); gap:7px; }
  .fld label { display:block; font-size:0.7rem; color:var(--sub); margin-bottom:2px; }
  .fld input, .fld select {
    width:100%; padding:5px 6px; border:1px solid var(--line); border-radius:5px; font-size:0.8rem;
  }
  button.primary {
    background:var(--accent); color:#fff; border:none; border-radius:6px; padding:9px 14px;
    font-size:0.85rem; font-weight:700; cursor:pointer; width:100%;
  }
  button.primary:hover { background:var(--accent-dark); }
  button.ghost {
    background:#fff; border:1px solid var(--line); border-radius:6px; padding:6px 10px;
    font-size:0.77rem; cursor:pointer; color:var(--sub);
  }
  button.ghost:hover { border-color:var(--accent); color:var(--accent); }
  table { width:100%; border-collapse:collapse; font-size:0.72rem; }
  th, td { text-align:right; padding:4px 5px; border-bottom:1px solid var(--line); white-space:nowrap; }
  th:first-child, td:first-child { text-align:left; }
  th { color:var(--sub); font-weight:600; background:var(--panel); position:sticky; top:0; }
  tr.clickable { cursor:pointer; }
  tr.clickable:hover { background:#f2f7ff; }
  tr.sel { background:#fff6e6; }
  .loadrow {
    display:flex; justify-content:space-between; align-items:center; gap:8px; padding:6px 9px;
    border:1px solid var(--line); border-radius:6px; margin-bottom:5px; font-size:0.75rem;
    background:var(--panel);
  }
  .loadrow button { border:none; background:none; color:var(--bad); cursor:pointer; font-size:0.95rem; }
  .empty { color:var(--sub); font-size:0.77rem; font-style:italic; padding:5px 0; }
  .u-ok { color:var(--good); font-weight:700; }
  .u-warn { color:var(--warn); font-weight:700; }
  .u-bad { color:var(--bad); font-weight:700; }
  .note { font-size:0.72rem; color:var(--sub); line-height:1.5; }
  .checkrow { display:flex; align-items:flex-start; gap:7px; font-size:0.78rem; margin-bottom:7px; }
  .scroll { max-height:250px; overflow:auto; }
  .banner { padding:7px 10px; border-radius:6px; font-size:0.77rem; margin-bottom:10px; }
  .banner-ok { background:#e7f5ec; color:var(--good); }
  .banner-err { background:#fdecea; color:var(--bad); }
  .banner-idle { background:var(--panel); color:var(--sub); }
</style>
</head>
<body>

<div class="ribbon">
  <div class="ribbon-top">
    <h1>REV3 Structural Solver</h1>
    <div class="sep"></div>
    <div class="ctl">
      <label for="sel-units">Units</label>
      <select id="sel-units" onchange="onUnitsChanged()"></select>
    </div>
    <div class="ctl">
      <label for="sel-material">Material</label>
      <select id="sel-material" onchange="onMaterialChanged()"></select>
    </div>
    <div class="ctl">
      <label for="sel-section">Section</label>
      <select id="sel-section" onchange="onSectionChanged()"></select>
    </div>
    <div class="sep"></div>
    <button class="ghost" onclick="runAnalysis()" style="font-weight:700;color:var(--accent);border-color:var(--accent);">&#9654; Solve</button>
  </div>
  <div class="ribbon-status">
    <span>Nodes: <b id="st-nodes">8</b></span>
    <span>Members: <b id="st-members">12</b></span>
    <span>DOF: <b id="st-dof">--</b></span>
    <span>Loads: <b id="st-loads">0</b></span>
    <span>Edge: <b id="st-edge">--</b></span>
    <span id="st-result"><span class="pill pill-idle">Not solved</span></span>
    <span style="margin-left:auto;">Solve: Direct Stiffness (JS) &middot; Excel-driven</span>
  </div>
</div>

<div class="layout">
  <div class="viewer-col">
    <canvas id="view3d"></canvas>
    <div class="view-controls">
      <button id="btn-loads" class="active" onclick="toggleLoads()">Load arrows</button>
      <button id="btn-deformed" onclick="toggleDeformed()">Deformed</button>
      <button id="btn-labels" class="active" onclick="toggleLabels()">Labels</button>
      <button onclick="resetView()">Reset view</button>
    </div>
    <div class="selection-box" id="selection-box"></div>
    <div class="view-hint">Drag to orbit &middot; Shift+drag to pan &middot; Wheel to zoom &middot; Click a node or member to select</div>
  </div>

  <div class="side-col">
    <div class="tabs">
      <button class="active" onclick="setTab(0)">Model</button>
      <button onclick="setTab(1)">Loads</button>
      <button onclick="setTab(2)">Results</button>
    </div>

    <!-- MODEL -->
    <div class="panel active" id="panel-0">
      <div class="card">
        <h3>Material <span style="font-weight:400;color:var(--sub);font-size:0.75rem;">(from Excel)</span></h3>
        <div class="kv" id="kv-material"></div>
      </div>
      <div class="card">
        <h3>Section <span style="font-weight:400;color:var(--sub);font-size:0.75rem;">(from Excel)</span></h3>
        <div class="kv" id="kv-section"></div>
      </div>
      <div class="card">
        <h3>Geometry &amp; supports</h3>
        <div class="kv" id="kv-geometry"></div>
      </div>
      <div class="card">
        <h3>Data sources</h3>
        <div class="note" id="src-list"></div>
      </div>
      <div class="card">
        <h3>Modelling assumptions</h3>
        <div class="note">
          Linear-elastic, first-order 3D space-frame analysis (axial + torsion +
          biaxial bending, Euler-Bernoulli, no shear deformation). All 12 members
          are rigidly (moment) connected at both ends; nodes 1&ndash;4 are pinned.
          Bending about a member's local y-axis uses the section's weak-axis
          property (Iy), about local z the strong-axis property (Ix) &mdash; a stated
          modelling convention. The utilisation ratio is an elastic superposition
          stress estimate, (|N|/A + |My|/Sy + |Mz|/Sx) / Fy, <b>not</b> a full AISC
          interaction-equation capacity check.
        </div>
      </div>
    </div>

    <!-- LOADS -->
    <div class="panel" id="panel-1">
      <div class="card">
        <h3>Add nodal load</h3>
        <div class="fld" style="margin-bottom:7px;">
          <label>Node (or click one in the 3D view)</label>
          <select id="ld-node" onchange="onLoadNodeChanged()"></select>
        </div>
        <div class="grid3" style="margin-bottom:7px;">
          <div class="fld"><label id="lbl-fx">Fx</label><input id="ld-fx" type="number" value="0" step="any"></div>
          <div class="fld"><label id="lbl-fy">Fy</label><input id="ld-fy" type="number" value="0" step="any"></div>
          <div class="fld"><label id="lbl-fz">Fz</label><input id="ld-fz" type="number" value="0" step="any"></div>
        </div>
        <div class="grid3" style="margin-bottom:9px;">
          <div class="fld"><label id="lbl-mx">Mx</label><input id="ld-mx" type="number" value="0" step="any"></div>
          <div class="fld"><label id="lbl-my">My</label><input id="ld-my" type="number" value="0" step="any"></div>
          <div class="fld"><label id="lbl-mz">Mz</label><input id="ld-mz" type="number" value="0" step="any"></div>
        </div>
        <button class="ghost" style="width:100%;" onclick="addLoad()">+ Add load</button>
      </div>

      <div class="card">
        <h3>Self-weight</h3>
        <div class="checkrow">
          <input type="checkbox" id="chk-selfweight" onchange="onSelfWeightToggled()">
          <label for="chk-selfweight">Include member self-weight (section area &times; material density from Excel), lumped to end nodes</label>
        </div>
      </div>

      <div class="card">
        <h3>Applied loads <button class="ghost" style="float:right;padding:2px 8px;" onclick="clearLoads()">Clear all</button></h3>
        <div id="load-list"></div>
      </div>

      <button class="primary" onclick="runAnalysis()">Solve</button>
    </div>

    <!-- RESULTS -->
    <div class="panel" id="panel-2">
      <div id="banner" class="banner banner-idle">Not solved yet &mdash; add loads and press Solve.</div>
      <div class="card">
        <h3>Summary</h3>
        <div class="kv" id="kv-summary"></div>
      </div>
      <div class="card">
        <h3>Nodal displacements <span id="disp-unit" style="font-weight:400;color:var(--sub);font-size:0.72rem;"></span></h3>
        <div class="scroll">
          <table id="tbl-disp"><thead><tr><th>Node</th><th>UX</th><th>UY</th><th>UZ</th><th>RX</th><th>RY</th><th>RZ</th></tr></thead><tbody></tbody></table>
        </div>
      </div>
      <div class="card">
        <h3>Support reactions <span id="react-unit" style="font-weight:400;color:var(--sub);font-size:0.72rem;"></span></h3>
        <div class="scroll">
          <table id="tbl-react"><thead><tr><th>Node</th><th>FX</th><th>FY</th><th>FZ</th></tr></thead><tbody></tbody></table>
        </div>
      </div>
      <div class="card">
        <h3>Member end forces &amp; utilisation</h3>
        <div class="scroll">
          <table id="tbl-member"><thead><tr><th>Mbr</th><th>End</th><th>N</th><th>Vy</th><th>Vz</th><th>T</th><th>My</th><th>Mz</th><th>Util</th></tr></thead><tbody></tbody></table>
        </div>
        <div class="note" style="margin-top:6px;" id="member-unit"></div>
      </div>
    </div>
  </div>
</div>

<script>
const CATALOG = __CATALOG_JSON__;
const GEOMETRY = __GEOMETRY_JSON__;
</script>
<script>
__SOLVER_CORE__
</script>
<script>
__VIEWER_JS__
</script>
<script>
// ---------------------------------------------------------------------------
// APP STATE + UI
// ---------------------------------------------------------------------------
let selection = Object.assign({}, CATALOG.start);
let currentModel = null;
let loads = [];
let lastResult = null;
let selectedNodeId = null;
let selectedMemberId = null;
let loadSeq = 1;

function includeSelfWeight() {
  const el = document.getElementById("chk-selfweight");
  return el ? el.checked : false;
}

function fmt(v, d) {
  if (v === null || v === undefined || Number.isNaN(v)) return "--";
  d = d === undefined ? 3 : d;
  const a = Math.abs(v);
  if (a !== 0 && (a < 1e-4 || a >= 1e7)) return Number(v).toExponential(2);
  return Number(v).toLocaleString(undefined, { minimumFractionDigits: d, maximumFractionDigits: d });
}

function setTab(i) {
  document.querySelectorAll(".panel").forEach((p, idx) => p.classList.toggle("active", idx === i));
  document.querySelectorAll(".tabs button").forEach((b, idx) => b.classList.toggle("active", idx === i));
}

// ---- dropdown population ----------------------------------------------
function populateControls() {
  const unitsSel = document.getElementById("sel-units");
  unitsSel.innerHTML = "";
  Object.entries(CATALOG.systems).forEach(([key, sys]) => {
    const o = document.createElement("option");
    o.value = key; o.textContent = sys.name;
    unitsSel.appendChild(o);
  });
  unitsSel.value = selection.units;

  populateMaterialSelect();
  populateSectionSelect();
  populateNodeSelect();
}

function populateMaterialSelect() {
  const sys = CATALOG.systems[selection.units];
  const sel = document.getElementById("sel-material");
  sel.innerHTML = "";
  // group by category so the list is navigable
  const byCat = {};
  sys.material_labels.forEach((lbl) => {
    const cat = sys.materials[lbl].category;
    (byCat[cat] = byCat[cat] || []).push(lbl);
  });
  Object.keys(byCat).sort().forEach((cat) => {
    const g = document.createElement("optgroup");
    g.label = cat;
    byCat[cat].forEach((lbl) => {
      const o = document.createElement("option");
      o.value = lbl;
      o.textContent = lbl + (lbl === CATALOG.declared_material_label ? "  (" + CATALOG.declared_material_name + ")" : "");
      g.appendChild(o);
    });
    sel.appendChild(g);
  });
  sel.value = selection.material;
}

function populateSectionSelect() {
  const sys = CATALOG.systems[selection.units];
  const sel = document.getElementById("sel-section");
  sel.innerHTML = "";
  sys.section_labels.forEach((lbl) => {
    const o = document.createElement("option");
    o.value = lbl; o.textContent = lbl;
    sel.appendChild(o);
  });
  sel.value = selection.section;
}

function populateNodeSelect() {
  const sel = document.getElementById("ld-node");
  const keep = sel.value;
  sel.innerHTML = "";
  GEOMETRY.nodes.forEach((n) => {
    const o = document.createElement("option");
    o.value = n.id;
    o.textContent = "Node " + n.id + (n.support ? " (pinned support)" : "");
    sel.appendChild(o);
  });
  if (keep) sel.value = keep;
}

// ---- selection change handlers ----------------------------------------
function onUnitsChanged() {
  const newUnits = document.getElementById("sel-units").value;
  const newSys = CATALOG.systems[newUnits];

  // keep the same physical shape across the unit switch (W150X22.5 <-> W6X15)
  const equiv = CATALOG.section_equiv[selection.section];
  const newSection = newSys.sections[equiv] ? equiv
    : (newSys.sections[selection.section] ? selection.section : newSys.section_labels[0]);
  // material labels are identical across both workbooks, but be defensive
  const newMaterial = newSys.materials[selection.material]
    ? selection.material : newSys.material_labels[0];

  // Convert already-entered load magnitudes into the new display units, so a
  // load the user typed keeps its PHYSICAL magnitude instead of being
  // silently reinterpreted (100 kN becomes 22.48 kip, not "100 kip").
  const fConv = unitForceConversion(selection.units, newUnits);
  const mConv = unitMomentConversion(selection.units, newUnits);
  loads.forEach((ld) => {
    ld.fx *= fConv; ld.fy *= fConv; ld.fz *= fConv;
    ld.mx *= mConv; ld.my *= mConv; ld.mz *= mConv;
  });

  selection = { units: newUnits, material: newMaterial, section: newSection };
  populateMaterialSelect();
  populateSectionSelect();
  rebuildModel();
}

// kN <-> kip and kN·m <-> kip·in, derived from the catalog's own scales:
// both systems' load_*_input_scale convert display -> solver units, and the
// solver base units (N vs kip, N·mm vs kip·in) differ, so go through the
// declared cross-system factors carried in the catalog.
function unitForceConversion(fromKey, toKey) {
  if (fromKey === toKey) return 1;
  const kNtoKip = CATALOG.force_kN_to_kip;
  return fromKey === "metric" ? kNtoKip : 1 / kNtoKip;
}
function unitMomentConversion(fromKey, toKey) {
  if (fromKey === toKey) return 1;
  const kNmToKipIn = CATALOG.moment_kNm_to_kipin;
  return fromKey === "metric" ? kNmToKipIn : 1 / kNmToKipIn;
}

function onMaterialChanged() {
  selection.material = document.getElementById("sel-material").value;
  rebuildModel();
}
function onSectionChanged() {
  selection.section = document.getElementById("sel-section").value;
  rebuildModel();
}
function onLoadNodeChanged() {
  selectedNodeId = parseInt(document.getElementById("ld-node").value, 10);
  selectedMemberId = null;
  onSelectionChanged();
}
function onSelfWeightToggled() {
  lastResult = null;
  renderStatus();
  drawScene();
}

// ---- model rebuild (the heart of live switching) -----------------------
function rebuildModel() {
  currentModel = deriveModel(selection);
  lastResult = null;           // previous results belong to the old model
  renderModelPanel();
  renderLoadLabels();
  renderLoadList();
  renderStatus();
  clearResultTables();
  drawScene();
}

function renderModelPanel() {
  const M = currentModel, sys = M.sys, mat = M.material, sec = M.section;
  const su = sys.section_unit, lu = sys.solver_length_unit;

  document.getElementById("kv-material").innerHTML = `
    <div class="k">Label</div><div class="v">${mat.label}</div>
    <div class="k">Category</div><div class="v">${mat.category}</div>
    <div class="k">E</div><div class="v">${fmt(mat.E,0)} ${mat.e_unit}</div>
    <div class="k">G</div><div class="v">${fmt(mat.G,0)} ${mat.g_unit}</div>
    <div class="k">Fy</div><div class="v">${fmt(mat.Fy,1)} ${mat.fy_unit}</div>
    <div class="k">Fu</div><div class="v">${mat.Fu === null ? "--" : fmt(mat.Fu,1) + " " + mat.fu_unit}</div>
    <div class="k">Density</div><div class="v">${mat.density === null ? "--" : fmt(mat.density,2) + " " + mat.density_unit}</div>`;

  document.getElementById("kv-section").innerHTML = `
    <div class="k">Designation</div><div class="v">${sec.label}</div>
    <div class="k">d</div><div class="v">${fmt(sec.d,1)} ${su}</div>
    <div class="k">bf</div><div class="v">${fmt(sec.bf,1)} ${su}</div>
    <div class="k">tf / tw</div><div class="v">${fmt(sec.tf,2)} / ${fmt(sec.tw,2)} ${su}</div>
    <div class="k">A</div><div class="v">${fmt(sec.A_display,2)} ${su}\u00b2</div>
    <div class="k">Ix (strong)</div><div class="v">${fmt(sec.Ix,0)} ${lu}\u2074</div>
    <div class="k">Iy (weak)</div><div class="v">${fmt(sec.Iy,0)} ${lu}\u2074</div>
    <div class="k">J</div><div class="v">${fmt(sec.J,0)} ${lu}\u2074</div>
    <div class="k">Sx / Sy</div><div class="v">${fmt(sec.Sx,0)} / ${fmt(sec.Sy,0)} ${lu}\u00b3</div>`;

  document.getElementById("kv-geometry").innerHTML = `
    <div class="k">Cube edge</div><div class="v">${fmt(sys.edge_length,3)} ${sys.length_unit}</div>
    <div class="k">Nodes / Members</div><div class="v">${M.nodes.length} / ${M.members.length}</div>
    <div class="k">Supports</div><div class="v">Pinned, nodes 1-4</div>
    <div class="k">Member ends</div><div class="v">Fixed (rigid)</div>
    <div class="k">Total DOF</div><div class="v">${GEOMETRY.dof.total}</div>
    <div class="k">Restrained</div><div class="v">${GEOMETRY.dof.restrained}</div>
    <div class="k">Active DOF</div><div class="v">${GEOMETRY.dof.active}</div>
    <div class="k">Solver length unit</div><div class="v">${lu}</div>`;

  document.getElementById("src-list").innerHTML =
    `units/ &rarr; ${CATALOG.sources.units}<br>` +
    `materials/ &rarr; ${sys.materials_source}<br>` +
    `member_size/ &rarr; ${CATALOG.sources.member_size}<br>` +
    `<span style="color:var(--sub)">${Object.keys(sys.materials).length} materials, ` +
    `${sys.section_labels.length} W-shapes loaded for ${sys.name}</span>`;
}

function renderLoadLabels() {
  const sys = currentModel.sys;
  ["fx","fy","fz"].forEach((k) => {
    document.getElementById("lbl-" + k).textContent =
      k.toUpperCase().replace("F","F") + " (" + sys.display_force_unit + ")";
  });
  ["mx","my","mz"].forEach((k) => {
    document.getElementById("lbl-" + k).textContent =
      k.toUpperCase() + " (" + sys.display_moment_unit + ")";
  });
  document.getElementById("disp-unit").textContent = "(" + sys.solver_length_unit + " / rad)";
  document.getElementById("react-unit").textContent = "(" + sys.display_force_unit + ")";
  document.getElementById("member-unit").textContent =
    "Forces in " + sys.display_force_unit + ", moments in " + sys.display_moment_unit +
    ". Util = (|N|/A + |My|/Sy + |Mz|/Sx) / Fy, elastic estimate only.";
}

// ---- loads -------------------------------------------------------------
function addLoad() {
  const node = parseInt(document.getElementById("ld-node").value, 10);
  const get = (id) => parseFloat(document.getElementById(id).value) || 0;
  const ld = { id: loadSeq++, node,
    fx: get("ld-fx"), fy: get("ld-fy"), fz: get("ld-fz"),
    mx: get("ld-mx"), my: get("ld-my"), mz: get("ld-mz") };
  if (!ld.fx && !ld.fy && !ld.fz && !ld.mx && !ld.my && !ld.mz) return;
  loads.push(ld);
  ["ld-fx","ld-fy","ld-fz","ld-mx","ld-my","ld-mz"].forEach((id) => {
    document.getElementById(id).value = 0;
  });
  lastResult = null;
  renderLoadList();
  renderStatus();
  clearResultTables();
  drawScene();
}

function removeLoad(id) {
  loads = loads.filter((l) => l.id !== id);
  lastResult = null;
  renderLoadList();
  renderStatus();
  clearResultTables();
  drawScene();
}

function clearLoads() {
  loads = [];
  lastResult = null;
  renderLoadList();
  renderStatus();
  clearResultTables();
  drawScene();
}

function renderLoadList() {
  const box = document.getElementById("load-list");
  if (loads.length === 0) {
    box.innerHTML = '<div class="empty">No nodal loads yet. Add one above, or enable self-weight.</div>';
  } else {
    const sys = currentModel.sys;
    const fu = sys.display_force_unit, mu = sys.display_moment_unit;
    box.innerHTML = loads.map((l) => {
      const parts = [];
      if (l.fx) parts.push(`Fx ${fmt(l.fx,2)} ${fu}`);
      if (l.fy) parts.push(`Fy ${fmt(l.fy,2)} ${fu}`);
      if (l.fz) parts.push(`Fz ${fmt(l.fz,2)} ${fu}`);
      if (l.mx) parts.push(`Mx ${fmt(l.mx,2)} ${mu}`);
      if (l.my) parts.push(`My ${fmt(l.my,2)} ${mu}`);
      if (l.mz) parts.push(`Mz ${fmt(l.mz,2)} ${mu}`);
      return `<div class="loadrow"><span><b>Node ${l.node}</b> &mdash; ${parts.join(", ")}</span>
              <button onclick="removeLoad(${l.id})" title="Remove">&times;</button></div>`;
    }).join("");
  }
  renderStatus();
}

// ---- solve -------------------------------------------------------------
function runAnalysis() {
  const banner = document.getElementById("banner");
  const sw = includeSelfWeight();
  if (loads.length === 0 && !sw) {
    banner.className = "banner banner-err";
    banner.textContent = "Add at least one load, or enable self-weight, before solving.";
    setTab(2);
    return;
  }
  try {
    lastResult = solveModel(currentModel, loads, sw);
    banner.className = "banner banner-ok";
    banner.textContent = `Analysis complete \u2014 no errors. ${loads.length} nodal load(s)` +
      (sw ? " + self-weight" : "") + `, ${GEOMETRY.dof.active} active DOF.`;
    renderResults();
    renderStatus();
    view.showDeformed = true;
    document.getElementById("btn-deformed").classList.add("active");
    drawScene();
    setTab(2);
  } catch (err) {
    lastResult = null;
    banner.className = "banner banner-err";
    banner.textContent = "Solve failed: " + err.message;
    renderStatus();
    setTab(2);
  }
}

function clearResultTables() {
  ["tbl-disp","tbl-react","tbl-member"].forEach((id) => {
    document.querySelector("#" + id + " tbody").innerHTML = "";
  });
  document.getElementById("kv-summary").innerHTML =
    '<div class="k">Status</div><div class="v">Not solved</div>';
  const banner = document.getElementById("banner");
  banner.className = "banner banner-idle";
  banner.textContent = "Model changed \u2014 press Solve to analyse.";
}

function renderResults() {
  const M = currentModel, sys = M.sys, R = lastResult;

  const utilClass = (u) => u > 1 ? "u-bad" : (u > 0.7 ? "u-warn" : "u-ok");
  document.getElementById("kv-summary").innerHTML = `
    <div class="k">Max displacement</div><div class="v">${fmt(R.maxDisp,4)} ${sys.solver_length_unit} (node ${R.maxDispNode})</div>
    <div class="k">Max utilisation</div><div class="v"><span class="${utilClass(R.maxUtil)}">${fmt(R.maxUtil,3)}</span></div>
    <div class="k">Material</div><div class="v">${M.material.label}</div>
    <div class="k">Section</div><div class="v">${M.section.label}</div>
    <div class="k">Units</div><div class="v">${sys.name}</div>`;

  const dispBody = document.querySelector("#tbl-disp tbody");
  dispBody.innerHTML = M.nodes.map((nd) => {
    const b = (nd.id - 1) * 6;
    const u = R.u.slice(b, b + 6);
    const cls = selectedNodeId === nd.id ? ' class="sel"' : "";
    return `<tr${cls}><td>${nd.id}${nd.support ? " (P)" : ""}</td>
      <td>${fmt(u[0],4)}</td><td>${fmt(u[1],4)}</td><td>${fmt(u[2],4)}</td>
      <td>${fmt(u[3],5)}</td><td>${fmt(u[4],5)}</td><td>${fmt(u[5],5)}</td></tr>`;
  }).join("");

  const reactBody = document.querySelector("#tbl-react tbody");
  reactBody.innerHTML = M.nodes.filter((n) => n.support).map((nd) => {
    const b = (nd.id - 1) * 6;
    const rx = (R.reactions[b + 0] || 0) * sys.force_display_scale;
    const ry = (R.reactions[b + 1] || 0) * sys.force_display_scale;
    const rz = (R.reactions[b + 2] || 0) * sys.force_display_scale;
    return `<tr><td>${nd.id}</td><td>${fmt(rx,3)}</td><td>${fmt(ry,3)}</td><td>${fmt(rz,3)}</td></tr>`;
  }).join("");

  const memberBody = document.querySelector("#tbl-member tbody");
  let html = "";
  R.memberResults.forEach((mr) => {
    ["i","j"].forEach((end, idx) => {
      const off = idx === 0 ? 0 : 6;
      const f = mr.forces_local;
      const cls = selectedMemberId === mr.id ? ' class="sel clickable"' : ' class="clickable"';
      html += `<tr${cls} onclick="selectMember(${mr.id})">
        <td>M${mr.id} ${mr.type.replace(" Beam","")}</td><td>${end}</td>
        <td>${fmt(f[off+0]*sys.force_display_scale,2)}</td>
        <td>${fmt(f[off+1]*sys.force_display_scale,2)}</td>
        <td>${fmt(f[off+2]*sys.force_display_scale,2)}</td>
        <td>${fmt(f[off+3]*sys.moment_display_scale,2)}</td>
        <td>${fmt(f[off+4]*sys.moment_display_scale,2)}</td>
        <td>${fmt(f[off+5]*sys.moment_display_scale,2)}</td>
        <td class="${utilClass(mr.util)}">${fmt(mr.util,2)}</td></tr>`;
    });
  });
  memberBody.innerHTML = html;
}

function renderStatus() {
  const M = currentModel;
  document.getElementById("st-nodes").textContent = M.nodes.length;
  document.getElementById("st-members").textContent = M.members.length;
  document.getElementById("st-dof").textContent = GEOMETRY.dof.active + " active / " + GEOMETRY.dof.total;
  document.getElementById("st-loads").textContent = loads.length + (includeSelfWeight() ? " + SW" : "");
  document.getElementById("st-edge").textContent = fmt(M.sys.edge_length,3) + " " + M.sys.length_unit;
  const el = document.getElementById("st-result");
  if (!lastResult) {
    el.innerHTML = '<span class="pill pill-idle">Not solved</span>';
  } else {
    const u = lastResult.maxUtil;
    const cls = u > 1 ? "pill-bad" : (u > 0.7 ? "pill-warn" : "pill-ok");
    el.innerHTML = `<span class="pill ${cls}">Solved &middot; max util ${fmt(u,2)}</span>`;
  }
}

// ---- selection ---------------------------------------------------------
function selectMember(id) {
  selectedMemberId = id;
  selectedNodeId = null;
  onSelectionChanged();
}

function onSelectionChanged() {
  const box = document.getElementById("selection-box");
  const M = currentModel;
  if (selectedNodeId !== null) {
    const n = M.nodes.find((x) => x.id === selectedNodeId);
    const sys = M.sys;
    let html = `<h4>Node ${n.id}${n.support ? " (pinned)" : ""}</h4>`;
    html += `<div class="note">x, y, z = ${fmt(n.x,1)}, ${fmt(n.y,1)}, ${fmt(n.z,1)} ${sys.solver_length_unit}</div>`;
    if (lastResult) {
      const b = (n.id - 1) * 6;
      html += `<div class="note" style="margin-top:4px;">U = ${fmt(lastResult.u[b],4)}, ${fmt(lastResult.u[b+1],4)}, ${fmt(lastResult.u[b+2],4)} ${sys.solver_length_unit}</div>`;
    }
    html += `<div class="note" style="margin-top:4px;color:var(--accent)">Selected as load target.</div>`;
    box.innerHTML = html;
    box.classList.add("show");
  } else if (selectedMemberId !== null) {
    const m = M.members.find((x) => x.id === selectedMemberId);
    let html = `<h4>Member M${m.id} &mdash; ${m.type}</h4>`;
    html += `<div class="note">Nodes ${m.i} &rarr; ${m.j}, L = ${fmt(m.length,1)} ${M.sys.solver_length_unit}</div>`;
    if (lastResult) {
      const mr = lastResult.memberResults.find((r) => r.id === m.id);
      html += `<div class="note" style="margin-top:4px;">Utilisation = <b>${fmt(mr.util,3)}</b></div>`;
    }
    box.innerHTML = html;
    box.classList.add("show");
  } else {
    box.classList.remove("show");
  }
  if (lastResult) renderResults();
  drawScene();
}

// ---- view toggles ------------------------------------------------------
function toggleLoads() {
  view.showLoads = !view.showLoads;
  document.getElementById("btn-loads").classList.toggle("active", view.showLoads);
  drawScene();
}
function toggleDeformed() {
  view.showDeformed = !view.showDeformed;
  document.getElementById("btn-deformed").classList.toggle("active", view.showDeformed);
  drawScene();
}
function toggleLabels() {
  view.showLabels = !view.showLabels;
  document.getElementById("btn-labels").classList.toggle("active", view.showLabels);
  drawScene();
}

// ---- boot --------------------------------------------------------------
populateControls();
rebuildModel();
initViewerEvents();
renderLoadList();
drawScene();
</script>
</body>
</html>
"""

SOLVER_CORE_JS = r"""// ---------------------------------------------------------------------------
// SOLVER CORE
// Derives the analysis model from the current catalog selection (units /
// material / section), then runs a 3D space-frame finite-element solve:
// 12x12 Euler-Bernoulli element stiffness, global assembly, partitioned
// solve, member end-force recovery. Independently cross-validated against a
// numpy reference implementation to ~1e-10 relative agreement.
// ---------------------------------------------------------------------------

function zeros(rows, cols) {
  const M = new Array(rows);
  for (let i = 0; i < rows; i++) M[i] = new Array(cols).fill(0);
  return M;
}

function matMul(A, B) {
  const n = A.length, p = B.length, m = B[0].length;
  const C = zeros(n, m);
  for (let i = 0; i < n; i++) {
    for (let k = 0; k < p; k++) {
      const a = A[i][k];
      if (a === 0) continue;
      for (let j = 0; j < m; j++) C[i][j] += a * B[k][j];
    }
  }
  return C;
}

function transpose(A) {
  const n = A.length, m = A[0].length;
  const T = zeros(m, n);
  for (let i = 0; i < n; i++) for (let j = 0; j < m; j++) T[j][i] = A[i][j];
  return T;
}

function matVec(A, v) {
  const n = A.length;
  const r = new Array(n).fill(0);
  for (let i = 0; i < n; i++) {
    let s = 0;
    for (let j = 0; j < v.length; j++) s += A[i][j] * v[j];
    r[i] = s;
  }
  return r;
}

function nodeDofIndices(nodeId) {
  const base = (nodeId - 1) * 6;
  return [base, base + 1, base + 2, base + 3, base + 4, base + 5];
}

// Build the live analysis model from the current selection. Node coordinates
// are the dimensionless cube coefficients scaled by the active system's edge
// length, converted into that system's solver-consistent length unit.
function deriveModel(sel) {
  const sys = CATALOG.systems[sel.units];
  const mat = sys.materials[sel.material];
  const sec = sys.sections[sel.section];
  const edgeSolver = sys.edge_length * sys.node_to_solver_length;

  const nodes = GEOMETRY.nodes.map((n) => ({
    id: n.id,
    x: n.cx * edgeSolver, y: n.cy * edgeSolver, z: n.cz * edgeSolver,
    support: n.support, restraint: n.restraint,
  }));

  const members = GEOMETRY.members.map((m) => {
    const ni = nodes.find((n) => n.id === m.i);
    const nj = nodes.find((n) => n.id === m.j);
    const L = Math.hypot(nj.x - ni.x, nj.y - ni.y, nj.z - ni.z);
    return Object.assign({}, m, { length: L });
  });

  return { sys, material: mat, section: sec, nodes, members, edgeSolver };
}

// 12x12 space-frame element stiffness. Local DOF order per node:
// [u, v, w, rx, ry, rz]. Iy = bending about local y (weak axis here),
// Iz = bending about local z (strong axis here).
function localStiffness(E, G, A, Iy, Iz, J, L) {
  const k = zeros(12, 12);
  const EAL = E * A / L;
  const GJL = G * J / L;

  k[0][0] = k[6][6] = EAL;
  k[0][6] = k[6][0] = -EAL;

  k[3][3] = k[9][9] = GJL;
  k[3][9] = k[9][3] = -GJL;

  const L2 = L * L, L3 = L2 * L;
  const EIz = E * Iz, EIy = E * Iy;

  k[1][1] = k[7][7] = 12 * EIz / L3;
  k[1][7] = k[7][1] = -12 * EIz / L3;
  k[1][5] = k[5][1] = 6 * EIz / L2;
  k[1][11] = k[11][1] = 6 * EIz / L2;
  k[7][5] = k[5][7] = -6 * EIz / L2;
  k[7][11] = k[11][7] = -6 * EIz / L2;
  k[5][5] = k[11][11] = 4 * EIz / L;
  k[5][11] = k[11][5] = 2 * EIz / L;

  k[2][2] = k[8][8] = 12 * EIy / L3;
  k[2][8] = k[8][2] = -12 * EIy / L3;
  k[2][4] = k[4][2] = -6 * EIy / L2;
  k[2][10] = k[10][2] = -6 * EIy / L2;
  k[8][4] = k[4][8] = 6 * EIy / L2;
  k[8][10] = k[10][8] = 6 * EIy / L2;
  k[4][4] = k[10][10] = 4 * EIy / L;
  k[4][10] = k[10][4] = 2 * EIy / L;

  return k;
}

function transformMatrix(lx, ly, lz) {
  const R = [lx, ly, lz];
  const T = zeros(12, 12);
  for (let b = 0; b < 4; b++) {
    for (let r = 0; r < 3; r++) {
      for (let c = 0; c < 3; c++) T[b * 3 + r][b * 3 + c] = R[r][c];
    }
  }
  return T;
}

function assembleGlobalStiffness(M) {
  const n = M.nodes.length * 6;
  const K = zeros(n, n);
  M.members.forEach((m) => {
    const kl = localStiffness(M.material.E, M.material.G, M.section.A,
      M.section.Iy, M.section.Ix, M.section.J, m.length);
    const T = transformMatrix(m.local_x, m.local_y, m.local_z);
    const kg = matMul(matMul(transpose(T), kl), T);
    const dof = nodeDofIndices(m.i).concat(nodeDofIndices(m.j));
    for (let a = 0; a < 12; a++) {
      for (let b = 0; b < 12; b++) K[dof[a]][dof[b]] += kg[a][b];
    }
  });
  return K;
}

// Gaussian elimination with partial pivoting -- small dense systems only
// (this model has at most 36 free DOF).
function gaussSolve(Amat, bvec) {
  const n = bvec.length;
  const A = Amat.map((row) => row.slice());
  const b = bvec.slice();
  for (let col = 0; col < n; col++) {
    let piv = col, maxAbs = Math.abs(A[col][col]);
    for (let r = col + 1; r < n; r++) {
      if (Math.abs(A[r][col]) > maxAbs) { maxAbs = Math.abs(A[r][col]); piv = r; }
    }
    if (maxAbs < 1e-9) {
      throw new Error("Singular stiffness matrix at column " + col +
        " -- the structure may be an unstable mechanism for this load case.");
    }
    if (piv !== col) {
      const tr = A[col]; A[col] = A[piv]; A[piv] = tr;
      const tb = b[col]; b[col] = b[piv]; b[piv] = tb;
    }
    const diag = A[col][col];
    for (let r = col + 1; r < n; r++) {
      const f = A[r][col] / diag;
      if (f === 0) continue;
      for (let c = col; c < n; c++) A[r][c] -= f * A[col][c];
      b[r] -= f * b[col];
    }
  }
  const x = new Array(n).fill(0);
  for (let r = n - 1; r >= 0; r--) {
    let s = b[r];
    for (let c = r + 1; c < n; c++) s -= A[r][c] * x[c];
    x[r] = s / A[r][r];
  }
  return x;
}

// Self-weight, lumped half-member-weight to each end node, applied in -Y.
// Density from the materials workbook is a unit weight (kN/m^3 or kip/ft^3);
// density_to_solver converts it to the solver-consistent force/length^3.
function selfWeightNodalForces(M) {
  const dens = M.material.density;
  if (!dens) return {};
  const densSolver = dens * M.sys.density_to_solver;
  const nodal = {};
  M.members.forEach((m) => {
    const w = densSolver * M.section.A * m.length;
    [m.i, m.j].forEach((n) => { nodal[n] = (nodal[n] || 0) + w / 2; });
  });
  return nodal;
}

// loads: [{node, fx, fy, fz, mx, my, mz}] in DISPLAY units (kN / kN·m, or
// kip / kip·in). includeSelfWeight adds lumped member self-weight.
function solveModel(M, loads, includeSelfWeight) {
  const n = M.nodes.length * 6;
  const K = assembleGlobalStiffness(M);
  const F = new Array(n).fill(0);

  loads.forEach((ld) => {
    const b = (ld.node - 1) * 6;
    F[b + 0] += (ld.fx || 0) * M.sys.load_force_input_scale;
    F[b + 1] += (ld.fy || 0) * M.sys.load_force_input_scale;
    F[b + 2] += (ld.fz || 0) * M.sys.load_force_input_scale;
    F[b + 3] += (ld.mx || 0) * M.sys.load_moment_input_scale;
    F[b + 4] += (ld.my || 0) * M.sys.load_moment_input_scale;
    F[b + 5] += (ld.mz || 0) * M.sys.load_moment_input_scale;
  });

  let swForces = {};
  if (includeSelfWeight) {
    swForces = selfWeightNodalForces(M);
    Object.entries(swForces).forEach(([node, w]) => {
      F[(parseInt(node, 10) - 1) * 6 + 1] -= w;   // already in solver force units
    });
  }

  const restrained = [], free = [];
  M.nodes.forEach((nd) => {
    const b = (nd.id - 1) * 6;
    for (let d = 0; d < 6; d++) {
      if (nd.restraint[d] === 1) restrained.push(b + d); else free.push(b + d);
    }
  });

  const nf = free.length;
  const Kff = zeros(nf, nf);
  for (let a = 0; a < nf; a++) {
    for (let b = 0; b < nf; b++) Kff[a][b] = K[free[a]][free[b]];
  }
  const uf = gaussSolve(Kff, free.map((i) => F[i]));

  const u = new Array(n).fill(0);
  free.forEach((idx, k) => { u[idx] = uf[k]; });

  const reactions = {};
  restrained.forEach((idx) => {
    let s = 0;
    for (let c = 0; c < n; c++) s += K[idx][c] * u[c];
    reactions[idx] = s - F[idx];
  });

  const memberResults = M.members.map((m) => {
    const T = transformMatrix(m.local_x, m.local_y, m.local_z);
    const dof = nodeDofIndices(m.i).concat(nodeDofIndices(m.j));
    const ul = matVec(T, dof.map((i) => u[i]));
    const kl = localStiffness(M.material.E, M.material.G, M.section.A,
      M.section.Iy, M.section.Ix, M.section.J, m.length);
    const fl = matVec(kl, ul);

    // Elastic utilisation, computed in solver-consistent units so the stress
    // lands in the material's own stress unit (MPa / ksi) alongside Fy.
    let maxUtil = 0;
    for (const off of [0, 6]) {
      const stress = Math.abs(fl[off + 0]) / M.section.A
        + Math.abs(fl[off + 4]) / M.section.Sy
        + Math.abs(fl[off + 5]) / M.section.Sx;
      maxUtil = Math.max(maxUtil, stress / M.material.Fy);
    }
    return { id: m.id, i: m.i, j: m.j, type: m.type, forces_local: fl, util: maxUtil };
  });

  const maxUtil = memberResults.reduce((a, m) => Math.max(a, m.util), 0);
  let maxDisp = 0, maxDispNode = null;
  M.nodes.forEach((nd) => {
    const b = (nd.id - 1) * 6;
    const d = Math.hypot(u[b], u[b + 1], u[b + 2]);
    if (d > maxDisp) { maxDisp = d; maxDispNode = nd.id; }
  });

  return { u, reactions, restrained, free, memberResults, maxUtil, maxDisp, maxDispNode, swForces };
}
"""

VIEWER_JS = r"""// ---------------------------------------------------------------------------
// VIEWER  -  dependency-free 3D wireframe renderer on a plain 2D canvas.
// Manual rotation + perspective projection, mouse-drag orbit, wheel zoom.
// Deliberately uses NO external library: an earlier revision loaded three.js
// from a CDN and the 3D view silently failed on machines where that fetch
// was blocked (corporate network, offline, file:// restrictions). This has
// no network dependency at all.
// ---------------------------------------------------------------------------

const view = {
  yaw: -0.6, pitch: -0.35, zoom: 1.0,
  panX: 0, panY: 0,
  dragging: false, panning: false, lastX: 0, lastY: 0,
  showLoads: true, showDeformed: false, showLabels: true,
  dispScale: null,   // null = auto
};

function rotatePoint(p, yaw, pitch) {
  // yaw about global Y (vertical), then pitch about the camera's X.
  const cy = Math.cos(yaw), sy = Math.sin(yaw);
  const x1 = p[0] * cy + p[2] * sy;
  const z1 = -p[0] * sy + p[2] * cy;
  const y1 = p[1];
  const cp = Math.cos(pitch), sp = Math.sin(pitch);
  const y2 = y1 * cp - z1 * sp;
  const z2 = y1 * sp + z1 * cp;
  return [x1, y2, z2];
}

function makeProjector(canvas, M) {
  const w = canvas.clientWidth, h = canvas.clientHeight;
  const nodes = M.nodes;
  const cx = nodes.reduce((a, n) => a + n.x, 0) / nodes.length;
  const cy = nodes.reduce((a, n) => a + n.y, 0) / nodes.length;
  const cz = nodes.reduce((a, n) => a + n.z, 0) / nodes.length;
  let span = 0;
  nodes.forEach((n) => {
    span = Math.max(span, Math.hypot(n.x - cx, n.y - cy, n.z - cz));
  });
  span = span || 1;

  const fit = Math.min(w, h) * 0.34 * view.zoom;
  const camDist = span * 3.2;

  return function project(x, y, z) {
    const p = rotatePoint([x - cx, y - cy, z - cz], view.yaw, view.pitch);
    const depth = p[2] + camDist;
    const persp = camDist / Math.max(depth, camDist * 0.15);
    return {
      sx: w / 2 + view.panX + (p[0] / span) * fit * persp,
      sy: h / 2 + view.panY - (p[1] / span) * fit * persp,
      depth: depth,
    };
  };
}

function drawScene() {
  const canvas = document.getElementById("view3d");
  if (!canvas) return;
  const ctx = canvas.getContext("2d");
  const dpr = window.devicePixelRatio || 1;
  const w = canvas.clientWidth, h = canvas.clientHeight;
  if (canvas.width !== w * dpr || canvas.height !== h * dpr) {
    canvas.width = w * dpr;
    canvas.height = h * dpr;
  }
  ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
  ctx.clearRect(0, 0, w, h);

  const M = currentModel;
  if (!M) return;
  const project = makeProjector(canvas, M);
  const nodeById = {};
  M.nodes.forEach((n) => { nodeById[n.id] = n; });

  // ---- ground grid, for depth cues ----
  const edge = M.edgeSolver;
  ctx.strokeStyle = "#e6e9ed";
  ctx.lineWidth = 1;
  const div = 4;
  for (let i = -1; i <= div + 1; i++) {
    const t = (i / div) * edge;
    const a1 = project(t, 0, -edge * 0.25), a2 = project(t, 0, edge * 1.25);
    ctx.beginPath(); ctx.moveTo(a1.sx, a1.sy); ctx.lineTo(a2.sx, a2.sy); ctx.stroke();
    const b1 = project(-edge * 0.25, 0, t), b2 = project(edge * 1.25, 0, t);
    ctx.beginPath(); ctx.moveTo(b1.sx, b1.sy); ctx.lineTo(b2.sx, b2.sy); ctx.stroke();
  }

  // ---- undeformed members, painter's algorithm (far to near) ----
  const memberDraw = M.members.map((m) => {
    const ni = nodeById[m.i], nj = nodeById[m.j];
    const pi = project(ni.x, ni.y, ni.z);
    const pj = project(nj.x, nj.y, nj.z);
    return { m, pi, pj, depth: (pi.depth + pj.depth) / 2 };
  }).sort((a, b) => b.depth - a.depth);

  memberDraw.forEach((d) => {
    const selected = selectedMemberId === d.m.id;
    ctx.strokeStyle = selected ? "#f59e0b" : d.m.color;
    ctx.lineWidth = selected ? 5 : 3.5;
    ctx.beginPath();
    ctx.moveTo(d.pi.sx, d.pi.sy);
    ctx.lineTo(d.pj.sx, d.pj.sy);
    ctx.stroke();
  });

  // ---- deformed shape overlay ----
  if (view.showDeformed && lastResult) {
    const s = deformedScale();
    ctx.strokeStyle = "#e63946";
    ctx.lineWidth = 2;
    ctx.setLineDash([6, 4]);
    M.members.forEach((m) => {
      const ni = nodeById[m.i], nj = nodeById[m.j];
      const bi = (m.i - 1) * 6, bj = (m.j - 1) * 6;
      const pi = project(ni.x + lastResult.u[bi] * s, ni.y + lastResult.u[bi + 1] * s, ni.z + lastResult.u[bi + 2] * s);
      const pj = project(nj.x + lastResult.u[bj] * s, nj.y + lastResult.u[bj + 1] * s, nj.z + lastResult.u[bj + 2] * s);
      ctx.beginPath();
      ctx.moveTo(pi.sx, pi.sy);
      ctx.lineTo(pj.sx, pj.sy);
      ctx.stroke();
    });
    ctx.setLineDash([]);
  }

  // ---- nodes, supports, labels ----
  M.nodes.forEach((n) => {
    const p = project(n.x, n.y, n.z);
    const selected = selectedNodeId === n.id;
    ctx.beginPath();
    ctx.arc(p.sx, p.sy, selected ? 8 : 6, 0, Math.PI * 2);
    ctx.fillStyle = selected ? "#f59e0b" : (n.support ? "#b42318" : "#333");
    ctx.fill();
    ctx.strokeStyle = "#fff";
    ctx.lineWidth = 1.5;
    ctx.stroke();

    if (n.support) {
      // little pinned-support triangle under the node
      ctx.beginPath();
      ctx.moveTo(p.sx, p.sy + 6);
      ctx.lineTo(p.sx - 9, p.sy + 20);
      ctx.lineTo(p.sx + 9, p.sy + 20);
      ctx.closePath();
      ctx.fillStyle = "#b45309";
      ctx.fill();
      ctx.beginPath();
      ctx.moveTo(p.sx - 13, p.sy + 20);
      ctx.lineTo(p.sx + 13, p.sy + 20);
      ctx.strokeStyle = "#b45309";
      ctx.lineWidth = 2.5;
      ctx.stroke();
    }

    if (view.showLabels) {
      ctx.fillStyle = "#1b1f24";
      ctx.font = "600 12px -apple-system, Segoe UI, Helvetica, Arial, sans-serif";
      ctx.fillText("N" + n.id, p.sx + 9, p.sy - 8);
    }
  });

  // ---- load arrows ----
  if (view.showLoads) drawLoadArrows(ctx, project, M);

  // ---- axis triad, bottom-left ----
  drawAxisTriad(ctx, w, h);

  // ---- deformation scale caption, so the exaggeration is never misleading ----
  if (view.showDeformed && lastResult) {
    ctx.fillStyle = "#e63946";
    ctx.font = "600 11px -apple-system, Segoe UI, Helvetica, Arial, sans-serif";
    ctx.fillText("Deformed shape \u00d7" + deformedScale().toFixed(1) + " (exaggerated)", 12, 20);
  }
}

function drawLoadArrows(ctx, project, M) {
  const byNode = {};
  loads.forEach((ld) => {
    if (!byNode[ld.node]) byNode[ld.node] = { fx: 0, fy: 0, fz: 0 };
    byNode[ld.node].fx += ld.fx || 0;
    byNode[ld.node].fy += ld.fy || 0;
    byNode[ld.node].fz += ld.fz || 0;
  });
  if (includeSelfWeight() && currentModel) {
    const sw = selfWeightNodalForces(currentModel);
    Object.entries(sw).forEach(([node, w]) => {
      if (!byNode[node]) byNode[node] = { fx: 0, fy: 0, fz: 0 };
      byNode[node].fy -= w * currentModel.sys.force_display_scale;
    });
  }

  // Scale arrow length by magnitude relative to the largest load present, so
  // a small self-weight lump doesn't draw the same size as a 100 kN point
  // load. Clamped to a sensible min/max fraction of the cube edge.
  let maxMag = 0;
  Object.values(byNode).forEach((f) => {
    maxMag = Math.max(maxMag, Math.abs(f.fx), Math.abs(f.fy), Math.abs(f.fz));
  });
  if (maxMag <= 0) return;
  const maxLen = M.edgeSolver * 0.30;
  const minLen = M.edgeSolver * 0.08;

  Object.entries(byNode).forEach(([nodeId, f]) => {
    const n = M.nodes.find((nd) => nd.id === parseInt(nodeId, 10));
    if (!n) return;
    [["fx", [1, 0, 0]], ["fy", [0, 1, 0]], ["fz", [0, 0, 1]]].forEach(([key, axis]) => {
      const val = f[key];
      if (!val) return;
      const sgn = Math.sign(val);
      const frac = Math.abs(val) / maxMag;
      const arrowLen = minLen + (maxLen - minLen) * frac;
      const tail = [
        n.x - axis[0] * arrowLen * sgn,
        n.y - axis[1] * arrowLen * sgn,
        n.z - axis[2] * arrowLen * sgn,
      ];
      const pTail = project(tail[0], tail[1], tail[2]);
      const pHead = project(n.x, n.y, n.z);
      ctx.strokeStyle = "#e63946";
      ctx.fillStyle = "#e63946";
      ctx.lineWidth = 2.5;
      ctx.beginPath();
      ctx.moveTo(pTail.sx, pTail.sy);
      ctx.lineTo(pHead.sx, pHead.sy);
      ctx.stroke();
      const ang = Math.atan2(pHead.sy - pTail.sy, pHead.sx - pTail.sx);
      const hl = 11;
      ctx.beginPath();
      ctx.moveTo(pHead.sx, pHead.sy);
      ctx.lineTo(pHead.sx - hl * Math.cos(ang - 0.4), pHead.sy - hl * Math.sin(ang - 0.4));
      ctx.lineTo(pHead.sx - hl * Math.cos(ang + 0.4), pHead.sy - hl * Math.sin(ang + 0.4));
      ctx.closePath();
      ctx.fill();
    });
  });
}

function drawAxisTriad(ctx, w, h) {
  const ox = 46, oy = h - 46, len = 26;
  const axes = [
    { v: [1, 0, 0], c: "#d11", label: "X" },
    { v: [0, 1, 0], c: "#1a1", label: "Y" },
    { v: [0, 0, 1], c: "#16c", label: "Z" },
  ];
  axes.forEach((a) => {
    const p = rotatePoint(a.v, view.yaw, view.pitch);
    const ex = ox + p[0] * len, ey = oy - p[1] * len;
    ctx.strokeStyle = a.c;
    ctx.lineWidth = 2;
    ctx.beginPath();
    ctx.moveTo(ox, oy);
    ctx.lineTo(ex, ey);
    ctx.stroke();
    ctx.fillStyle = a.c;
    ctx.font = "600 11px -apple-system, Segoe UI, Helvetica, Arial, sans-serif";
    ctx.fillText(a.label, ex + 3, ey + 3);
  });
}

function deformedScale() {
  if (view.dispScale !== null) return view.dispScale;
  if (!lastResult || !currentModel || lastResult.maxDisp <= 0) return 1;
  // Auto-fit so the largest displacement draws at ~18% of the cube edge, but
  // cap the exaggeration: a displacement that is genuinely negligible should
  // LOOK negligible rather than being magnified into an alarming shape.
  const fit = (currentModel.edgeSolver * 0.18) / lastResult.maxDisp;
  return Math.min(fit, 2000);
}

// ---- picking: click a node to select it as the load target -------------
function pickAt(clientX, clientY) {
  const canvas = document.getElementById("view3d");
  const rect = canvas.getBoundingClientRect();
  const mx = clientX - rect.left, my = clientY - rect.top;
  const M = currentModel;
  if (!M) return;
  const project = makeProjector(canvas, M);

  let bestNode = null, bestNodeDist = 16;
  M.nodes.forEach((n) => {
    const p = project(n.x, n.y, n.z);
    const d = Math.hypot(p.sx - mx, p.sy - my);
    if (d < bestNodeDist) { bestNodeDist = d; bestNode = n.id; }
  });
  if (bestNode !== null) {
    selectedNodeId = bestNode;
    selectedMemberId = null;
    const sel = document.getElementById("ld-node");
    if (sel) sel.value = String(bestNode);
    onSelectionChanged();
    return;
  }

  const nodeById = {};
  M.nodes.forEach((n) => { nodeById[n.id] = n; });
  let bestMember = null, bestMemberDist = 10;
  M.members.forEach((m) => {
    const a = project(nodeById[m.i].x, nodeById[m.i].y, nodeById[m.i].z);
    const b = project(nodeById[m.j].x, nodeById[m.j].y, nodeById[m.j].z);
    const d = pointSegmentDistance(mx, my, a.sx, a.sy, b.sx, b.sy);
    if (d < bestMemberDist) { bestMemberDist = d; bestMember = m.id; }
  });
  selectedMemberId = bestMember;
  if (bestMember !== null) selectedNodeId = null;
  onSelectionChanged();
}

function pointSegmentDistance(px, py, x1, y1, x2, y2) {
  const dx = x2 - x1, dy = y2 - y1;
  const len2 = dx * dx + dy * dy;
  if (len2 === 0) return Math.hypot(px - x1, py - y1);
  let t = ((px - x1) * dx + (py - y1) * dy) / len2;
  t = Math.max(0, Math.min(1, t));
  return Math.hypot(px - (x1 + t * dx), py - (y1 + t * dy));
}

function initViewerEvents() {
  const canvas = document.getElementById("view3d");

  canvas.addEventListener("mousedown", (e) => {
    view.lastX = e.clientX; view.lastY = e.clientY;
    if (e.button === 1 || e.shiftKey) { view.panning = true; e.preventDefault(); }
    else { view.dragging = true; view.moved = false; }
  });

  window.addEventListener("mousemove", (e) => {
    const dx = e.clientX - view.lastX, dy = e.clientY - view.lastY;
    if (view.dragging) {
      if (Math.abs(dx) > 2 || Math.abs(dy) > 2) view.moved = true;
      view.yaw += dx * 0.0115;
      view.pitch += dy * 0.0115;
      view.pitch = Math.max(-1.45, Math.min(1.45, view.pitch));
      view.lastX = e.clientX; view.lastY = e.clientY;
      drawScene();
    } else if (view.panning) {
      view.panX += dx; view.panY += dy;
      view.lastX = e.clientX; view.lastY = e.clientY;
      drawScene();
    }
  });

  window.addEventListener("mouseup", (e) => {
    if (view.dragging && !view.moved) pickAt(e.clientX, e.clientY);
    view.dragging = false;
    view.panning = false;
  });

  canvas.addEventListener("wheel", (e) => {
    e.preventDefault();
    view.zoom *= e.deltaY < 0 ? 1.12 : 1 / 1.12;
    view.zoom = Math.max(0.25, Math.min(6, view.zoom));
    drawScene();
  }, { passive: false });

  canvas.addEventListener("contextmenu", (e) => e.preventDefault());
  window.addEventListener("resize", drawScene);
}

function resetView() {
  view.yaw = -0.6; view.pitch = -0.35; view.zoom = 1.0;
  view.panX = 0; view.panY = 0;
  drawScene();
}
"""





class DataNotFoundError(Exception):
    """Raised when the Excel workbooks don't contain what the solver needs."""


def banner(step, title):
    print(f"\n[REV3] Phase {step} -- {title}")


# =============================================================================
# 1. UNIT DATABASE  -  units/*.xlsx is the source of truth
# =============================================================================
class UnitDatabase:
    def __init__(self, folder=os.path.join(HERE, "units")):
        self.folder = folder
        self.path = self._locate()
        self.systems = []
        self.quantity_units = {}
        self.factors = {}
        self.base_definitions = []
        self._load()

    def _locate(self):
        candidates = glob.glob(os.path.join(self.folder, "*.xlsx"))
        if not candidates:
            raise DataNotFoundError(f"No Excel workbook found in {self.folder!r}.")
        preferred = [c for c in candidates if "imperial" in c.lower() and "metric" in c.lower()]
        return preferred[0] if preferred else candidates[0]

    @staticmethod
    def _find_header_row(ws, first_cell, second_cell=None):
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
            v0 = row[0].value
            if isinstance(v0, str) and v0.strip() == first_cell:
                if second_cell is None:
                    return row[0].row
                v1 = row[1].value if len(row) > 1 else None
                if isinstance(v1, str) and v1.strip() == second_cell:
                    return row[0].row
        raise DataNotFoundError(
            f"Could not find a header row starting with {first_cell!r} in sheet {ws.title!r}.")

    def _load(self):
        wb = openpyxl.load_workbook(self.path, data_only=True)
        for required in ("Unit Systems", "Conversion Factors"):
            if required not in wb.sheetnames:
                raise DataNotFoundError(f"{self.path} is missing the {required!r} sheet.")

        ws = wb["Unit Systems"]
        header_row = self._find_header_row(ws, "Quantity")
        headers = [c.value.strip() if isinstance(c.value, str) else c.value for c in ws[header_row]]
        imp_col, met_col = 1, 2
        self.systems = [headers[imp_col], headers[met_col]]
        for row in ws.iter_rows(min_row=header_row + 1, max_row=ws.max_row):
            qty = row[0].value
            if not isinstance(qty, str) or not qty.strip():
                continue
            imp_unit, met_unit = row[imp_col].value, row[met_col].value
            self.quantity_units[qty.strip()] = {
                self.systems[0]: imp_unit.strip() if isinstance(imp_unit, str) else imp_unit,
                self.systems[1]: met_unit.strip() if isinstance(met_unit, str) else met_unit,
            }

        ws = wb["Conversion Factors"]
        header_row = self._find_header_row(ws, "Quantity", "Imperial unit")
        for row in ws.iter_rows(min_row=1, max_row=header_row - 1):
            desc, val, unit = row[0].value, row[1].value, row[2].value
            if isinstance(desc, str) and desc.strip() and isinstance(val, (int, float)):
                self.base_definitions.append((desc.strip(), float(val),
                                               unit.strip() if isinstance(unit, str) else unit))
        for row in ws.iter_rows(min_row=header_row + 1, max_row=ws.max_row):
            qty = row[0].value
            if not isinstance(qty, str) or not qty.strip():
                continue
            imp_unit, met_unit, imp_to_met, met_to_imp = row[1].value, row[2].value, row[3].value, row[4].value
            if not isinstance(imp_to_met, (int, float)):
                continue
            self.factors[(imp_unit, met_unit)] = {
                "quantity": qty.strip(),
                "imperial_to_metric": float(imp_to_met),
                "metric_to_imperial": float(met_to_imp) if isinstance(met_to_imp, (int, float))
                else 1.0 / float(imp_to_met),
            }
        if len(self.systems) != 2:
            raise DataNotFoundError(f"Expected exactly two unit systems in {self.path}, found {self.systems!r}.")

    def label(self, quantity, system):
        entry = self.quantity_units.get(quantity)
        if entry is None:
            raise DataNotFoundError(f"Quantity {quantity!r} not found in {self.path}.")
        for name, unit in entry.items():
            if name.lower().startswith(system.lower()) or system.lower().startswith(name.lower().split()[0].lower()):
                return unit
        raise DataNotFoundError(f"Unit system {system!r} not recognised for {quantity!r}.")

    def factor_pair(self, imperial_unit, metric_unit):
        key = (imperial_unit, metric_unit)
        if key not in self.factors:
            raise DataNotFoundError(f"No conversion factor for unit pair {key} in {self.path}.")
        return self.factors[key]

    def base_definition(self, keyword):
        for desc, val, _unit in self.base_definitions:
            if keyword.lower() in desc.lower():
                return val
        raise DataNotFoundError(f"No base definition containing {keyword!r} found in {self.path}.")


# =============================================================================
# 2. MATERIAL DATABASE  -  materials/*.xlsx is the source of truth
# =============================================================================
class MaterialDatabase:
    def __init__(self, folder=os.path.join(HERE, "materials")):
        self.folder = folder
        self.system = None
        self.path = None
        self.materials = {}
        self.units = {}

    def _locate(self, system):
        candidates = glob.glob(os.path.join(self.folder, "*.xlsx"))
        if not candidates:
            raise DataNotFoundError(f"No Excel workbook found in {self.folder!r}.")
        tag = "metric" if system.lower().startswith(("metric", "standard")) else "imperial"
        matches = [c for c in candidates if tag in os.path.basename(c).lower()]
        if not matches:
            raise DataNotFoundError(f"No materials workbook for unit system {system!r} in {self.folder!r}.")
        return matches[0]

    def load(self, system):
        self.system = system
        self.path = self._locate(system)
        wb = openpyxl.load_workbook(self.path, data_only=True)
        if "All Materials" not in wb.sheetnames:
            raise DataNotFoundError(f"{self.path} is missing the 'All Materials' sheet.")
        ws = wb["All Materials"]
        header_row = None
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
            if row[0].value == "Category" and row[1].value == "Label":
                header_row = row[0].row
                break
        if header_row is None:
            raise DataNotFoundError(f"No 'Category'/'Label' header row in {self.path}.")
        headers = [c.value for c in ws[header_row]]
        for h in headers:
            if isinstance(h, str):
                m = re.search(r"\[([^\]]+)\]", h)
                if m:
                    self.units[h] = m.group(1)
        self.materials = {}
        for row in ws.iter_rows(min_row=header_row + 1, max_row=ws.max_row):
            category, label = row[0].value, row[1].value
            if not isinstance(label, str) or not label.strip():
                continue
            if not isinstance(category, str) or not category.strip():
                continue
            props = {"Category": category}
            for h, cell in zip(headers[2:], row[2:]):
                props[h] = cell.value
            self.materials[label.strip()] = props
        if not self.materials:
            raise DataNotFoundError(f"No materials rows parsed from {self.path}.")
        return self

    def unit_of(self, header):
        return self.units.get(header, "")

    def value(self, material, base_header_prefix):
        for h, v in material.items():
            if h.startswith(base_header_prefix + " ") or h.startswith(base_header_prefix + " ["):
                return v, self.unit_of(h)
        for h, v in material.items():
            if h.startswith(base_header_prefix):
                return v, self.unit_of(h)
        return None, ""


# =============================================================================
# 3. MEMBER SIZE DATABASE  -  member_size/*.xlsx is the source of truth
# =============================================================================
class MemberSizeDatabase:
    def __init__(self, folder=os.path.join(HERE, "member_size"), sheet="Database v16.0"):
        self.folder = folder
        self.sheet_name = sheet
        self.path = self._locate()
        self.imperial_idx = {}
        self.metric_idx = {}
        self._rows = []
        self._load()

    def _locate(self):
        candidates = glob.glob(os.path.join(self.folder, "*.xlsx"))
        if not candidates:
            raise DataNotFoundError(f"No Excel workbook found in {self.folder!r}.")
        matches = [c for c in candidates if "shape" in os.path.basename(c).lower()]
        return matches[0] if matches else candidates[0]

    def _load(self):
        wb = openpyxl.load_workbook(self.path, data_only=True)
        if self.sheet_name not in wb.sheetnames:
            raise DataNotFoundError(f"{self.path} is missing the {self.sheet_name!r} sheet.")
        ws = wb[self.sheet_name]
        header = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
        if "EDI_Std_Nomenclature" not in header:
            raise DataNotFoundError(f"Could not find 'EDI_Std_Nomenclature' in {self.path}.")
        first = header.index("EDI_Std_Nomenclature")
        second = header.index("EDI_Std_Nomenclature", first + 1)
        self.imperial_idx = {name: i for i, name in enumerate(header[1:second], start=1)}
        self.metric_idx = {name: i for i, name in enumerate(header[second:], start=second)}
        self._rows = list(ws.iter_rows(min_row=2, values_only=True))
        if not self._rows:
            raise DataNotFoundError(f"No shape rows parsed from {self.path}.")

    def rows_of_type(self, shape_type="W"):
        return [r for r in self._rows if r[0] == shape_type]


# =============================================================================
# CONFIGURATION  (now just the STARTING selection shown when the page opens
# -- the browser can change all three of these live afterwards)
# =============================================================================
DEFAULT_UNITS = "metric"          # Philippines-based project -> Standard Metric
DEFAULT_MATERIAL_LABEL = "A36 Gr.36"
DEFAULT_MATERIAL_DECLARED_NAME = "ASTM A36 Steel"
DEFAULT_MEMBER_SIZE_LABEL = "W150X22.5"
EDGE_LENGTH_M = 6.0

DOF_PER_NODE = 6
_METRIC_SCALE = {"Ix": 1e6, "Iy": 1e6, "J": 1e3, "Sx": 1e3, "Sy": 1e3, "Zx": 1e3, "Zy": 1e3}
SECTION_KEYS = ("d", "bf", "tf", "tw", "A", "Ix", "Iy", "J", "Sx", "Sy")


def parse_args(argv=None):
    ap = argparse.ArgumentParser(
        description="REV3 cube solver -- launches an interactive, catalog-driven FE "
                    "solver in your browser; Units/Material/Member Size can all be "
                    "changed live on the page.")
    ap.add_argument("--units", choices=["metric", "imperial"], default=DEFAULT_UNITS,
                     help="STARTING unit system when the page opens (default: Standard "
                          "Metric). Changeable live in the browser afterwards.")
    ap.add_argument("--material", default=DEFAULT_MATERIAL_LABEL,
                     help=f"STARTING material library label (default: {DEFAULT_MATERIAL_LABEL!r}).")
    ap.add_argument("--member-size", dest="member_size", default=DEFAULT_MEMBER_SIZE_LABEL,
                     help=f"STARTING AISC W-shape label (default: {DEFAULT_MEMBER_SIZE_LABEL!r}).")
    ap.add_argument("--outdir", default=HERE, help="Folder to write the solver page into.")
    ap.add_argument("--no-browser", action="store_true",
                     help="Write the solver page but don't open a browser tab for it.")
    ap.add_argument("--source", choices=["auto", "excel", "embedded"], default="auto",
                     help="Where to read the unit/material/section catalog from. "
                          "'auto' (default) uses the Excel folders when they are beside "
                          "this script and the embedded snapshot otherwise; 'excel' "
                          "requires the workbooks; 'embedded' ignores them.")
    ap.add_argument("--embed-catalog", action="store_true",
                     help="Read the Excel workbooks and bake them into this .py file, "
                          "then exit. Afterwards REV3.py runs standalone with no Excel "
                          "files needed. Re-run this after editing the workbooks.")
    ap.add_argument("--list-materials", action="store_true",
                     help="Print every material in the catalog, then exit.")
    ap.add_argument("--list-sections", nargs="?", const="W", default=None, metavar="TYPE",
                     help="Print every AISC shape of the given type (default 'W'), then exit.")
    return ap.parse_args(argv)


def descale_section(raw_row_dict, is_metric):
    """Return a section property dict with Ix/Iy/J/Sx/Sy converted out of the
    AISC workbook's documented magnitude convention (metric only; imperial
    values are already plain in/in^2/in^3/in^4)."""
    out = dict(raw_row_dict)
    if is_metric:
        for key, scale in _METRIC_SCALE.items():
            if key in out and isinstance(out[key], (int, float)):
                out[key] = float(out[key]) * scale
    return out


# =============================================================================
# GEOMETRY  (dimensionless cube coefficients; the browser scales these into
# whichever unit system is currently selected)
# =============================================================================
NODE_COEFFS = {
    1: (0, 0, 0), 2: (1, 0, 0), 3: (1, 0, 1), 4: (0, 0, 1),
    5: (0, 1, 0), 6: (1, 1, 0), 7: (1, 1, 1), 8: (0, 1, 1),
}
MEMBERS_RAW = [
    (1, 2, "Base Beam"), (2, 3, "Base Beam"), (3, 4, "Base Beam"), (4, 1, "Base Beam"),
    (5, 6, "Roof Beam"), (6, 7, "Roof Beam"), (7, 8, "Roof Beam"), (8, 5, "Roof Beam"),
    (1, 5, "Column"), (2, 6, "Column"), (3, 7, "Column"), (4, 8, "Column"),
]
BETA_ANGLE = {"Base Beam": 0.0, "Roof Beam": 0.0, "Column": 90.0}
MEMBER_COLOR = {"Base Beam": "#4169e1", "Roof Beam": "#4169e1", "Column": "#2e8b57"}
SUPPORT_NODES = {1, 2, 3, 4}
PINNED_RESTRAINT = [1, 1, 1, 0, 0, 0]
FREE_RESTRAINT = [0, 0, 0, 0, 0, 0]


def node_restraint(node):
    return PINNED_RESTRAINT if node in SUPPORT_NODES else FREE_RESTRAINT


def local_axes_base(p_i, p_j):
    """Local axis triad from the DIMENSIONLESS coefficient coordinates --
    direction is scale-invariant, so this never needs to be redone when the
    browser switches unit systems."""
    p_i, p_j = np.array(p_i, dtype=float), np.array(p_j, dtype=float)
    local_x = p_j - p_i
    local_x /= np.linalg.norm(local_x)
    global_y = np.array([0.0, 1.0, 0.0])
    vertical = np.isclose(abs(np.dot(local_x, global_y)), 1.0)
    reference = np.array([0.0, 0.0, 1.0]) if vertical else global_y
    local_z = np.cross(local_x, reference)
    local_z /= np.linalg.norm(local_z)
    local_y = np.cross(local_z, local_x)
    local_y /= np.linalg.norm(local_y)
    return local_x, local_y, local_z


def build_static_geometry():
    """Everything about the cube that does NOT depend on the active unit
    system: node coefficients, member connectivity + local axes, DOF
    bookkeeping. Lengths are all "1 edge" in coefficient space; the browser
    multiplies by the current edge length in the current solver length unit."""
    nodes = []
    for n, coeffs in NODE_COEFFS.items():
        nodes.append({"id": n, "cx": coeffs[0], "cy": coeffs[1], "cz": coeffs[2],
                      "support": n in SUPPORT_NODES, "restraint": node_restraint(n)})
    members = []
    for idx, (i_node, j_node, mtype) in enumerate(MEMBERS_RAW, start=1):
        lx, ly, lz = local_axes_base(NODE_COEFFS[i_node], NODE_COEFFS[j_node])
        members.append({
            "id": idx, "i": i_node, "j": j_node, "type": mtype, "beta": BETA_ANGLE[mtype],
            "color": MEMBER_COLOR[mtype],
            "local_x": lx.tolist(), "local_y": ly.tolist(), "local_z": lz.tolist(),
        })
    total_dof = len(NODE_COEFFS) * DOF_PER_NODE
    restrained_dof = sum(sum(node_restraint(n)) for n in NODE_COEFFS)
    return {
        "nodes": nodes, "members": members,
        "dof": {"total": total_dof, "restrained": restrained_dof, "active": total_dof - restrained_dof},
    }


# =============================================================================
# CATALOG BUILDER  -  read EVERYTHING out of the Excel workbooks, both unit
# systems at once, so the browser can switch live without re-running Python.
# =============================================================================
def build_catalog(unit_db, size_db, args):
    systems = {}
    for short, system_name in (("metric", "Standard Metric"), ("imperial", "Imperial")):
        is_metric = short == "metric"
        mat_db = MaterialDatabase().load(system_name)

        materials = {}
        for label, props in mat_db.materials.items():
            e_val, e_unit = mat_db.value(props, "E")
            g_val, g_unit = mat_db.value(props, "G")
            fy_val, fy_unit = mat_db.value(props, "Yield")
            fu_val, fu_unit = mat_db.value(props, "Fu")
            dens_val, dens_unit = mat_db.value(props, "Density")
            # Only materials with the full set a frame solve needs are offered.
            if not all(isinstance(v, (int, float)) for v in (e_val, g_val, fy_val)):
                continue
            materials[label] = {
                "label": label, "category": props["Category"],
                "E": float(e_val), "G": float(g_val), "Fy": float(fy_val),
                "Fu": float(fu_val) if isinstance(fu_val, (int, float)) else None,
                "density": float(dens_val) if isinstance(dens_val, (int, float)) else None,
                "e_unit": e_unit, "g_unit": g_unit, "fy_unit": fy_unit,
                "fu_unit": fu_unit, "density_unit": dens_unit,
            }
        if not materials:
            raise DataNotFoundError(
                f"No usable materials (needing E, G and Fy) found for {system_name}.")

        idx = size_db.metric_idx if is_metric else size_db.imperial_idx
        sections = {}
        for row in size_db.rows_of_type("W"):
            label = row[idx["AISC_Manual_Label"]]
            if not isinstance(label, str):
                continue
            raw = {k: row[idx[k]] for k in SECTION_KEYS if k in idx}
            if not all(isinstance(raw.get(k), (int, float)) and raw[k] > 0 for k in SECTION_KEYS):
                continue  # skip any shape missing a property the solve needs
            sec = descale_section(raw, is_metric)
            sec["label"] = label
            sec["A_display"] = float(raw["A"])   # as tabulated, for the info panel
            sections[label] = sec
        if not sections:
            raise DataNotFoundError(f"No usable W-shapes found for {system_name}.")

        length_unit = unit_db.label("Node coordinates", system_name)
        section_unit = unit_db.label("Section dimensions", system_name)
        stress_unit = unit_db.label("Stress, modulus", system_name)

        # Edge length in this system's display length unit, converted from the
        # declared 6.000 m using the workbook's own factor.
        len_pair = unit_db.factor_pair(
            unit_db.label("Node coordinates", "Imperial"),
            unit_db.label("Node coordinates", "Standard Metric"))
        edge_length = EDGE_LENGTH_M if is_metric else EDGE_LENGTH_M * len_pair["metric_to_imperial"]

        if is_metric:
            # m -> mm is the SI-prefix definition of a millimetre.
            node_to_solver_length = 1000.0
            solver_length_unit = "mm"
            display_force_unit, display_moment_unit = "kN", "kN\u00b7m"
            force_display_scale = 1.0 / 1000.0     # N -> kN
            moment_display_scale = 1.0 / 1.0e6     # N mm -> kN m
            load_force_input_scale = 1000.0        # kN -> N
            load_moment_input_scale = 1.0e6        # kN m -> N mm
            density_to_solver = 1000.0 / 1.0e9     # kN/m^3 -> N/mm^3
        else:
            # ft -> in, read from the workbook's own "Base definitions" block.
            node_to_solver_length = unit_db.base_definition("Foot")
            solver_length_unit = "in"
            display_force_unit, display_moment_unit = "kip", "kip\u00b7in"
            force_display_scale = 1.0
            moment_display_scale = 1.0
            load_force_input_scale = 1.0
            load_moment_input_scale = 1.0
            density_to_solver = 1.0 / (node_to_solver_length ** 3)  # kip/ft^3 -> kip/in^3

        systems[short] = {
            "name": system_name,
            "materials": materials,
            "sections": sections,
            "section_labels": sorted(sections.keys()),
            "material_labels": sorted(materials.keys()),
            "length_unit": length_unit,
            "section_unit": section_unit,
            "stress_unit": stress_unit,
            "solver_length_unit": solver_length_unit,
            "edge_length": edge_length,
            "node_to_solver_length": node_to_solver_length,
            "display_force_unit": display_force_unit,
            "display_moment_unit": display_moment_unit,
            "force_display_scale": force_display_scale,
            "moment_display_scale": moment_display_scale,
            "load_force_input_scale": load_force_input_scale,
            "load_moment_input_scale": load_moment_input_scale,
            "density_to_solver": density_to_solver,
            "materials_source": os.path.basename(mat_db.path),
        }

    # Equivalent labels across systems, so switching units keeps the same
    # physical shape selected (W150X22.5 <-> W6X15) instead of resetting.
    section_equiv = {}
    for row in size_db.rows_of_type("W"):
        met = row[size_db.metric_idx["AISC_Manual_Label"]]
        imp = row[size_db.imperial_idx["AISC_Manual_Label"]]
        if isinstance(met, str) and isinstance(imp, str):
            section_equiv[met] = imp
            section_equiv[imp] = met

    return {
        "systems": systems,
        "section_equiv": section_equiv,
        "declared_material_name": DEFAULT_MATERIAL_DECLARED_NAME,
        "declared_material_label": DEFAULT_MATERIAL_LABEL,
        "edge_length_m": EDGE_LENGTH_M,
        # Cross-system load conversions, read from the workbook's own factor
        # table, so a load typed in kN keeps its physical magnitude when the
        # page is switched to Imperial (and back).
        "force_kN_to_kip": unit_db.factor_pair("kip", "kN")["metric_to_imperial"],
        "moment_kNm_to_kipin": unit_db.factor_pair("kip\u00b7in", "kN\u00b7m")["metric_to_imperial"],
        # Carried so the Phase-4 cross-system consistency checks can run even in
        # standalone (embedded) mode, where no workbook is available to re-read.
        "check_factors": {
            "ksi_to_MPa": unit_db.factor_pair("ksi", "MPa")["imperial_to_metric"],
            "in4_to_mm4": unit_db.factor_pair("in\u2074", "mm\u2074")["imperial_to_metric"],
        },
        "start": {"units": DEFAULT_UNITS, "material": DEFAULT_MATERIAL_LABEL,
                   "section": DEFAULT_MEMBER_SIZE_LABEL},
        "sources": {
            "units": os.path.basename(unit_db.path),
            "member_size": os.path.basename(size_db.path),
        },
    }


def apply_start_selection(catalog, args):
    """Resolve the opening Units/Material/Section shown when the page loads.
    Kept separate from build_catalog() so it applies identically whether the
    catalog came from the Excel workbooks or from the embedded snapshot."""
    start_units = args.units
    sys_c = catalog["systems"][start_units]

    start_material = args.material
    if start_material not in sys_c["materials"]:
        start_material = DEFAULT_MATERIAL_LABEL
    if start_material not in sys_c["materials"]:
        start_material = sys_c["material_labels"][0]

    start_section = args.member_size
    if start_section not in sys_c["sections"]:
        alt = catalog["section_equiv"].get(start_section)
        start_section = alt if alt in sys_c["sections"] else DEFAULT_MEMBER_SIZE_LABEL
    if start_section not in sys_c["sections"]:
        start_section = sys_c["section_labels"][0]

    catalog["start"] = {"units": start_units, "material": start_material,
                         "section": start_section}
    return catalog


# =============================================================================
# CATALOG SOURCING  -  Excel when available, embedded snapshot otherwise
# =============================================================================
def excel_folders_present():
    """True only if all three data folders exist AND each holds a workbook."""
    for folder in ("units", "materials", "member_size"):
        path = os.path.join(HERE, folder)
        if not os.path.isdir(path) or not glob.glob(os.path.join(path, "*.xlsx")):
            return False
    return True


def load_embedded_catalog():
    """Decode the catalog snapshot baked into this file. The snapshot is a
    gzip-compressed JSON blob produced by --embed-catalog from the very same
    build_catalog() code path that reads the workbooks, so the embedded data
    is identical to what Excel would have produced -- it is a cache, not a
    hand-maintained second copy that could drift."""
    if not EMBEDDED_CATALOG_B64.strip():
        raise DataNotFoundError(
            "No Excel data folders were found next to this script, and this copy "
            "has no embedded catalog baked in.\n"
            "        Either put the units/, materials/ and member_size/ folders "
            "beside REV3.py,\n"
            "        or run 'python REV3.py --embed-catalog' once while they ARE "
            "present to bake them in.")
    raw = zlib.decompress(base64.b64decode(EMBEDDED_CATALOG_B64))
    return json.loads(raw.decode("utf-8"))


def load_catalog(args):
    """Return (catalog, source_label). Prefers the Excel workbooks so the
    'Excel is the source of truth' architecture still holds when they're
    there; falls back to the embedded snapshot so the .py runs standalone."""
    want = args.source
    if want == "excel" or (want == "auto" and excel_folders_present()):
        if openpyxl is None:
            raise DataNotFoundError(
                "Reading the Excel workbooks needs openpyxl, which isn't installed.\n"
                "        Install it with 'pip install openpyxl', or run with "
                "--source embedded\n"
                "        to use the catalog baked into this file (no openpyxl needed).")
        unit_db = UnitDatabase()
        size_db = MemberSizeDatabase()
        catalog = build_catalog(unit_db, size_db, args)
        return apply_start_selection(catalog, args), "excel", unit_db
    catalog = load_embedded_catalog()
    return apply_start_selection(catalog, args), "embedded", None


def cmd_embed_catalog(args):
    """Rebuild this script's embedded snapshot from the current Excel files
    and rewrite REV3.py in place. Run this after editing the workbooks if you
    want the standalone copy to carry the new data."""
    if not excel_folders_present():
        print("[REV3] Cannot embed: the units/, materials/ and member_size/ folders "
              "must be present beside this script.", file=sys.stderr)
        return 2
    unit_db = UnitDatabase()
    size_db = MemberSizeDatabase()
    catalog = build_catalog(unit_db, size_db, args)
    blob = json.dumps(catalog, separators=(",", ":")).encode("utf-8")
    packed = base64.b64encode(zlib.compress(blob, 9)).decode("ascii")

    src_path = os.path.abspath(__file__)
    with open(src_path, encoding="utf-8") as f:
        src = f.read()

    start_marker = "EMBEDDED_CATALOG_B64 = ("
    start = src.index(start_marker)
    end = src.index("# --- end embedded catalog ---", start)

    wrapped = "\n".join('    "%s"' % packed[i:i + 96] for i in range(0, len(packed), 96))
    replacement = start_marker + "\n" + wrapped + "\n)\n"
    src = src[:start] + replacement + src[end:]

    with open(src_path, "w", encoding="utf-8") as f:
        f.write(src)

    print(f"[REV3] Embedded catalog refreshed from Excel into {os.path.basename(src_path)}")
    print(f"       raw JSON {len(blob)/1024:.0f} KB -> compressed+base64 {len(packed)/1024:.0f} KB")
    for key in ("metric", "imperial"):
        s = catalog["systems"][key]
        print(f"       {s['name']:<16} {len(s['materials'])} materials, {len(s['sections'])} W-shapes")
    print("       This script can now run standalone, with no Excel files beside it.")
    return 0


# =============================================================================
# HTML RENDERER
# =============================================================================
def render_html(catalog, geometry, out_path, page_template, solver_core, viewer_js):
    page = page_template.replace("__CATALOG_JSON__", json.dumps(catalog))
    page = page.replace("__GEOMETRY_JSON__", json.dumps(geometry))
    page = page.replace("__SOLVER_CORE__", solver_core)
    page = page.replace("__VIEWER_JS__", viewer_js)
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(page)
    return out_path


# =============================================================================
# PHASE 4 -- VALIDATION  (real checks, run every time)
# =============================================================================
def run_validation(catalog, geometry, source, out_path):
    checks = []

    def check(name, passed, detail=""):
        checks.append((name, bool(passed), detail))

    met = catalog["systems"]["metric"]
    imp = catalog["systems"]["imperial"]

    check("Catalog source resolved",
          source in ("excel", "embedded"),
          "read from the Excel workbooks" if source == "excel"
          else "read from the embedded snapshot (standalone mode)")
    check("Both unit systems present in the catalog",
          set(catalog["systems"]) == {"metric", "imperial"},
          f"{met['name']} + {imp['name']}")
    check("Default unit system is Standard Metric when --units is omitted",
          parse_args([]).units == "metric", "parse_args([]).units == 'metric'")
    check("Metric and Imperial length units are distinct",
          met["length_unit"] != imp["length_unit"],
          f"{met['length_unit']!r} vs {imp['length_unit']!r}")
    check("Materials loaded for BOTH unit systems",
          len(met["materials"]) > 0 and len(imp["materials"]) > 0,
          f"{len(met['materials'])} metric / {len(imp['materials'])} imperial")
    check("W-shape catalog loaded for BOTH unit systems",
          len(met["sections"]) > 0 and len(imp["sections"]) > 0,
          f"{len(met['sections'])} metric / {len(imp['sections'])} imperial")
    check("Declared material (A36) present in both systems",
          DEFAULT_MATERIAL_LABEL in met["materials"] and DEFAULT_MATERIAL_LABEL in imp["materials"],
          f"{DEFAULT_MATERIAL_LABEL!r}")

    cf = catalog.get("check_factors", {})
    a36_met = met["materials"].get(DEFAULT_MATERIAL_LABEL)
    a36_imp = imp["materials"].get(DEFAULT_MATERIAL_LABEL)
    if a36_met and a36_imp and "ksi_to_MPa" in cf:
        # E in ksi -> MPa should land on the metric figure (same physical steel)
        e_converted = a36_imp["E"] * cf["ksi_to_MPa"]
        rel = abs(e_converted - a36_met["E"]) / a36_met["E"]
        check("A36 E agrees across unit systems (cross-workbook consistency)",
              rel < 0.01, f"{a36_imp['E']} ksi -> {e_converted:.0f} MPa vs {a36_met['E']} MPa "
                          f"({rel*100:.2f}% diff)")

    # Section cross-check: metric Ix should equal imperial Ix converted.
    in4_to_mm4 = cf.get("in4_to_mm4")
    start_sec_met = met["sections"].get(DEFAULT_MEMBER_SIZE_LABEL)
    equiv_label = catalog["section_equiv"].get(DEFAULT_MEMBER_SIZE_LABEL)
    start_sec_imp = imp["sections"].get(equiv_label)
    if start_sec_met and start_sec_imp and in4_to_mm4:
        ix_conv = start_sec_imp["Ix"] * in4_to_mm4
        rel = abs(ix_conv - start_sec_met["Ix"]) / start_sec_met["Ix"]
        check("Section Ix agrees across unit systems (descaling is correct)",
              rel < 0.02,
              f"{equiv_label} {start_sec_imp['Ix']} in\u2074 -> {ix_conv:.3e} mm\u2074 vs "
              f"{DEFAULT_MEMBER_SIZE_LABEL} {start_sec_met['Ix']:.3e} mm\u2074 ({rel*100:.2f}% diff)")

    check("Every catalogued section has all solver-required properties",
          all(all(isinstance(s.get(k), (int, float)) and s[k] > 0 for k in SECTION_KEYS)
              for s in met["sections"].values()),
          f"{len(SECTION_KEYS)} properties x {len(met['sections'])} metric shapes")
    check("Cross-system section equivalence map is populated",
          catalog["section_equiv"].get(DEFAULT_MEMBER_SIZE_LABEL) is not None,
          f"{DEFAULT_MEMBER_SIZE_LABEL} <-> {catalog['section_equiv'].get(DEFAULT_MEMBER_SIZE_LABEL)}")
    check("Load-unit conversion factors sourced from Excel",
          catalog["force_kN_to_kip"] > 0 and catalog["moment_kNm_to_kipin"] > 0,
          f"1 kN = {catalog['force_kN_to_kip']:.6f} kip; "
          f"1 kN\u00b7m = {catalog['moment_kNm_to_kipin']:.6f} kip\u00b7in")
    check("DOF bookkeeping is internally consistent",
          geometry["dof"]["total"] == len(NODE_COEFFS) * DOF_PER_NODE
          and geometry["dof"]["active"] == geometry["dof"]["total"] - geometry["dof"]["restrained"],
          f"total={geometry['dof']['total']}, restrained={geometry['dof']['restrained']}, "
          f"active={geometry['dof']['active']}")
    check("Every member's local axis triad is orthonormal and right-handed",
          all(_axes_ok(m) for m in geometry["members"]),
          "R R^T = I and det(R) = +1 checked per member")
    check("Geometry has the expected 8 nodes / 12 members",
          len(geometry["nodes"]) == 8 and len(geometry["members"]) == 12,
          f"{len(geometry['nodes'])} nodes, {len(geometry['members'])} members")
    check("Solver page written and non-empty",
          os.path.isfile(out_path) and os.path.getsize(out_path) > 0,
          f"{os.path.basename(out_path)}, {os.path.getsize(out_path)/1024:.0f} KB")
    with open(out_path, encoding="utf-8") as f:
        page = f.read()
    check("Page has no unresolved template placeholders",
          "__CATALOG_JSON__" not in page and "__SOLVER_CORE__" not in page
          and "__VIEWER_JS__" not in page and "__GEOMETRY_JSON__" not in page)
    check("Page loads no external resources (works fully offline)",
          "http://" not in page and "https://" not in page,
          "no CDN <script>/<link> tags -- viewer is dependency-free Canvas2D")

    return all(c[1] for c in checks), checks


def _axes_ok(m, tol=1e-8):
    R = np.array([m["local_x"], m["local_y"], m["local_z"]])
    return np.max(np.abs(R @ R.T - np.eye(3))) < tol and abs(np.linalg.det(R) - 1.0) < tol


def print_validation_report(checks):
    width = max(len(c[0]) for c in checks) + 2
    for name, passed, detail in checks:
        print(f"  [{'PASS' if passed else 'FAIL'}] {name.ljust(width)} {detail}")


# =============================================================================
# DIAGNOSTIC MODES  (work from the catalog, so they run standalone too)
# =============================================================================
def cmd_list_materials(args):
    catalog, source, _ = load_catalog(args)
    sys_c = catalog["systems"][args.units]
    print(f"Materials for {sys_c['name']} ({len(sys_c['materials'])} total, "
          f"from the {source} catalog):\n")
    for label in sys_c["material_labels"]:
        m = sys_c["materials"][label]
        print(f"  {label:<20} {m['category']:<16} E={m['E']:<10g} Fy={m['Fy']:g} {m['fy_unit']}")


def cmd_list_sections(args, shape_type):
    catalog, source, _ = load_catalog(args)
    sys_c = catalog["systems"][args.units]
    if shape_type.upper() != "W":
        print(f"Only W-shapes are carried in the solver catalog "
              f"(asked for {shape_type!r}).", file=sys.stderr)
        return
    print(f"W-shapes for {sys_c['name']} ({len(sys_c['section_labels'])} total, "
          f"from the {source} catalog):\n")
    for label in sys_c["section_labels"]:
        print(f"  {label}")


# =============================================================================
# MAIN
# =============================================================================
def main(argv=None):
    args = parse_args(argv)

    try:
        if args.embed_catalog:
            return cmd_embed_catalog(args)
        if args.list_materials:
            cmd_list_materials(args)
            return 0
        if args.list_sections is not None:
            cmd_list_sections(args, args.list_sections)
            return 0

        banner(1, "INSPECT -- locate the unit / material / section data")
        if args.source == "embedded":
            print("  --source embedded : ignoring any Excel folders on purpose.")
        elif excel_folders_present():
            print(f"  Excel data folders found beside this script:")
            print(f"    units/        -> {os.path.basename(UnitDatabase().path)}")
            print(f"    member_size/  -> {os.path.basename(MemberSizeDatabase().path)}")
            print(f"    materials/    -> read per unit system in Phase 2")
        elif args.source == "excel":
            print("  --source excel given, but the units/ materials/ member_size/ "
                  "folders are not beside this script.")
            raise DataNotFoundError(
                "--source excel requires the units/, materials/ and member_size/ "
                "folders next to REV3.py.\n"
                "        Drop them beside the script, or use --source auto "
                "(the default) / --source embedded\n"
                "        to run from the catalog baked into this file.")
        else:
            print("  No Excel data folders beside this script.")
            print("  Falling back to the catalog embedded in REV3.py (standalone mode).")

        banner(2, "CATALOG -- every material and section, both unit systems")
        catalog, source, _unit_db = load_catalog(args)
        print(f"  Source: {'Excel workbooks' if source == 'excel' else 'embedded snapshot'}")
        for key in ("metric", "imperial"):
            sys_c = catalog["systems"][key]
            print(f"  {sys_c['name']:<16} {len(sys_c['materials']):>3} materials, "
                  f"{len(sys_c['sections']):>3} W-shapes  "
                  f"(edge {sys_c['edge_length']:.4g} {sys_c['length_unit']}, "
                  f"solver length {sys_c['solver_length_unit']}, "
                  f"loads in {sys_c['display_force_unit']})")
        start = catalog["start"]
        print(f"  Opening selection: {catalog['systems'][start['units']]['name']} / "
              f"{start['material']} / {start['section']}")
        print(f"  (all three are changeable live in the browser)")

        banner(3, "BUILD -- geometry and the interactive solver page")
        geometry = build_static_geometry()
        os.makedirs(args.outdir, exist_ok=True)
        out_path = os.path.join(args.outdir, "REV3_solver.html")
        render_html(catalog, geometry, out_path, PAGE_TEMPLATE, SOLVER_CORE_JS, VIEWER_JS)
        print(f"  Solver page: {out_path} ({os.path.getsize(out_path)/1024:.0f} KB, self-contained)")

        banner(4, "VALIDATE")
        all_passed, checks = run_validation(catalog, geometry, source, out_path)
        print_validation_report(checks)
        print()
        print("  ALL CHECKS PASSED" if all_passed else "  ONE OR MORE CHECKS FAILED -- see above")

        if not args.no_browser:
            print("\n[REV3] Opening the solver in your browser ...")
            webbrowser.open("file://" + os.path.abspath(out_path))
        else:
            print(f"\n[REV3] --no-browser given; open this file yourself:\n       {out_path}")

        return 0 if all_passed else 1

    except DataNotFoundError as e:
        print(f"\n[REV3] FATAL: {e}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
