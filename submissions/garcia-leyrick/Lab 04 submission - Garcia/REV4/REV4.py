"""
REV3.py  (Rev. 3 -- interactive solver, catalog-driven, load framework)

Excel -> Full Catalog -> Interactive Web Solver with Load Cases,
Combinations, Diaphragm, and Temperature Loads.

Rev 3 builds on the catalog-driven FE solver:
  * Load-case framework (9 cases: dead, live, wind, seismic, temperature)
  * NSCP LRFD + ASD load combinations
  * A roof diaphragm constraint DEFINITION (documented DOFs, not enforced)
  * Real temperature loads (thermal strain, not fake nodal forces)
  * A load viewer: pick a case, see its loads drawn on the cube
  * A Cartesian 3D view with labeled X/Y/Z axes and a ground grid
  * 11 automated tests (python REV3.py --test)

SCOPE NOTE:
  The browser FE solver (Solve button) solves whatever nodal loads are
  currently present -- hand-entered loads AND the nodal loads of the
  selected load case.  Distributed member loads and member point loads
  are DISPLAYED for the selected load case but are NOT fed into the FE
  solve (equivalent-nodal-load conversion is out of scope for this
  revision).  Temperature loads are shown and validated but not solved.

Excel is the source of truth for:
  units/*.xlsx          -> unit systems and conversion factors
  materials/*.xlsx       -> RISA material library (both unit systems), incl. G
                            and Therm. Coeff. (coefficient of thermal expansion)
  member_size/*.xlsx     -> AISC v16.0 shapes database (both unit systems)
"""

import argparse
import glob
import json
import math
import os
import re
import sys
import unittest
import webbrowser

import numpy as np
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))


# =============================================================================
# CONFIGURATION
# =============================================================================
DEFAULT_MATERIAL_LABEL = "A36 Gr.36"
DEFAULT_MATERIAL_DECLARED_NAME = "ASTM A36 Steel"
DEFAULT_MEMBER_SIZE_LABEL = "W150X22.5"
EDGE_LENGTH_M = 6.0

DOF_PER_NODE = 6
_METRIC_SCALE = {"Ix": 1e6, "Iy": 1e6, "J": 1e3, "Sx": 1e3, "Sy": 1e3, "Zx": 1e3, "Zy": 1e3}
SECTION_KEYS = ("d", "bf", "tf", "tw", "A", "Ix", "Iy", "J", "Sx", "Sy")

# Rev 3 roof-level load magnitudes
ROOF_DEAD_UDL_KN_PER_M = 5.0
ROOF_LIVE_UDL_KN_PER_M = 3.0
CENTER_POINT_LOAD_KN = 5.0
WIND_TOTAL_KN = 10.0
SEISMIC_TOTAL_KN = 15.0
NUM_LATERAL_NODES = 4
TEMPERATURE_CHANGE_DEGC = 15.0

ROOF_NODES = [5, 6, 7, 8]
ROOF_BEAMS = [5, 6, 7, 8]
DIAPHRAGM_MASTER_NODE = 5
DIAPHRAGM_DOF = ["UX", "UZ", "RY"]

# NSCP load combination factors (see verification report for edition caveat)
LRFD_COMBINATIONS = [
    ("1.4D",                    {"D": 1.4}),
    ("1.2D+1.6L",               {"D": 1.2, "L": 1.6}),
    ("1.2D+1.0W+1.0L",          {"D": 1.2, "W": 1.0, "L": 1.0}),
    ("1.2D+1.0E+1.0L",          {"D": 1.2, "E": 1.0, "L": 1.0}),
    ("0.9D+1.0W",               {"D": 0.9, "W": 1.0}),
    ("0.9D+1.0E",               {"D": 0.9, "E": 1.0}),
]
ASD_COMBINATIONS = [
    ("D",                       {"D": 1.0}),
    ("D+L",                     {"D": 1.0, "L": 1.0}),
    ("D+W",                     {"D": 1.0, "W": 1.0}),
    ("D+E",                     {"D": 1.0, "E": 1.0}),
    ("D+0.75L+0.75W",           {"D": 1.0, "L": 0.75, "W": 0.75}),
    ("D+0.75L+0.75E",           {"D": 1.0, "L": 0.75, "E": 0.75}),
    ("0.6D+W",                  {"D": 0.6, "W": 1.0}),
    ("0.6D+E",                  {"D": 0.6, "E": 1.0}),
]
LRFD_TEMP_COMBINATIONS = [
    ("1.2D+1.0T",               {"D": 1.2, "T": 1.0}),
    ("1.2D+1.0T+1.0L",          {"D": 1.2, "T": 1.0, "L": 1.0}),
]
ASD_TEMP_COMBINATIONS = [
    ("D+T",                     {"D": 1.0, "T": 1.0}),
    ("D+T+L",                   {"D": 1.0, "T": 1.0, "L": 1.0}),
]


# =============================================================================
# LOAD CASE DATA MODEL
# =============================================================================

class LoadCase:
    def __init__(self, case_id, name, category, description="",
                 self_weight_factor=0.0, loads=None):
        self.id = case_id
        self.name = name
        self.category = category
        self.description = description
        self.self_weight_factor = self_weight_factor
        self.loads = loads if loads is not None else []

    def total_force(self):
        fx = sum(getattr(l, "fx", 0.0) for l in self.loads)
        fy = sum(getattr(l, "fy", 0.0) for l in self.loads)
        fz = sum(getattr(l, "fz", 0.0) for l in self.loads)
        return fx, fy, fz

    def total_udl_force(self):
        total = 0.0
        for l in self.loads:
            if isinstance(l, MemberDistributedLoad):
                total += l.magnitude * l.length * l.direction_factor
        return total

    def __repr__(self):
        return (f"LoadCase(id={self.id}, name={self.name!r}, "
                f"category={self.category!r}, n_loads={len(self.loads)})")


class NodalLoad:
    def __init__(self, node_id, fx=0.0, fy=0.0, fz=0.0,
                 mx=0.0, my=0.0, mz=0.0, label=""):
        self.node_id = node_id
        self.fx, self.fy, self.fz = fx, fy, fz
        self.mx, self.my, self.mz = mx, my, mz
        self.label = label

    def __repr__(self):
        return f"NodalLoad(node={self.node_id}, F=({self.fx},{self.fy},{self.fz}))"


class MemberDistributedLoad:
    def __init__(self, member_id, direction, magnitude,
                 length=0.0, direction_factor=-1.0, label=""):
        self.member_id = member_id
        self.direction = direction
        self.magnitude = magnitude
        self.length = length
        self.direction_factor = direction_factor
        self.label = label

    def resultant(self):
        return self.magnitude * self.length * self.direction_factor

    def __repr__(self):
        return (f"MemberDistributedLoad(mbr={self.member_id}, "
                f"{self.direction}, {self.magnitude} kN/m)")


class MemberPointLoad:
    def __init__(self, member_id, location, direction, magnitude,
                 direction_factor=-1.0, label=""):
        self.member_id = member_id
        self.location = location
        self.direction = direction
        self.magnitude = magnitude
        self.direction_factor = direction_factor
        self.label = label

    def __repr__(self):
        return (f"MemberPointLoad(mbr={self.member_id}, loc={self.location}, "
                f"{self.direction}, {self.magnitude} kN)")


class Diaphragm:
    def __init__(self, dia_id, name, master_node, constrained_nodes,
                 degrees_of_freedom):
        self.id = dia_id
        self.name = name
        self.master_node = master_node
        self.constrained_nodes = list(constrained_nodes)
        self.degrees_of_freedom = list(degrees_of_freedom)

    def generate_constraint_equations(self):
        equations = []
        for n in self.constrained_nodes:
            if n == self.master_node:
                continue
            for dof in self.degrees_of_freedom:
                equations.append((n, self.master_node, dof))
        return equations

    def __repr__(self):
        return (f"Diaphragm(id={self.id}, master={self.master_node}, "
                f"nodes={self.constrained_nodes}, DOF={self.degrees_of_freedom})")


class TemperatureLoad:
    def __init__(self, member_ids, temperature_change, alpha,
                 reference_temperature=20.0, label=""):
        self.member_ids = list(member_ids)
        self.temperature_change = temperature_change
        self.alpha = alpha
        self.reference_temperature = reference_temperature
        self.label = label

    def thermal_strain(self):
        return self.alpha * self.temperature_change

    def free_expansion(self, length):
        return self.thermal_strain() * length

    def restrained_force(self, E, A):
        return E * A * self.thermal_strain()

    def __repr__(self):
        return (f"TemperatureLoad(mbrs={self.member_ids}, "
                f"dT={self.temperature_change} degC, alpha={self.alpha})")


class LoadCombination:
    def __init__(self, comb_id, name, design_method, factors):
        self.id = comb_id
        self.name = name
        self.design_method = design_method
        self.factors = dict(factors)

    def __repr__(self):
        return (f"LoadCombination(id={self.id}, name={self.name!r}, "
                f"method={self.design_method}, factors={self.factors})")


# =============================================================================
# EXCEL DATABASES
# =============================================================================

class DataNotFoundError(Exception):
    """Raised when the Excel workbooks don't contain what the solver needs."""


def banner(step, title):
    print(f"\n[REV3] Phase {step} -- {title}")


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

    def thermal_coefficient(self, material):
        """Coefficient of thermal expansion, 1/degC.

        The RISA workbook stores this as 11.7 in units of 1e-6/degC, so the
        stored value is scaled by 1e-6.  Falls back to structural steel if
        the column is missing.
        """
        for h, v in material.items():
            if isinstance(h, str) and "Therm" in h and isinstance(v, (int, float)):
                return float(v) * 1e-6
        return 11.7e-6

    def density(self, material):
        for h, v in material.items():
            if isinstance(h, str) and h.startswith("Density") and isinstance(v, (int, float)):
                return float(v)
        return None


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
# GEOMETRY  (dimensionless cube coefficients)
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
# LOAD CASE BUILDERS
# =============================================================================

def build_load_cases(geometry, material_db):
    edge = EDGE_LENGTH_M
    member_lengths = {}
    for m in geometry["members"]:
        ci = NODE_COEFFS[m["i"]]
        cj = NODE_COEFFS[m["j"]]
        member_lengths[m["id"]] = math.sqrt(
            sum((cj[k] - ci[k]) ** 2 for k in range(3))) * edge

    notes = []
    load_cases = []

    lc1 = LoadCase(1, "DEAD / SELF WEIGHT", "D",
                   "Self-weight of all members, from density x A x L, -Y.",
                   self_weight_factor=1.0)
    load_cases.append(lc1)

    lc2 = LoadCase(2, "ROOF DEAD", "D",
                   "5 kN/m uniformly distributed on roof beams, -Y.")
    for mbr_id in ROOF_BEAMS:
        lc2.loads.append(MemberDistributedLoad(
            member_id=mbr_id, direction="Y", magnitude=ROOF_DEAD_UDL_KN_PER_M,
            length=member_lengths[mbr_id], direction_factor=-1.0,
            label=f"Roof dead 5 kN/m on M{mbr_id}"))
    load_cases.append(lc2)

    lc3 = LoadCase(3, "ROOF LIVE", "L",
                   "3 kN/m uniformly distributed on roof beams, -Y.")
    for mbr_id in ROOF_BEAMS:
        lc3.loads.append(MemberDistributedLoad(
            member_id=mbr_id, direction="Y", magnitude=ROOF_LIVE_UDL_KN_PER_M,
            length=member_lengths[mbr_id], direction_factor=-1.0,
            label=f"Roof live 3 kN/m on M{mbr_id}"))
    load_cases.append(lc3)

    lc4 = LoadCase(4, "ROOF BEAM CENTER LOAD", "L",
                   "5 kN point load at the center of each roof beam, -Y.")
    for mbr_id in ROOF_BEAMS:
        lc4.loads.append(MemberPointLoad(
            member_id=mbr_id, location=0.5, direction="Y",
            magnitude=CENTER_POINT_LOAD_KN, direction_factor=-1.0,
            label=f"5 kN at center of M{mbr_id}"))
    load_cases.append(lc4)

    per_node_wind = WIND_TOTAL_KN / NUM_LATERAL_NODES
    lc5 = LoadCase(5, "WIND X", "W",
                   "10 kN total lateral load in +X, equally to 4 roof nodes.")
    for node_id in ROOF_NODES:
        lc5.loads.append(NodalLoad(node_id=node_id, fx=per_node_wind,
                                   label=f"Wind X at node {node_id}"))
    load_cases.append(lc5)

    lc6 = LoadCase(6, "WIND Z", "W",
                   "10 kN total lateral load in +Z, equally to 4 roof nodes.")
    for node_id in ROOF_NODES:
        lc6.loads.append(NodalLoad(node_id=node_id, fz=per_node_wind,
                                   label=f"Wind Z at node {node_id}"))
    load_cases.append(lc6)

    per_node_seis = SEISMIC_TOTAL_KN / NUM_LATERAL_NODES
    lc7 = LoadCase(7, "SEISMIC X", "E",
                   "15 kN total seismic load in +X, equally to 4 roof nodes.")
    for node_id in ROOF_NODES:
        lc7.loads.append(NodalLoad(node_id=node_id, fx=per_node_seis,
                                   label=f"Seismic X at node {node_id}"))
    load_cases.append(lc7)

    lc8 = LoadCase(8, "SEISMIC Z", "E",
                   "15 kN total seismic load in +Z, equally to 4 roof nodes.")
    for node_id in ROOF_NODES:
        lc8.loads.append(NodalLoad(node_id=node_id, fz=per_node_seis,
                                   label=f"Seismic Z at node {node_id}"))
    load_cases.append(lc8)

    a36 = material_db.materials.get(DEFAULT_MATERIAL_LABEL)
    if a36 is not None:
        alpha = material_db.thermal_coefficient(a36)
    else:
        alpha = 11.7e-6
        notes.append("A36 not found; using default steel alpha.")

    lc9 = LoadCase(9, "TEMPERATURE +15 degC", "T",
                   "+15 degC uniform temperature change on all members.")
    all_member_ids = [m["id"] for m in geometry["members"]]
    temp_load = TemperatureLoad(
        member_ids=all_member_ids,
        temperature_change=TEMPERATURE_CHANGE_DEGC,
        alpha=alpha,
        reference_temperature=20.0,
        label="+15 degC on all members")
    lc9.loads.append(temp_load)
    load_cases.append(lc9)

    diaphragm = Diaphragm(
        dia_id=1, name="Roof Diaphragm (Y=6m)",
        master_node=DIAPHRAGM_MASTER_NODE,
        constrained_nodes=ROOF_NODES,
        degrees_of_freedom=DIAPHRAGM_DOF)

    return load_cases, diaphragm, [temp_load], notes


def compute_self_weight(geometry, material_db, section_db):
    density = None
    a36 = material_db.materials.get(DEFAULT_MATERIAL_LABEL)
    if a36 is not None:
        density = material_db.density(a36)
    if density is None:
        density = 77.0

    idx = section_db.metric_idx
    area_m2 = None
    for row in section_db.rows_of_type("W"):
        label = row[idx["AISC_Manual_Label"]]
        if label == DEFAULT_MEMBER_SIZE_LABEL:
            area_m2 = float(row[idx["A"]]) * 1e-6
            break
    if area_m2 is None:
        area_m2 = 2.87e-3

    total_weight = 0.0
    for m in geometry["members"]:
        ci = NODE_COEFFS[m["i"]]
        cj = NODE_COEFFS[m["j"]]
        length = math.sqrt(sum((cj[k] - ci[k]) ** 2 for k in range(3))) * EDGE_LENGTH_M
        total_weight += density * area_m2 * length

    return total_weight, density, area_m2


def build_combinations(load_cases):
    by_category = {}
    for lc in load_cases:
        by_category.setdefault(lc.category, []).append(lc.id)

    combinations = []
    cid = 1
    for name, factors in LRFD_COMBINATIONS:
        combinations.append(LoadCombination(cid, f"LRFD: {name}", "LRFD", factors)); cid += 1
    for name, factors in ASD_COMBINATIONS:
        combinations.append(LoadCombination(cid, f"ASD: {name}", "ASD", factors)); cid += 1
    for name, factors in LRFD_TEMP_COMBINATIONS:
        combinations.append(LoadCombination(cid, f"LRFD+T: {name}", "LRFD", factors)); cid += 1
    for name, factors in ASD_TEMP_COMBINATIONS:
        combinations.append(LoadCombination(cid, f"ASD+T: {name}", "ASD", factors)); cid += 1
    return combinations, by_category


class LoadValidationReport:
    def __init__(self):
        self.case_summaries = []
        self.combination_summaries = []
        self.diaphragm_summary = None
        self.temperature_summary = None

    def to_text(self):
        L = []
        L.append("=" * 78)
        L.append("REV 3 LOAD VALIDATION REPORT")
        L.append("=" * 78)
        L.append("")
        L.append("NOTE: This report validates LOAD APPLICATION and LOAD TOTALS.")
        L.append("      The FE solver is separate; this report does not re-run it.")
        L.append("      Distributed and member point loads are DISPLAYED but not")
        L.append("      converted to equivalent nodal loads in this revision.")
        L.append("")
        L.append("-" * 78); L.append("LOAD CASE SUMMARY"); L.append("-" * 78)
        for s in self.case_summaries:
            L.append("")
            L.append(f"  Load Case {s['id']} / {s['name']}")
            L.append(f"    Category: {s['category']}")
            L.append(f"    Description: {s['description']}")
            if s.get("nodal_total") is not None:
                L.append(f"    Nodal total: Fx={s['nodal_total'][0]:.4f}, "
                         f"Fy={s['nodal_total'][1]:.4f}, Fz={s['nodal_total'][2]:.4f} kN")
            if s.get("udl_total") is not None:
                L.append(f"    Distributed total: {s['udl_total']:.4f} kN")
            if s.get("self_weight_total") is not None:
                L.append(f"    Self-weight total: {s['self_weight_total']:.4f} kN")
            if s.get("temperature") is not None:
                t = s["temperature"]
                L.append(f"    Temperature change: {t['dT']:.2f} degC")
                L.append(f"    alpha: {t['alpha']:.4e} 1/degC")
                L.append(f"    Thermal strain: {t['strain']:.6e}")
                L.append(f"    Free expansion (6 m): {t['free_expansion_mm']:.4f} mm")
                L.append(f"    Fully restrained force: {t['restrained_force_kN']:.4f} kN")
                L.append(f"    Net force: {t['net_force']:.4f} kN (self-straining)")
            L.append(f"    Number of loads: {s['n_loads']}")

        L.append(""); L.append("-" * 78)
        L.append("LOAD COMBINATION SUMMARY"); L.append("-" * 78)
        for s in self.combination_summaries:
            L.append(f"  {s['name']:<24} {s['design_method']:<5} "
                     f"factors={s['factors']}  n_ref={s['n_loads']}")

        L.append(""); L.append("-" * 78)
        L.append("DIAPHRAGM DEFINITION"); L.append("-" * 78)
        if self.diaphragm_summary:
            d = self.diaphragm_summary
            L.append(f"  Name: {d['name']}")
            L.append(f"  Master node: N{d['master_node']}")
            L.append(f"  Constrained nodes: {d['constrained_nodes']}")
            L.append(f"  Constrained DOF: {d['degrees_of_freedom']}")
            L.append(f"  Free DOF: {d['free_dof']}")
            L.append(f"  Constraint equations: {d['n_equations']}")
            L.append("  NOTE: Defined and tested; NOT eliminated into the equation")
            L.append("        set (out of scope for this revision).")

        L.append(""); L.append("-" * 78)
        L.append("TEMPERATURE LOAD VERIFICATION"); L.append("-" * 78)
        if self.temperature_summary:
            t = self.temperature_summary
            L.append(f"  Load case: {t['name']}")
            L.append(f"  dT: +{t['dT']:.1f} degC   Material: {t['material']}")
            L.append(f"  alpha: {t['alpha']:.4e} 1/degC")
            L.append(f"  eps_T = alpha*dT: {t['strain']:.6e}")
            L.append(f"  L: {t['length']:.4f} m   dL = {t['free_expansion']*1000:.4f} mm")
            L.append(f"  EA: {t['EA']:.1f} kN   N_restrained = {t['restrained_force']:.4f} kN")
            L.append(f"  Net force: {t['net_force']:.4f} kN (self-equilibrating)")
            L.append("  Partially restrained case: INDETERMINATE (no solve on this load)")
        L.append(""); L.append("=" * 78); L.append("END OF REPORT"); L.append("=" * 78)
        return "\n".join(L)


def validate_load_cases(load_cases, geometry, material_db, section_db,
                        diaphragm, self_weight_total, density, area_m2):
    report = LoadValidationReport()
    for lc in load_cases:
        s = {"id": lc.id, "name": lc.name, "category": lc.category,
             "description": lc.description, "n_loads": len(lc.loads),
             "nodal_total": None, "udl_total": None,
             "self_weight_total": None, "temperature": None}
        if lc.id == 1:
            s["self_weight_total"] = self_weight_total
        elif lc.category == "T":
            t = lc.loads[0]
            strain = t.thermal_strain()
            free_exp = t.free_expansion(EDGE_LENGTH_M)
            E = 200e6
            EA = E * area_m2
            N = t.restrained_force(E, area_m2)
            s["temperature"] = {"dT": t.temperature_change, "alpha": t.alpha,
                                "strain": strain, "free_expansion_mm": free_exp * 1000,
                                "restrained_force_kN": N, "net_force": 0.0}
            report.temperature_summary = {
                "name": lc.name, "dT": t.temperature_change,
                "material": DEFAULT_MATERIAL_LABEL, "alpha": t.alpha,
                "strain": strain, "length": EDGE_LENGTH_M,
                "free_expansion": free_exp, "EA": EA,
                "restrained_force": N, "net_force": 0.0}
        else:
            fx, fy, fz = lc.total_force()
            s["nodal_total"] = (fx, fy, fz)
            udl = lc.total_udl_force()
            if udl != 0.0:
                s["udl_total"] = udl
        report.case_summaries.append(s)

    eqs = diaphragm.generate_constraint_equations()
    free_dof = [d for d in ["UX", "UY", "UZ", "RX", "RY", "RZ"]
                if d not in diaphragm.degrees_of_freedom]
    report.diaphragm_summary = {
        "name": diaphragm.name, "master_node": diaphragm.master_node,
        "constrained_nodes": diaphragm.constrained_nodes,
        "degrees_of_freedom": diaphragm.degrees_of_freedom,
        "free_dof": free_dof, "n_equations": len(eqs)}

    combos, _ = build_combinations(load_cases)
    for c in combos:
        report.combination_summaries.append({
            "name": c.name, "design_method": c.design_method,
            "factors": c.factors, "n_loads": len(c.factors)})
    return report


# =============================================================================
# AUTOMATED TESTS  (Rev 3 load framework)
# =============================================================================

class Rev3LoadTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.size_db = MemberSizeDatabase()
        cls.mat_db = MaterialDatabase().load("Standard Metric")
        cls.geometry = build_static_geometry()
        cls.load_cases, cls.diaphragm, cls.temp_loads, cls.notes = \
            build_load_cases(cls.geometry, cls.mat_db)
        cls.combinations, cls.by_category = build_combinations(cls.load_cases)

    def test_01_self_weight(self):
        lc1 = next(c for c in self.load_cases if c.id == 1)
        self.assertEqual(lc1.category, "D")
        self.assertEqual(lc1.self_weight_factor, 1.0)
        total, density, area = compute_self_weight(self.geometry, self.mat_db, self.size_db)
        self.assertGreater(total, 0.0)

    def test_02_roof_dead(self):
        lc2 = next(c for c in self.load_cases if c.id == 2)
        udls = [l for l in lc2.loads if isinstance(l, MemberDistributedLoad)]
        self.assertEqual(len(udls), 4)
        for l in udls:
            self.assertAlmostEqual(l.magnitude, 5.0, places=6)
            self.assertEqual(l.direction, "Y")
        self.assertAlmostEqual(lc2.total_udl_force(), -120.0, places=3)

    def test_03_roof_live(self):
        lc3 = next(c for c in self.load_cases if c.id == 3)
        self.assertAlmostEqual(lc3.total_udl_force(), -72.0, places=3)

    def test_04_center_point(self):
        lc4 = next(c for c in self.load_cases if c.id == 4)
        pts = [l for l in lc4.loads if isinstance(l, MemberPointLoad)]
        self.assertEqual(len(pts), 4)
        for l in pts:
            self.assertAlmostEqual(l.location, 0.5, places=6)
            self.assertAlmostEqual(l.magnitude, 5.0, places=6)

    def test_05_wind_x(self):
        lc5 = next(c for c in self.load_cases if c.id == 5)
        nodal = [l for l in lc5.loads if isinstance(l, NodalLoad)]
        self.assertAlmostEqual(sum(l.fx for l in nodal), 10.0, places=6)

    def test_06_wind_z(self):
        lc6 = next(c for c in self.load_cases if c.id == 6)
        nodal = [l for l in lc6.loads if isinstance(l, NodalLoad)]
        self.assertAlmostEqual(sum(l.fz for l in nodal), 10.0, places=6)

    def test_07_seismic_x(self):
        lc7 = next(c for c in self.load_cases if c.id == 7)
        nodal = [l for l in lc7.loads if isinstance(l, NodalLoad)]
        self.assertAlmostEqual(sum(l.fx for l in nodal), 15.0, places=6)

    def test_08_seismic_z(self):
        lc8 = next(c for c in self.load_cases if c.id == 8)
        nodal = [l for l in lc8.loads if isinstance(l, NodalLoad)]
        self.assertAlmostEqual(sum(l.fz for l in nodal), 15.0, places=6)

    def test_09_diaphragm(self):
        d = self.diaphragm
        self.assertEqual(d.master_node, 5)
        self.assertEqual(sorted(d.constrained_nodes), [5, 6, 7, 8])
        self.assertEqual(len(d.generate_constraint_equations()), 9)

    def test_10_combinations(self):
        self.assertGreater(len(self.combinations), 0)
        c14d = next(c for c in self.combinations if "1.4D" in c.name and "T" not in c.name)
        self.assertEqual(c14d.factors.get("D"), 1.4)

    def test_11_temperature(self):
        lc9 = next(c for c in self.load_cases if c.id == 9)
        t = lc9.loads[0]
        self.assertAlmostEqual(t.temperature_change, 15.0, places=6)
        self.assertAlmostEqual(t.alpha, 11.7e-6, places=12)
        self.assertAlmostEqual(t.thermal_strain(), 1.755e-4, places=8)
        self.assertFalse(hasattr(t, "fx"))


def run_load_tests(verbose=False):
    loader = unittest.TestLoader()
    suite = loader.loadTestsFromTestCase(Rev3LoadTests)
    runner = unittest.TextTestRunner(verbosity=2 if verbose else 1)
    result = runner.run(suite)
    return 0 if result.wasSuccessful() else 1


# =============================================================================
# CATALOG BUILDER
# =============================================================================

def parse_args(argv=None):
    ap = argparse.ArgumentParser(
        description="REV3 cube solver -- catalog-driven FE solver with a Rev 3 "
                    "load framework (cases, combinations, diaphragm, temperature).")
    ap.add_argument("--units", choices=["metric", "imperial"], default="metric")
    ap.add_argument("--material", default=DEFAULT_MATERIAL_LABEL)
    ap.add_argument("--member-size", dest="member_size", default=DEFAULT_MEMBER_SIZE_LABEL)
    ap.add_argument("--outdir", default=HERE)
    ap.add_argument("--no-browser", action="store_true")
    ap.add_argument("--verify-loads", action="store_true",
                     help="Write the load-validation report headlessly, then exit.")
    ap.add_argument("--test", action="store_true",
                     help="Run the load-framework test suite, then exit.")
    ap.add_argument("--verbose", action="store_true")
    ap.add_argument("--list-materials", action="store_true")
    ap.add_argument("--list-sections", nargs="?", const="W", default=None, metavar="TYPE")
    return ap.parse_args(argv)


def descale_section(raw_row_dict, is_metric):
    out = dict(raw_row_dict)
    if is_metric:
        for key, scale in _METRIC_SCALE.items():
            if key in out and isinstance(out[key], (int, float)):
                out[key] = float(out[key]) * scale
    return out


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
            alpha = mat_db.thermal_coefficient(props)
            if not all(isinstance(v, (int, float)) for v in (e_val, g_val, fy_val)):
                continue
            materials[label] = {
                "label": label, "category": props["Category"],
                "E": float(e_val), "G": float(g_val), "Fy": float(fy_val),
                "Fu": float(fu_val) if isinstance(fu_val, (int, float)) else None,
                "density": float(dens_val) if isinstance(dens_val, (int, float)) else None,
                "alpha": float(alpha),
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
                continue
            sec = descale_section(raw, is_metric)
            sec["label"] = label
            sec["A_display"] = float(raw["A"])
            sections[label] = sec
        if not sections:
            raise DataNotFoundError(f"No usable W-shapes found for {system_name}.")

        length_unit = unit_db.label("Node coordinates", system_name)
        section_unit = unit_db.label("Section dimensions", system_name)
        stress_unit = unit_db.label("Stress, modulus", system_name)

        len_pair = unit_db.factor_pair(
            unit_db.label("Node coordinates", "Imperial"),
            unit_db.label("Node coordinates", "Standard Metric"))
        edge_length = EDGE_LENGTH_M if is_metric else EDGE_LENGTH_M * len_pair["metric_to_imperial"]

        if is_metric:
            node_to_solver_length = 1000.0
            solver_length_unit = "mm"
            display_force_unit, display_moment_unit = "kN", "kN\u00b7m"
            force_display_scale = 1.0 / 1000.0
            moment_display_scale = 1.0 / 1.0e6
            load_force_input_scale = 1000.0
            load_moment_input_scale = 1.0e6
            density_to_solver = 1000.0 / 1.0e9
        else:
            node_to_solver_length = unit_db.base_definition("Foot")
            solver_length_unit = "in"
            display_force_unit, display_moment_unit = "kip", "kip\u00b7in"
            force_display_scale = 1.0
            moment_display_scale = 1.0
            load_force_input_scale = 1.0
            load_moment_input_scale = 1.0
            density_to_solver = 1.0 / (node_to_solver_length ** 3)

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

    section_equiv = {}
    for row in size_db.rows_of_type("W"):
        met = row[size_db.metric_idx["AISC_Manual_Label"]]
        imp = row[size_db.imperial_idx["AISC_Manual_Label"]]
        if isinstance(met, str) and isinstance(imp, str):
            section_equiv[met] = imp
            section_equiv[imp] = met

    start_units = args.units
    start_sys = systems[start_units]
    start_material = args.material if args.material in start_sys["materials"] else DEFAULT_MATERIAL_LABEL
    if start_material not in start_sys["materials"]:
        start_material = start_sys["material_labels"][0]
    start_section = args.member_size
    if start_section not in start_sys["sections"]:
        alt = section_equiv.get(start_section)
        start_section = alt if alt in start_sys["sections"] else DEFAULT_MEMBER_SIZE_LABEL
    if start_section not in start_sys["sections"]:
        start_section = start_sys["section_labels"][0]

    return {
        "systems": systems,
        "section_equiv": section_equiv,
        "declared_material_name": DEFAULT_MATERIAL_DECLARED_NAME,
        "declared_material_label": DEFAULT_MATERIAL_LABEL,
        "edge_length_m": EDGE_LENGTH_M,
        "force_kN_to_kip": unit_db.factor_pair("kip", "kN")["metric_to_imperial"],
        "moment_kNm_to_kipin": unit_db.factor_pair("kip\u00b7in", "kN\u00b7m")["metric_to_imperial"],
        "start": {"units": start_units, "material": start_material, "section": start_section},
        "sources": {
            "units": os.path.basename(unit_db.path),
            "member_size": os.path.basename(size_db.path),
        },
    }


# =============================================================================
# HTML PAGE  (the web app)
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
    --accent:#1f6feb; --accent-dark:#123a75; --good:#2d6a4f;
    --warn:#b45309; --bad:#b42318; --temp:#b45309;
  }
  * { box-sizing:border-box; }
  html, body { height:100%; margin:0; }
  body {
    font-family:-apple-system,"Segoe UI",Helvetica,Arial,sans-serif;
    color:var(--ink); background:#fff; display:flex; flex-direction:column;
    font-size:13px; overflow:hidden;
  }
  .ribbon { border-bottom:1px solid var(--line); background:#fff; flex-shrink:0; }
  .ribbon-top {
    display:flex; align-items:center; gap:10px; padding:8px 16px 6px;
    flex-wrap:wrap;
  }
  .ribbon-top h1 { font-size:1.02rem; margin:0; white-space:nowrap; }
  .ribbon-top .sep { width:1px; height:22px; background:var(--line); }
  .ctl { display:flex; align-items:center; gap:6px; }
  .ctl label {
    font-size:0.72rem; color:var(--sub); text-transform:uppercase;
    letter-spacing:.03em;
  }
  .ctl select {
    padding:5px 7px; border:1px solid var(--line); border-radius:6px;
    font-size:0.82rem; background:#fff; max-width:260px;
  }
  .ribbon-status {
    display:flex; align-items:center; gap:14px; padding:5px 16px;
    background:var(--panel); border-top:1px solid var(--line);
    font-size:0.76rem; color:var(--sub); flex-wrap:wrap;
  }
  .ribbon-status b { color:var(--ink); }
  .pill { padding:2px 9px; border-radius:99px; font-size:0.72rem; font-weight:700; }
  .pill-ok { background:#e7f5ec; color:var(--good); }
  .pill-warn { background:#fdf3e4; color:var(--warn); }
  .pill-bad { background:#fdecea; color:var(--bad); }
  .pill-idle { background:#eceff2; color:var(--sub); }
  .layout { flex:1; display:flex; min-height:0; }
  .viewer-col {
    flex:1 1 60%; position:relative; border-right:1px solid var(--line);
    min-width:0; background:#fbfcfd;
  }
  #view3d { width:100%; height:100%; display:block; cursor:grab; }
  #view3d:active { cursor:grabbing; }
  .view-controls {
    position:absolute; top:10px; left:10px; display:flex; gap:6px;
    flex-wrap:wrap;
  }
  .view-controls button {
    font-size:0.76rem; padding:5px 10px; border:1px solid var(--line);
    background:#fff; border-radius:6px; cursor:pointer;
  }
  .view-controls button.active {
    background:var(--accent); color:#fff; border-color:var(--accent);
  }
  .view-hint {
    position:absolute; bottom:10px; left:10px; font-size:0.72rem;
    color:var(--sub); background:rgba(255,255,255,.9);
    border:1px solid var(--line); border-radius:7px; padding:6px 10px;
  }
  .view-legend {
    position:absolute; bottom:10px; right:10px; background:rgba(255,255,255,.95);
    border:1px solid var(--line); border-radius:8px; padding:8px 12px;
    font-size:0.72rem; max-width:280px;
  }
  .view-legend h4 { margin:0 0 5px 0; font-size:0.78rem; }
  .view-legend .row { display:flex; align-items:center; gap:6px; margin-bottom:3px; }
  .view-legend .swatch { width:14px; height:3px; border-radius:2px; }
  .side-col {
    flex:1 1 40%; display:flex; flex-direction:column; min-width:370px;
    max-width:640px;
  }
  .tabs {
    display:flex; border-bottom:1px solid var(--line); background:var(--panel);
    flex-shrink:0;
  }
  .tabs button {
    flex:1; padding:9px 4px; border:none; background:none; cursor:pointer;
    font-size:0.8rem; font-weight:600; color:var(--sub);
    border-bottom:2px solid transparent;
  }
  .tabs button.active {
    color:var(--accent); border-bottom-color:var(--accent); background:#fff;
  }
  .panel { flex:1; overflow-y:auto; padding:13px 15px; display:none; }
  .panel.active { display:block; }
  .card {
    border:1px solid var(--line); border-radius:8px; padding:11px 13px;
    margin-bottom:11px;
  }
  .card h3 { margin:0 0 8px 0; font-size:0.85rem; }
  .kv { display:grid; grid-template-columns:1fr auto; gap:3px 10px; font-size:0.78rem; }
  .kv .k { color:var(--sub); }
  .kv .v { font-weight:600; text-align:right; }
  .grid3 { display:grid; grid-template-columns:repeat(3,1fr); gap:7px; }
  .fld label { display:block; font-size:0.7rem; color:var(--sub); margin-bottom:2px; }
  .fld input, .fld select {
    width:100%; padding:5px 6px; border:1px solid var(--line);
    border-radius:5px; font-size:0.8rem;
  }
  button.primary {
    background:var(--accent); color:#fff; border:none; border-radius:6px;
    padding:9px 14px; font-size:0.85rem; font-weight:700; cursor:pointer;
    width:100%;
  }
  button.primary:hover { background:var(--accent-dark); }
  button.ghost {
    background:#fff; border:1px solid var(--line); border-radius:6px;
    padding:6px 10px; font-size:0.77rem; cursor:pointer; color:var(--sub);
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
    display:flex; justify-content:space-between; align-items:center; gap:8px;
    padding:6px 9px; border:1px solid var(--line); border-radius:6px;
    margin-bottom:5px; font-size:0.75rem; background:var(--panel);
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
  .combo-row { display:flex; justify-content:space-between; padding:4px 7px;
    border-bottom:1px solid var(--line); font-size:0.75rem; }
  .combo-row .name { font-weight:600; }
  .combo-row .method { color:var(--sub); font-size:0.7rem; margin-left:5px; }
  .temp-badge {
    display:inline-block; background:#fdf3e4; color:var(--temp);
    padding:1px 7px; border-radius:99px; font-size:0.7rem; font-weight:700;
  }
  .case-item {
    display:flex; justify-content:space-between; align-items:center;
    gap:8px; padding:6px 9px; border:1px solid var(--line); border-radius:6px;
    margin-bottom:5px; font-size:0.75rem; background:var(--panel);
    cursor:pointer;
  }
  .case-item.active { border-color:var(--accent); background:#f2f7ff; }
  .case-item .meta { color:var(--sub); font-size:0.7rem; }
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
    <div class="ctl">
      <label for="sel-view">View</label>
      <select id="sel-view" onchange="onViewChanged()"></select>
    </div>
    <div class="sep"></div>
    <button class="ghost" onclick="runAnalysis()" style="font-weight:700;color:var(--accent);border-color:var(--accent);">&#9654; Solve</button>
  </div>
  <div class="ribbon-status">
    <span>Nodes: <b id="st-nodes">8</b></span>
    <span>Members: <b id="st-members">12</b></span>
    <span>DOF: <b id="st-dof">--</b></span>
    <span>Loads: <b id="st-loads">0</b></span>
    <span>Cases: <b id="st-cases">9</b> &middot; Combos: <b id="st-combos">18</b></span>
    <span id="st-result"><span class="pill pill-idle">Not solved</span></span>
    <span style="margin-left:auto;">Rev 3 &middot; Direct Stiffness (JS) &middot; Excel-driven</span>
  </div>
</div>

<div class="layout">
  <div class="viewer-col">
    <canvas id="view3d"></canvas>
    <div class="view-controls">
      <button id="btn-loads" class="active" onclick="toggleLoads()">Load arrows</button>
      <button id="btn-deformed" onclick="toggleDeformed()">Deformed</button>
      <button id="btn-labels" class="active" onclick="toggleLabels()">Labels</button>
      <button id="btn-grid" class="active" onclick="toggleGrid()">Grid</button>
      <button id="btn-band" class="active" onclick="toggleBand()">Load band</button>
      <button onclick="resetView()">Reset view</button>
    </div>
    <div class="view-legend" id="view-legend"></div>
    <div class="view-hint">Drag to orbit &middot; Shift+drag to pan &middot; Wheel to zoom &middot; Click a node or member to select</div>
  </div>

  <div class="side-col">
    <div class="tabs">
      <button class="active" onclick="setTab(0)">Model</button>
      <button onclick="setTab(1)">Loads</button>
      <button onclick="setTab(2)">Cases</button>
      <button onclick="setTab(3)">Combos</button>
      <button onclick="setTab(4)">Validation</button>
      <button onclick="setTab(5)">Results</button>
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
        <h3>Diaphragm definition</h3>
        <div class="kv" id="kv-diaphragm"></div>
        <div class="note" id="note-diaphragm" style="margin-top:6px;"></div>
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
          property (Iy), about local z the strong-axis property (Ix).
          The utilisation ratio is an elastic superposition stress estimate,
          <b>not</b> a full AISC interaction-equation capacity check.<br><br>
          The load framework shows member distributed and point loads for the
          selected case but does <b>not</b> convert them to equivalent nodal
          loads in the FE solve. To solve a distributed-load case, use the
          approximate nodal loads (or enter equivalent nodal loads manually).
        </div>
      </div>
    </div>

    <!-- LOADS (hand-entered) -->
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
        <h3>Applied hand-entered loads <button class="ghost" style="float:right;padding:2px 8px;" onclick="clearLoads()">Clear all</button></h3>
        <div id="load-list"></div>
      </div>

      <button class="primary" onclick="runAnalysis()">Solve</button>
    </div>

    <!-- CASES -->
    <div class="panel" id="panel-2">
      <div class="card">
        <h3>Rev 3 load cases</h3>
        <div class="note" style="margin-bottom:8px;">
          Click a case to display its loads in the viewer. Nodal loads of the
          selected case are <b>also</b> fed into the FE solve (press Solve);
          member distributed / point loads are drawn but not solved.
        </div>
        <div id="case-list"></div>
      </div>
      <div class="card">
        <h3>Case detail</h3>
        <div id="case-detail" class="note"></div>
      </div>
    </div>

    <!-- COMBOS -->
    <div class="panel" id="panel-3">
      <div class="card">
        <h3>NSCP load combinations</h3>
        <div class="note" style="margin-bottom:8px;">
          LRFD and ASD combinations. Temperature combinations are added as
          separate entries. Combinations reference load cases by category;
          the underlying loads are not duplicated.
        </div>
        <div id="combo-list"></div>
      </div>
      <div class="card">
        <h3>Combination detail</h3>
        <div id="combo-detail" class="note"></div>
      </div>
    </div>

    <!-- VALIDATION -->
    <div class="panel" id="panel-4">
      <div class="card">
        <h3>Load validation summary</h3>
        <div class="note" style="margin-bottom:8px;">
          Total applied load per case, number of loaded nodes/members, and
          computed totals. Distributed loads are integrated and reported as
          total applied load.
        </div>
        <div id="validation-summary"></div>
      </div>
      <div class="card">
        <h3>Temperature load verification</h3>
        <div id="temp-verification" class="note"></div>
      </div>
      <div class="card">
        <h3>Scope note</h3>
        <div class="note">
          This validation panel verifies <b>load application</b> and
          <b>load totals</b>. It does not verify structural response &mdash;
          that is the browser FE solver's job (Results tab). Member distributed
          and point loads are drawn but not converted to equivalent nodal
          loads in this revision.
        </div>
      </div>
    </div>

    <!-- RESULTS -->
    <div class="panel" id="panel-5">
      <div id="banner" class="banner banner-idle">Not solved yet &mdash; add loads or pick a case, then press Solve.</div>
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
const LOAD_CASES = __LOAD_CASES_JSON__;
const COMBINATIONS = __COMBINATIONS_JSON__;
const DIAPHRAGM = __DIAPHRAGM_JSON__;
const VALIDATION = __VALIDATION_JSON__;
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
let currentView = "LC1";   // "LC1".."LC9" or "COMB1".."COMB18"

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
  populateViewSelect();
}

function populateMaterialSelect() {
  const sys = CATALOG.systems[selection.units];
  const sel = document.getElementById("sel-material");
  sel.innerHTML = "";
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

function populateViewSelect() {
  const sel = document.getElementById("sel-view");
  sel.innerHTML = "";
  const g1 = document.createElement("optgroup");
  g1.label = "Load Cases";
  LOAD_CASES.forEach((lc) => {
    const o = document.createElement("option");
    o.value = "LC" + lc.id;
    o.textContent = "LC" + lc.id + ": " + lc.name;
    g1.appendChild(o);
  });
  sel.appendChild(g1);
  const g2 = document.createElement("optgroup");
  g2.label = "Combinations";
  COMBINATIONS.forEach((c) => {
    const o = document.createElement("option");
    o.value = "COMB" + c.id;
    o.textContent = c.name;
    g2.appendChild(o);
  });
  sel.appendChild(g2);
  sel.value = currentView;
}

// ---- selection change handlers ----------------------------------------
function onUnitsChanged() {
  const newUnits = document.getElementById("sel-units").value;
  const newSys = CATALOG.systems[newUnits];
  const equiv = CATALOG.section_equiv[selection.section];
  const newSection = newSys.sections[equiv] ? equiv
    : (newSys.sections[selection.section] ? selection.section : newSys.section_labels[0]);
  const newMaterial = newSys.materials[selection.material]
    ? selection.material : newSys.material_labels[0];
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
function onViewChanged() {
  currentView = document.getElementById("sel-view").value;
  renderLegend();
  renderCaseDetail();
  renderComboDetail();
  drawScene();
}

// ---- model rebuild -----------------------------------------------------
function rebuildModel() {
  currentModel = deriveModel(selection);
  lastResult = null;
  renderModelPanel();
  renderLoadLabels();
  renderLoadList();
  renderCaseList();
  renderComboList();
  renderValidationSummary();
  renderTempVerification();
  renderLegend();
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
    <div class="k">Density</div><div class="v">${mat.density === null ? "--" : fmt(mat.density,2) + " " + mat.density_unit}</div>
    <div class="k">Therm. coeff.</div><div class="v">${fmt(mat.alpha,6)} 1/\u00b0C</div>`;

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

  const d = DIAPHRAGM;
  document.getElementById("kv-diaphragm").innerHTML = `
    <div class="k">Name</div><div class="v">${d.name}</div>
    <div class="k">Master node</div><div class="v">N${d.master_node}</div>
    <div class="k">Constrained nodes</div><div class="v">${d.constrained_nodes.join(", ")}</div>
    <div class="k">Constrained DOF</div><div class="v">${d.degrees_of_freedom.join(", ")}</div>
    <div class="k">Free DOF</div><div class="v">${d.free_dof.join(", ")}</div>
    <div class="k">Equations generated</div><div class="v">${d.n_equations}</div>`;
  document.getElementById("note-diaphragm").textContent =
    "The diaphragm constraint is a real structural constraint definition. " +
    "The constraint equations are generated and tested, but are not eliminated " +
    "into the FE equation set in this revision.";

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
      k.toUpperCase() + " (" + sys.display_force_unit + ")";
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

// ---- hand-entered loads -----------------------------------------------
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
    box.innerHTML = '<div class="empty">No hand-entered loads. Add one above, or select a load case to display its loads.</div>';
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

// ---- case / combo rendering -------------------------------------------
function renderCaseList() {
  const box = document.getElementById("case-list");
  box.innerHTML = LOAD_CASES.map((lc) => {
    const active = currentView === "LC" + lc.id;
    return `<div class="case-item ${active ? "active" : ""}" onclick="selectView('LC${lc.id}')">
      <span><b>LC${lc.id}: ${lc.name}</b><br>
        <span class="meta">${lc.category} &middot; ${lc.n_loads} load(s)</span></span>
      ${lc.category === "T" ? '<span class="temp-badge">T</span>' : ''}
    </div>`;
  }).join("");
}

function renderCaseDetail() {
  const box = document.getElementById("case-detail");
  if (!currentView.startsWith("LC")) { box.innerHTML = ""; return; }
  const id = parseInt(currentView.slice(2), 10);
  const lc = LOAD_CASES.find((c) => c.id === id);
  if (!lc) { box.innerHTML = ""; return; }
  let html = `<b>${lc.name}</b><br>${lc.description}<br><br>`;
  html += `<b>Category:</b> ${lc.category}<br>`;
  html += `<b>Number of loads:</b> ${lc.n_loads}<br>`;
  if (lc.loads) {
    html += `<br><b>Loads:</b><br>`;
    lc.loads.forEach((l) => {
      if (l.type === "nodal") {
        html += `&bull; Node ${l.node_id}: Fx=${fmt(l.fx,3)}, Fy=${fmt(l.fy,3)}, Fz=${fmt(l.fz,3)} kN<br>`;
      } else if (l.type === "udl") {
        html += `&bull; Member M${l.member_id}: ${l.magnitude} kN/m ${l.direction} (factor ${l.direction_factor}), L=${fmt(l.length,3)} m<br>`;
      } else if (l.type === "point") {
        html += `&bull; Member M${l.member_id}: ${l.magnitude} kN at ${l.location*100}% ${l.direction}<br>`;
      } else if (l.type === "temp") {
        html += `&bull; Temperature: +${l.temperature_change} \u00b0C on ${l.member_ids.length} members<br>`;
      }
    });
  }
  box.innerHTML = html;
}

function renderComboList() {
  const box = document.getElementById("combo-list");
  box.innerHTML = COMBINATIONS.map((c) => {
    const active = currentView === "COMB" + c.id;
    return `<div class="combo-row" style="cursor:pointer;${active ? 'background:#f2f7ff;' : ''}"
              onclick="selectView('COMB${c.id}')">
      <span><span class="name">${c.name}</span>
        <span class="method">${c.design_method}</span></span>
      <span class="note">${c.n_loads} ref(s)</span>
    </div>`;
  }).join("");
}

function renderComboDetail() {
  const box = document.getElementById("combo-detail");
  if (!currentView.startsWith("COMB")) { box.innerHTML = ""; return; }
  const id = parseInt(currentView.slice(4), 10);
  const c = COMBINATIONS.find((x) => x.id === id);
  if (!c) { box.innerHTML = ""; return; }
  let html = `<b>${c.name}</b><br>`;
  html += `Design method: ${c.design_method}<br>`;
  html += `Factors: ${JSON.stringify(c.factors)}<br>`;
  html += `Referenced categories: ${c.n_loads}<br>`;
  box.innerHTML = html;
}

function renderLegend() {
  const box = document.getElementById("view-legend");
  if (currentView.startsWith("COMB")) {
    const c = COMBINATIONS.find((x) => "COMB" + x.id === currentView);
    if (c) {
      box.innerHTML = `<h4>${c.name}</h4>` +
        Object.entries(c.factors).map(([k, v]) =>
          `<div class="row"><span class="swatch" style="background:#e63946;"></span>${k} \u00d7 ${v}</div>`
        ).join("");
      return;
    }
  }
  const id = parseInt(currentView.slice(2), 10);
  const lc = LOAD_CASES.find((x) => x.id === id);
  if (!lc) { box.innerHTML = ""; return; }
  let html = `<h4>LC${lc.id}: ${lc.name}</h4>`;
  if (lc.category === "T") {
    html += `<div class="row"><span class="swatch" style="background:#b45309;"></span>+15 \u00b0C (temperature)</div>`;
  } else if (lc.id === 1) {
    html += `<div class="row"><span class="swatch" style="background:#e63946;"></span>Self-weight, -Y</div>`;
  } else {
    html += `<div class="row"><span class="swatch" style="background:#e63946;"></span>Load, -Y</div>`;
  }
  html += `<div class="note" style="margin-top:5px;">${lc.description}</div>`;
  box.innerHTML = html;
}

function renderValidationSummary() {
  const box = document.getElementById("validation-summary");
  let html = "";
  VALIDATION.case_summaries.forEach((s) => {
    html += `<div style="margin-bottom:8px;">
      <b>LC${s.id}: ${s.name}</b><br>`;
    if (s.self_weight_total !== null) {
      html += `<span class="note">Self-weight total: ${fmt(s.self_weight_total,3)} kN</span><br>`;
    }
    if (s.nodal_total) {
      html += `<span class="note">Nodal: Fx=${fmt(s.nodal_total[0],3)}, Fy=${fmt(s.nodal_total[1],3)}, Fz=${fmt(s.nodal_total[2],3)} kN</span><br>`;
    }
    if (s.udl_total !== null && s.udl_total !== undefined) {
      html += `<span class="note">Distributed total: ${fmt(s.udl_total,3)} kN</span><br>`;
    }
    if (s.temperature) {
      html += `<span class="note">Temperature: +${s.temperature.dT} \u00b0C, \u03b5=${s.temperature.strain.toExponential(3)}</span><br>`;
    }
    html += `</div>`;
  });
  box.innerHTML = html;
}

function renderTempVerification() {
  const box = document.getElementById("temp-verification");
  const t = VALIDATION.temperature_summary;
  if (!t) { box.innerHTML = "No temperature load case."; return; }
  box.innerHTML = `
    <b>${t.name}</b><br>
    Temperature change: +${t.dT} \u00b0C<br>
    Material: ${t.material}<br>
    Coefficient of thermal expansion: ${t.alpha.toExponential(4)} 1/\u00b0C<br>
    Thermal strain: ${t.strain.toExponential(6)}<br>
    Member length: ${fmt(t.length,3)} m<br>
    Free expansion: ${fmt(t.free_expansion * 1000,4)} mm<br>
    Axial rigidity EA: ${fmt(t.EA,1)} kN<br>
    Fully restrained force: ${fmt(t.restrained_force,4)} kN (compression)<br>
    Net force on structure: ${fmt(t.net_force,4)} kN (self-equilibrating)<br>
    <br>
    <b>Restraint states distinguished:</b> free, partially restrained, fully restrained.<br>
    The partially restrained case is reported as <b>indeterminate</b> rather than guessed.
  `;
}

function selectView(v) {
  currentView = v;
  const sel = document.getElementById("sel-view");
  if (sel) sel.value = v;
  renderLegend();
  renderCaseDetail();
  renderComboDetail();
  renderCaseList();
  drawScene();
}

// ---- solve -------------------------------------------------------------
// Build the nodal-load array that the FE solver sees: hand-entered loads,
// plus the NODAL loads of the selected load case (if any). Distributed and
// member point loads are NOT converted to nodal loads in this revision.
function effectiveNodalLoads() {
  const out = loads.map((l) => ({ node: l.node, fx: l.fx, fy: l.fy, fz: l.fz,
                                  mx: l.mx, my: l.my, mz: l.mz }));
  if (currentView.startsWith("LC")) {
    const id = parseInt(currentView.slice(2), 10);
    const lc = LOAD_CASES.find((c) => c.id === id);
    if (lc && lc.loads) {
      lc.loads.forEach((l) => {
        if (l.type === "nodal") {
          out.push({ node: l.node_id, fx: l.fx, fy: l.fy, fz: l.fz,
                     mx: 0, my: 0, mz: 0 });
        }
      });
    }
  }
  return out;
}

function runAnalysis() {
  const banner = document.getElementById("banner");
  const sw = includeSelfWeight();
  const eff = effectiveNodalLoads();
  const selectedCase = currentView.startsWith("LC")
    ? LOAD_CASES.find((c) => "LC" + c.id === currentView) : null;
  if (eff.length === 0 && !sw && !(selectedCase && selectedCase.id === 1)) {
    banner.className = "banner banner-err";
    banner.textContent = "Nothing to solve: no hand-entered loads, no nodal loads in the selected case, and self-weight is off.";
    setTab(5);
    return;
  }
  try {
    // LC1 (self-weight case) implies self-weight ON for that solve.
    const useSW = sw || (selectedCase && selectedCase.id === 1);
    lastResult = solveModel(currentModel, eff, useSW);
    banner.className = "banner banner-ok";
    let msg = `Analysis complete \u2014 no errors. ${eff.length} nodal load(s)`;
    if (useSW) msg += " + self-weight";
    msg += `, ${GEOMETRY.dof.active} active DOF.`;
    if (selectedCase && selectedCase.loads.some((l) => l.type === "udl" || l.type === "point")) {
      msg += " NOTE: member distributed/point loads of this case were drawn but NOT solved.";
    }
    banner.textContent = msg;
    renderResults();
    renderStatus();
    view.showDeformed = true;
    document.getElementById("btn-deformed").classList.add("active");
    drawScene();
    setTab(5);
  } catch (err) {
    lastResult = null;
    banner.className = "banner banner-err";
    banner.textContent = "Solve failed: " + err.message;
    renderStatus();
    setTab(5);
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
  const eff = effectiveNodalLoads();
  document.getElementById("st-loads").textContent = eff.length + (includeSelfWeight() ? " + SW" : "");
  document.getElementById("st-cases").textContent = LOAD_CASES.length;
  document.getElementById("st-combos").textContent = COMBINATIONS.length;
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
function toggleGrid() {
  view.showGrid = !view.showGrid;
  document.getElementById("btn-grid").classList.toggle("active", view.showGrid);
  drawScene();
}
function toggleBand() {
  view.showBand = !view.showBand;
  document.getElementById("btn-band").classList.toggle("active", view.showBand);
  drawScene();
}

// ---- boot --------------------------------------------------------------
populateControls();
rebuildModel();
initViewerEvents();
renderLoadList();
renderCaseList();
renderComboList();
renderValidationSummary();
renderTempVerification();
renderLegend();
drawScene();
</script>
</body>
</html>
"""


# =============================================================================
# SOLVER CORE JS  (from the working Rev 3 -- unchanged)
# =============================================================================

SOLVER_CORE_JS = r"""// ---------------------------------------------------------------------------
// SOLVER CORE
// 12x12 Euler-Bernoulli space-frame element stiffness, global assembly,
// partitioned solve, member end-force recovery.  Cross-validated against a
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
      F[(parseInt(node, 10) - 1) * 6 + 1] -= w;
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


# =============================================================================
# VIEWER JS  (Cartesian grid, load-case-aware rendering)
# =============================================================================

VIEWER_JS = r"""// ---------------------------------------------------------------------------
// VIEWER -- dependency-free 3D wireframe renderer on a plain 2D canvas.
// Cartesian plane grid with labeled X/Y/Z axes, ground plane at Y=0, and
// load-case-aware rendering (nodal arrows, distributed load bands, point
// loads, temperature annotations).
// ---------------------------------------------------------------------------

const view = {
  yaw: -0.6, pitch: -0.35, zoom: 1.0,
  panX: 0, panY: 0,
  dragging: false, panning: false, lastX: 0, lastY: 0, moved: false,
  showLoads: true, showLabels: true, showGrid: true, showBand: true,
  showDeformed: false, dispScale: null,
};

const LOAD_COLOR = "#e63946";
const TEMP_COLOR = "#b45309";
const BAND_ALPHA = 0.30;
const SW_BAND_ALPHA = 0.18;

function rotatePoint(p, yaw, pitch) {
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

function currentLoadCase() {
  if (!currentView.startsWith("LC")) return null;
  const id = parseInt(currentView.slice(2), 10);
  return LOAD_CASES.find((c) => c.id === id);
}
function currentCombination() {
  if (!currentView.startsWith("COMB")) return null;
  const id = parseInt(currentView.slice(4), 10);
  return COMBINATIONS.find((c) => c.id === id);
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
  const edge = M.edgeSolver;

  // Cartesian grid at Y=0
  if (view.showGrid) {
    ctx.strokeStyle = "#e6e9ed";
    ctx.lineWidth = 1;
    const div = 6;
    for (let i = -1; i <= div + 1; i++) {
      const t = (i / div) * edge;
      const a1 = project(t, 0, -edge * 0.3), a2 = project(t, 0, edge * 1.3);
      ctx.beginPath(); ctx.moveTo(a1.sx, a1.sy); ctx.lineTo(a2.sx, a2.sy); ctx.stroke();
      const b1 = project(-edge * 0.3, 0, t), b2 = project(edge * 1.3, 0, t);
      ctx.beginPath(); ctx.moveTo(b1.sx, b1.sy); ctx.lineTo(b2.sx, b2.sy); ctx.stroke();
    }
    const corners = [
      project(-edge * 0.3, 0, -edge * 0.3),
      project(edge * 1.3, 0, -edge * 0.3),
      project(edge * 1.3, 0, edge * 1.3),
      project(-edge * 0.3, 0, edge * 1.3),
    ];
    ctx.beginPath();
    ctx.moveTo(corners[0].sx, corners[0].sy);
    for (let i = 1; i < corners.length; i++) ctx.lineTo(corners[i].sx, corners[i].sy);
    ctx.closePath();
    ctx.fillStyle = "rgba(220, 225, 230, 0.18)";
    ctx.fill();
  }

  drawWorldAxes(ctx, project, edge);

  // Undeformed members, painter's algorithm
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

  // Deformed shape overlay
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

  // Nodes, supports, labels
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

  // Loads
  if (view.showLoads) {
    // hand-entered + self-weight from the solver side
    drawHandEnteredLoads(ctx, project, M);
    // selected case / combination overlay
    const lc = currentLoadCase();
    const comb = currentCombination();
    if (lc) drawLoadCase(ctx, project, M, lc, 1.0);
    if (comb) drawCombination(ctx, project, M, comb);
  }

  drawAxisTriad(ctx, w, h);

  if (view.showDeformed && lastResult) {
    ctx.fillStyle = "#e63946";
    ctx.font = "600 11px -apple-system, Segoe UI, Helvetica, Arial, sans-serif";
    ctx.fillText("Deformed shape \u00d7" + deformedScale().toFixed(1) + " (exaggerated)", 12, 20);
  }
}

function drawWorldAxes(ctx, project, edge) {
  const origin = project(0, 0, 0);
  const xEnd = project(edge * 1.35, 0, 0);
  const yEnd = project(0, edge * 1.35, 0);
  const zEnd = project(0, 0, edge * 1.35);
  const axes = [
    { end: xEnd, color: "#d11", label: "X" },
    { end: yEnd, color: "#1a1", label: "Y" },
    { end: zEnd, color: "#16c", label: "Z" },
  ];
  axes.forEach((a) => {
    ctx.strokeStyle = a.color;
    ctx.lineWidth = 2;
    ctx.beginPath();
    ctx.moveTo(origin.sx, origin.sy);
    ctx.lineTo(a.end.sx, a.end.sy);
    ctx.stroke();
    ctx.fillStyle = a.color;
    ctx.font = "700 13px -apple-system, Segoe UI, Helvetica, Arial, sans-serif";
    ctx.fillText(a.label, a.end.sx + 4, a.end.sy + 4);
  });
}

function drawHandEnteredLoads(ctx, project, M) {
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

function drawLoadCase(ctx, project, M, lc, factor) {
  const nodeById = {};
  M.nodes.forEach((n) => { nodeById[n.id] = n; });

  if (lc.id === 1) {
    // Self-weight indicated along every member, arrow pointing -Y
    M.members.forEach((m) => {
      const ni = nodeById[m.i], nj = nodeById[m.j];
      const mid = { x: (ni.x + nj.x) / 2, y: (ni.y + nj.y) / 2, z: (ni.z + nj.z) / 2 };
      drawArrow(ctx, project, mid.x, mid.y, mid.z, 0, -1, 0, M.edgeSolver * 0.18, LOAD_COLOR);
    });
    return;
  }

  if (lc.category === "T") {
    M.members.forEach((m) => {
      const ni = nodeById[m.i], nj = nodeById[m.j];
      const mid = { x: (ni.x + nj.x) / 2, y: (ni.y + nj.y) / 2, z: (ni.z + nj.z) / 2 };
      const p = project(mid.x, mid.y, mid.z);
      ctx.fillStyle = TEMP_COLOR;
      ctx.font = "700 11px -apple-system, Segoe UI, Helvetica, Arial, sans-serif";
      ctx.fillText("+" + lc.loads[0].temperature_change + "\u00b0C", p.sx - 12, p.sy - 6);
    });
    return;
  }

  if (!lc.loads) return;
  lc.loads.forEach((l) => {
    if (l.type === "nodal") {
      const n = nodeById[l.node_id];
      if (!n) return;
      const fx = l.fx * factor, fy = l.fy * factor, fz = l.fz * factor;
      const mag = Math.hypot(fx, fy, fz);
      if (mag < 1e-9) return;
      drawArrow(ctx, project, n.x, n.y, n.z, fx / mag, fy / mag, fz / mag,
                M.edgeSolver * 0.22, LOAD_COLOR);
    } else if (l.type === "udl") {
      const m = M.members.find((x) => x.id === l.member_id);
      if (!m) return;
      const ni = nodeById[m.i], nj = nodeById[m.j];
      drawDistributedLoad(ctx, project, ni, nj, l.direction, l.direction_factor,
                          l.magnitude * factor, M.edgeSolver);
    } else if (l.type === "point") {
      const m = M.members.find((x) => x.id === l.member_id);
      if (!m) return;
      const ni = nodeById[m.i], nj = nodeById[m.j];
      const t = l.location;
      const px = ni.x + t * (nj.x - ni.x);
      const py = ni.y + t * (nj.y - ni.y);
      const pz = ni.z + t * (nj.z - ni.z);
      drawArrow(ctx, project, px, py, pz, 0, -1, 0,
                M.edgeSolver * 0.20, LOAD_COLOR);
    }
  });
}

function drawCombination(ctx, project, M, comb) {
  const nodeById = {};
  M.nodes.forEach((n) => { nodeById[n.id] = n; });
  Object.entries(comb.factors).forEach(([cat, f]) => {
    const lcs = LOAD_CASES.filter((c) => c.category === cat);
    lcs.forEach((lc) => drawLoadCase(ctx, project, M, lc, f));
  });
}

function drawArrow(ctx, project, x, y, z, dx, dy, dz, length, color) {
  const tail = { x: x - dx * length, y: y - dy * length, z: z - dz * length };
  const pTail = project(tail.x, tail.y, tail.z);
  const pHead = project(x, y, z);
  ctx.strokeStyle = color;
  ctx.fillStyle = color;
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
}

function drawDistributedLoad(ctx, project, ni, nj, direction, factor, magnitude, edge) {
  const pi = project(ni.x, ni.y, ni.z);
  const pj = project(nj.x, nj.y, nj.z);
  const dx = pj.sx - pi.sx, dy = pj.sy - pi.sy;
  const len = Math.hypot(dx, dy) || 1;
  const px = -dy / len, py = dx / len;
  const sign = factor < 0 ? 1 : -1;
  const bandWidth = Math.min(28, Math.max(14, magnitude * 3));
  const nArrows = Math.max(6, Math.round(len / 14));

  if (view.showBand) {
    ctx.beginPath();
    for (let i = 0; i <= nArrows; i++) {
      const t = i / nArrows;
      const sx = pi.sx + t * dx + sign * px * bandWidth;
      const sy = pi.sy + t * dy + sign * py * bandWidth;
      if (i === 0) ctx.moveTo(sx, sy); else ctx.lineTo(sx, sy);
    }
    for (let i = nArrows; i >= 0; i--) {
      const t = i / nArrows;
      const sx = pi.sx + t * dx;
      const sy = pi.sy + t * dy;
      ctx.lineTo(sx, sy);
    }
    ctx.closePath();
    ctx.fillStyle = "rgba(230, 57, 70, " + BAND_ALPHA + ")";
    ctx.fill();
    ctx.strokeStyle = "rgba(230, 57, 70, 0.9)";
    ctx.lineWidth = 1.5;
    ctx.stroke();
  }

  ctx.strokeStyle = LOAD_COLOR;
  ctx.fillStyle = LOAD_COLOR;
  ctx.lineWidth = 1.5;
  for (let i = 0; i <= nArrows; i++) {
    const t = i / nArrows;
    const bx = pi.sx + t * dx;
    const by = pi.sy + t * dy;
    const tx = bx + sign * px * bandWidth;
    const ty = by + sign * py * bandWidth;
    ctx.beginPath();
    ctx.moveTo(tx, ty);
    ctx.lineTo(bx, by);
    ctx.stroke();
    const ang = Math.atan2(by - ty, bx - tx);
    const hl = 5;
    ctx.beginPath();
    ctx.moveTo(bx, by);
    ctx.lineTo(bx - hl * Math.cos(ang - 0.5), by - hl * Math.sin(ang - 0.5));
    ctx.lineTo(bx - hl * Math.cos(ang + 0.5), by - hl * Math.sin(ang + 0.5));
    ctx.closePath();
    ctx.fill();
  }
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
  const fit = (currentModel.edgeSolver * 0.18) / lastResult.maxDisp;
  return Math.min(fit, 2000);
}

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


# =============================================================================
# HTML RENDERER
# =============================================================================

def render_html(catalog, geometry, load_cases, combinations, diaphragm,
                validation, out_path):
    page = PAGE_TEMPLATE
    page = page.replace("__CATALOG_JSON__", json.dumps(catalog))
    page = page.replace("__GEOMETRY_JSON__", json.dumps(geometry))

    lc_json = []
    for lc in load_cases:
        loads_json = []
        for l in lc.loads:
            if isinstance(l, NodalLoad):
                loads_json.append({
                    "type": "nodal", "node_id": l.node_id,
                    "fx": l.fx, "fy": l.fy, "fz": l.fz,
                    "mx": l.mx, "my": l.my, "mz": l.mz})
            elif isinstance(l, MemberDistributedLoad):
                loads_json.append({
                    "type": "udl", "member_id": l.member_id,
                    "direction": l.direction, "magnitude": l.magnitude,
                    "length": l.length, "direction_factor": l.direction_factor})
            elif isinstance(l, MemberPointLoad):
                loads_json.append({
                    "type": "point", "member_id": l.member_id,
                    "location": l.location, "direction": l.direction,
                    "magnitude": l.magnitude})
            elif isinstance(l, TemperatureLoad):
                loads_json.append({
                    "type": "temp", "member_ids": l.member_ids,
                    "temperature_change": l.temperature_change,
                    "alpha": l.alpha})
        lc_json.append({
            "id": lc.id, "name": lc.name, "category": lc.category,
            "description": lc.description, "n_loads": len(lc.loads),
            "loads": loads_json})
    page = page.replace("__LOAD_CASES_JSON__", json.dumps(lc_json))

    comb_json = []
    for c in combinations:
        comb_json.append({
            "id": c.id, "name": c.name, "design_method": c.design_method,
            "factors": c.factors, "n_loads": len(c.factors)})
    page = page.replace("__COMBINATIONS_JSON__", json.dumps(comb_json))

    dia_json = {
        "name": diaphragm.name, "master_node": diaphragm.master_node,
        "constrained_nodes": diaphragm.constrained_nodes,
        "degrees_of_freedom": diaphragm.degrees_of_freedom,
        "free_dof": [d for d in ["UX", "UY", "UZ", "RX", "RY", "RZ"]
                     if d not in diaphragm.degrees_of_freedom],
        "n_equations": len(diaphragm.generate_constraint_equations())}
    page = page.replace("__DIAPHRAGM_JSON__", json.dumps(dia_json))

    val_json = {
        "case_summaries": validation.case_summaries,
        "temperature_summary": validation.temperature_summary}
    page = page.replace("__VALIDATION_JSON__", json.dumps(val_json))

    page = page.replace("__SOLVER_CORE__", SOLVER_CORE_JS)
    page = page.replace("__VIEWER_JS__", VIEWER_JS)

    with open(out_path, "w", encoding="utf-8") as f:
        f.write(page)
    return out_path


# =============================================================================
# VALIDATION (solver-side, unchanged from working Rev 3)
# =============================================================================

def run_validation(catalog, geometry, unit_db, out_path):
    checks = []

    def check(name, passed, detail=""):
        checks.append((name, bool(passed), detail))

    met = catalog["systems"]["metric"]
    imp = catalog["systems"]["imperial"]

    check("Unit workbook located, both systems parsed",
          len(unit_db.systems) == 2, f"systems = {unit_db.systems}")
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

    a36_met = met["materials"].get(DEFAULT_MATERIAL_LABEL)
    a36_imp = imp["materials"].get(DEFAULT_MATERIAL_LABEL)
    if a36_met and a36_imp:
        ksi_to_mpa = unit_db.factor_pair("ksi", "MPa")["imperial_to_metric"]
        e_converted = a36_imp["E"] * ksi_to_mpa
        rel = abs(e_converted - a36_met["E"]) / a36_met["E"]
        check("A36 E agrees across unit systems (cross-workbook consistency)",
              rel < 0.01, f"{a36_imp['E']} ksi -> {e_converted:.0f} MPa vs "
                          f"{a36_met['E']} MPa ({rel*100:.2f}% diff)")

    in4_to_mm4 = unit_db.factor_pair("in\u2074", "mm\u2074")["imperial_to_metric"]
    start_sec_met = met["sections"].get(DEFAULT_MEMBER_SIZE_LABEL)
    equiv_label = catalog["section_equiv"].get(DEFAULT_MEMBER_SIZE_LABEL)
    start_sec_imp = imp["sections"].get(equiv_label)
    if start_sec_met and start_sec_imp:
        ix_conv = start_sec_imp["Ix"] * in4_to_mm4
        rel = abs(ix_conv - start_sec_met["Ix"]) / start_sec_met["Ix"]
        check("Section Ix agrees across unit systems (descaling is correct)",
              rel < 0.02,
              f"{equiv_label} {start_sec_imp['Ix']} in\u2074 -> {ix_conv:.3e} mm\u2074 vs "
              f"{DEFAULT_MEMBER_SIZE_LABEL} {start_sec_met['Ix']:.3e} mm\u2074 "
              f"({rel*100:.2f}% diff)")

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
          and "__VIEWER_JS__" not in page and "__GEOMETRY_JSON__" not in page
          and "__LOAD_CASES_JSON__" not in page
          and "__COMBINATIONS_JSON__" not in page
          and "__DIAPHRAGM_JSON__" not in page
          and "__VALIDATION_JSON__" not in page)
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
# DIAGNOSTIC MODES
# =============================================================================

def cmd_list_materials(args):
    system = "Standard Metric" if args.units == "metric" else "Imperial"
    db = MaterialDatabase().load(system)
    print(f"Materials in {db.path} ({len(db.materials)} total):\n")
    for label, props in sorted(db.materials.items()):
        print(f"  {label:<20} category={props['Category']}")


def cmd_list_sections(args, shape_type):
    db = MemberSizeDatabase()
    idx = db.metric_idx if args.units == "metric" else db.imperial_idx
    rows = db.rows_of_type(shape_type)
    print(f"{shape_type}-shapes in {db.path} ({len(rows)} total):\n")
    for row in rows:
        print(f"  {row[idx['AISC_Manual_Label']]}")


# =============================================================================
# MAIN
# =============================================================================

def main(argv=None):
    args = parse_args(argv)

    if args.test:
        return run_load_tests(args.verbose)
    if args.list_materials:
        cmd_list_materials(args)
        return 0
    if args.list_sections is not None:
        cmd_list_sections(args, args.list_sections)
        return 0

    try:
        banner(1, "INSPECT -- locate and read the Excel workbooks")
        unit_db = UnitDatabase()
        print(f"  units/        -> {os.path.relpath(unit_db.path, HERE)}")
        print(f"                   unit systems: {unit_db.systems}")
        size_db = MemberSizeDatabase()
        print(f"  member_size/  -> {os.path.relpath(size_db.path, HERE)}")
        print(f"                   {len(size_db.rows_of_type('W'))} W-shape rows")
        mat_db = MaterialDatabase().load("Standard Metric")
        print(f"  materials/    -> {os.path.relpath(mat_db.path, HERE)}")
        print(f"                   {len(mat_db.materials)} materials")

        banner(2, "CATALOG -- read EVERY material and section, both unit systems")
        catalog = build_catalog(unit_db, size_db, args)
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

        banner(3, "LOAD CASES -- build the Rev 3 load framework")
        geometry = build_static_geometry()
        load_cases, diaphragm, temp_loads, notes = build_load_cases(geometry, mat_db)
        for lc in load_cases:
            print(f"  LC{lc.id}: {lc.name:<28} {lc.category}  n_loads={len(lc.loads)}")
        for note in notes:
            print(f"  NOTE: {note}")

        banner(4, "COMBINATIONS -- NSCP LRFD + ASD")
        combinations, by_category = build_combinations(load_cases)
        print(f"  {len(combinations)} combinations generated")

        banner(5, "VALIDATION -- self-weight and load totals")
        self_weight_total, density, area = compute_self_weight(geometry, mat_db, size_db)
        print(f"  Self-weight total (gamma * A * L): {self_weight_total:.3f} kN")
        print(f"  Density used: {density:.3f} kN/m^3")
        print(f"  Section area: {area:.6f} m^2")
        validation = validate_load_cases(
            load_cases, geometry, mat_db, size_db, diaphragm,
            self_weight_total, density, area)

        banner(6, "RENDER -- build the interactive solver page")
        os.makedirs(args.outdir, exist_ok=True)
        out_path = os.path.join(args.outdir, "REV3_solver.html")
        render_html(catalog, geometry, load_cases, combinations, diaphragm,
                    validation, out_path)
        print(f"  Solver page: {out_path} ({os.path.getsize(out_path)/1024:.0f} KB)")

        banner(7, "VALIDATE (solver-side checks)")
        all_passed, checks = run_validation(catalog, geometry, unit_db, out_path)
        print_validation_report(checks)
        print()
        print("  ALL CHECKS PASSED" if all_passed else "  ONE OR MORE CHECKS FAILED -- see above")

        if args.verify_loads:
            banner(8, "VERIFY-LOADS -- write the load-validation report")
            report_path = os.path.join(args.outdir, "cube_rev3_verification_report.txt")
            with open(report_path, "w", encoding="utf-8") as f:
                f.write(validation.to_text())
            print(f"  Verification report: {report_path}")
            return 0 if all_passed else 1

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