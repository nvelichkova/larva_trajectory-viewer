#!/usr/bin/env python3
"""
Larva Trajectory Viewer
-----------------------
Small PyQt5 GUI for plotting the x/y trajectory of a single larva from an
Excel tracking export (e.g. FIM-Track). Reads two layouts:
  A) first row with  <point> | x | y  column blocks
  B) one column per coordinate, e.g.  larva(926)_head_x, larva(926)_head_y

Features
  * Trajectories mode: paths of chosen body points (head, tail, mom, spine)
  * Body lines mode: one head-tail (or head-spine-tail) line per frame, with
    optional head/tail/spine/mom dots and trajectory overlays (e.g. mom)
  * colour by body point or by time; adjustable colours, widths, styles
  * body skeleton every N frames, start/end markers
  * frame range, time conversion (fps -> s / min)
  * spatial calibration (px per cm for x and y; axes in px / cm / mm)
  * rotation (90 / 180 / 270), axis ranges, closed box, equal aspect
  * export the current figure as a vector PDF (editable text)

See README.md for full instructions.

Usage:  python larva_trajectory_gui.py [file.xlsx]
Requires: PyQt5, matplotlib, pandas, openpyxl, numpy
"""
import sys
import re
import os
import numpy as np
import pandas as pd

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QHBoxLayout, QVBoxLayout, QGroupBox,
    QCheckBox, QPushButton, QFileDialog, QLabel, QSpinBox, QComboBox,
    QFormLayout, QMessageBox, QLineEdit, QDoubleSpinBox, QScrollArea, QColorDialog,
)
from PyQt5.QtGui import QColor, QIcon, QPixmap
from matplotlib.backends.backend_qt5agg import (
    FigureCanvasQTAgg as FigureCanvas, NavigationToolbar2QT as NavToolbar,
)
from matplotlib.figure import Figure
from matplotlib.collections import LineCollection
import matplotlib as mpl

mpl.rcParams["pdf.fonttype"] = 42   # editable text in Illustrator
mpl.rcParams["ps.fonttype"] = 42

# Body order used for the skeleton (head -> tail)
SKELETON_ORDER = ["head", "spinepoint_1", "spinepoint_2", "spinepoint_3", "tail"]
POINT_COLOURS = {
    "spine": "#555555",   # colour for spine dots on body lines
    "mom": "#000000",
    "head": "#d62728",
    "spinepoint_1": "#ff7f0e",
    "spinepoint_2": "#bcbd22",
    "spinepoint_3": "#2ca02c",
    "tail": "#1f77b4",
}


# --------------------------------------------------------------------------
# Data loading
# --------------------------------------------------------------------------
def _blocks_layout(raw):
    """Layout A: first row has  <point> | x | y  blocks (FIM 'table for figure')."""
    header = raw.iloc[0].tolist()
    points = {}
    for c in range(len(header) - 2):
        name = header[c]
        if (isinstance(name, str) and name.strip()
                and str(header[c + 1]).strip().lower() == "x"
                and str(header[c + 2]).strip().lower() == "y"):
            xy = raw.iloc[1:, [c + 1, c + 2]].apply(pd.to_numeric, errors="coerce")
            points[name.strip()] = xy.to_numpy(dtype=float)
    return {"larva": points} if points else {}


_COL_RE = re.compile(r"^(?:(?P<larva>.*?)[_\s]+)?(?P<point>[A-Za-z][\w]*?)[_\s]*(?P<axis>[xXyY])$")


def _columns_layout(raw):
    """Layout B: one column per coordinate, e.g.  larva(926)_head_x,
    larva(926)_head_y, larva(926)_mom_x ...  (optionally a 'timepoint' column).
    Several larvae in one sheet are grouped by their prefix."""
    header = [str(h).strip() for h in raw.iloc[0].tolist()]
    cols = {}   # (larva, point) -> {'x': col, 'y': col}
    known = {"head", "tail", "mom"}
    for c, h in enumerate(header):
        if not h or h.lower() in ("nan", "timepoint", "time", "frame"):
            continue
        m = _COL_RE.match(h)
        if not m:
            continue
        larva = (m.group("larva") or "larva").strip("_ ")
        point = m.group("point")
        # handle names like  spinepoint_1_x  where the regex puts '_1' oddly
        parts = re.split(r"[_\s]+", h)
        if len(parts) >= 3 and parts[-1].lower() in ("x", "y"):
            # try longest known point name ending before the axis
            for k in range(len(parts) - 1, 0, -1):
                cand = "_".join(parts[k:-1])
                if cand.lower() in known or cand.lower().startswith("spinepoint"):
                    point = cand
                    larva = "_".join(parts[:k]) or "larva"
                    break
        cols.setdefault((larva, point), {})[m.group("axis").lower()] = c
    larvae = {}
    for (larva, point), xy in cols.items():
        if "x" in xy and "y" in xy:
            arr = raw.iloc[1:, [xy["x"], xy["y"]]].apply(pd.to_numeric, errors="coerce")
            larvae.setdefault(larva, {})[point] = arr.to_numpy(dtype=float)
    return larvae


def load_larva_xlsx(path, sheet=0):
    """Return dict  larva_id -> {point_name -> (N,2) float array of x,y}.

    Understands both layouts:
      A) first row with  <point> | x | y  column blocks
      B) one column per coordinate:  <larva>_<point>_x / _y
    Trailing rows where every point is empty are trimmed.
    """
    raw = pd.read_excel(path, sheet_name=sheet, header=None)
    larvae = _blocks_layout(raw) or _columns_layout(raw)
    if not larvae:
        raise ValueError(
            "Could not find coordinate columns.\nExpected either '<point> | x | y' "
            "blocks or columns named like 'larva_head_x', 'larva_head_y'.")
    for lid, points in larvae.items():
        n = max(len(v) for v in points.values())
        for k, v in points.items():
            if len(v) < n:
                points[k] = np.vstack([v, np.full((n - len(v), 2), np.nan)])
        # trim trailing frames where nothing was tracked
        any_valid = np.zeros(n, bool)
        for v in points.values():
            any_valid |= ~np.isnan(v).any(axis=1)
        last = np.where(any_valid)[0]
        if len(last):
            for k in points:
                points[k] = points[k][: last[-1] + 1]
    return larvae


# --------------------------------------------------------------------------
# GUI
# --------------------------------------------------------------------------
class TrajectoryViewer(QMainWindow):
    def __init__(self, path=None):
        super().__init__()
        self.setWindowTitle("Larva Trajectory Viewer")
        self.resize(1200, 800)
        self.points = {}
        self.path = None
        self.checks = {}

        central = QWidget()
        self.setCentralWidget(central)
        main = QHBoxLayout(central)

        # ---------------- left control panel ----------------
        # scrollable panel so all controls are reachable on small screens
        panel = QWidget()
        ctrl = QVBoxLayout(panel)
        scroll = QScrollArea()
        scroll.setWidget(panel)
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        scroll.setFixedWidth(300)
        main.addWidget(scroll, 0)

        btn_open = QPushButton("Open Excel file…")
        btn_open.clicked.connect(self.open_file)
        ctrl.addWidget(btn_open)
        self.lbl_file = QLabel("No file loaded")
        self.lbl_file.setWordWrap(True)
        self.lbl_file.setMaximumWidth(260)
        ctrl.addWidget(self.lbl_file)
        self.lbl_larva = QLabel("Larva (several in this file):")
        self.cmb_larva = QComboBox()
        self.cmb_larva.currentIndexChanged.connect(self.select_larva)
        ctrl.addWidget(self.lbl_larva)
        ctrl.addWidget(self.cmb_larva)
        self.lbl_larva.hide()
        self.cmb_larva.hide()

        # points
        self.grp_points = QGroupBox("Body points")
        self.pts_layout = QVBoxLayout(self.grp_points)
        ctrl.addWidget(self.grp_points)

        grp_pre = QGroupBox("Presets")
        pre = QVBoxLayout(grp_pre)
        for label, sel in [
            ("Head only", ["head"]),
            ("Tail only", ["tail"]),
            ("Head + tail", ["head", "tail"]),
            ("Head + spine points", ["head", "spinepoint_1", "spinepoint_2", "spinepoint_3"]),
            ("Tail + spine points", ["tail", "spinepoint_1", "spinepoint_2", "spinepoint_3"]),
            ("Head + tail + spine", SKELETON_ORDER),
            ("Centre of mass (mom)", ["mom"]),
            ("Mom + head + tail", ["mom", "head", "tail"]),
            ("All", None),
        ]:
            b = QPushButton(label)
            b.clicked.connect(lambda _, s=sel: self.apply_preset(s))
            pre.addWidget(b)
        ctrl.addWidget(grp_pre)

        # display options
        grp_mode = QGroupBox("Plot mode")
        mform = QFormLayout(grp_mode)
        self.cmb_mode = QComboBox()
        self.cmb_mode.addItems(["Trajectories (points over time)",
                                "Body lines (head–tail per frame)"])
        self._mode = 0
        self.mode_sel = {0: set(), 1: set()}   # remembered ticks per plot mode
        self.cmb_mode.currentIndexChanged.connect(self._mode_changed)
        mform.addRow(self.cmb_mode)
        self.spn_body_step = QSpinBox()
        self.spn_body_step.setRange(1, 10000)
        self.spn_body_step.setValue(1)
        self.spn_body_step.valueChanged.connect(self.redraw)
        mform.addRow("Body line every N frames:", self.spn_body_step)
        self.chk_body_spine = QCheckBox("Join all points (head–spine 1–2–3–tail)")
        self.chk_body_spine.setToolTip("Off: straight head–tail line. On: line through every body point.")
        self.chk_body_spine.toggled.connect(self.redraw)
        mform.addRow(self.chk_body_spine)
        self.chk_body_head = QCheckBox("Dot at head")
        self.chk_body_head.toggled.connect(self.redraw)
        mform.addRow(self.chk_body_head)
        self.chk_body_tail = QCheckBox("Dot at tail")
        self.chk_body_tail.toggled.connect(self.redraw)
        mform.addRow(self.chk_body_tail)
        self.chk_body_spdots = QCheckBox("Dots at spine points")
        self.chk_body_spdots.toggled.connect(self.redraw)
        mform.addRow(self.chk_body_spdots)
        self.chk_body_momdot = QCheckBox("Dot at mom")
        self.chk_body_momdot.toggled.connect(self.redraw)
        mform.addRow(self.chk_body_momdot)
        hint = QLabel("Tick points in the top box ('Overlay trajectories\non body lines') to draw their paths here.")
        hint.setStyleSheet("color: #666; font-size: 9pt;")
        mform.addRow(hint)
        self.spn_ov_lw = QDoubleSpinBox()
        self.spn_ov_lw.setRange(0.1, 20)
        self.spn_ov_lw.setSingleStep(0.25)
        self.spn_ov_lw.setDecimals(2)
        self.spn_ov_lw.setValue(1.2)
        self.spn_ov_lw.setKeyboardTracking(False)
        self.spn_ov_lw.valueChanged.connect(self.redraw)
        mform.addRow("Overlay line width:", self.spn_ov_lw)
        self.LINE_STYLES = {"Solid": "-", "Dashed": "--", "Dotted": ":",
                            "Dash-dot": "-.", "Long dash": (0, (8, 4)),
                            "Fine dots": (0, (1, 1.5))}
        self.cmb_ov_ls = QComboBox()
        self.cmb_ov_ls.addItems(list(self.LINE_STYLES))
        self.cmb_ov_ls.currentIndexChanged.connect(self.redraw)
        mform.addRow("Overlay line style:", self.cmb_ov_ls)
        self.chk_ov_halo = QCheckBox("White edge around overlay line")
        self.chk_ov_halo.setChecked(True)
        self.chk_ov_halo.toggled.connect(self.redraw)
        mform.addRow(self.chk_ov_halo)
        self.MOM_COLOURS = {
            "Black": "#000000", "Grey": "#808080", "Dark grey": "#404040",
            "Light grey": "#bfbfbf", "White": "#ffffff", "Blue": "#1f77b4",
            "Navy": "#08306b", "Red": "#d62728", "Orange": "#ff7f0e",
            "Magenta": "#e7298a", "Purple": "#7b3294", "Teal": "#1b9e77",
        }
        self.mom_custom = None
        self.cmb_mom_col = QComboBox()
        for name, hexc in self.MOM_COLOURS.items():
            pm = QPixmap(14, 14)
            pm.fill(QColor(hexc))
            self.cmb_mom_col.addItem(QIcon(pm), name)
        self.cmb_mom_col.addItem("Custom…")
        self.cmb_mom_col.activated.connect(self._mom_colour_chosen)
        mform.addRow("Mom line colour:", self.cmb_mom_col)
        self.spn_dot = QDoubleSpinBox()
        self.spn_dot.setRange(0.5, 50)
        self.spn_dot.setSingleStep(0.5)
        self.spn_dot.setDecimals(1)
        self.spn_dot.setValue(3)
        self.spn_dot.setKeyboardTracking(False)
        self.spn_dot.valueChanged.connect(self.redraw)
        mform.addRow("Dot size:", self.spn_dot)
        self.chk_dot_open = QCheckBox("Open circles (white fill)")
        self.chk_dot_open.toggled.connect(self.redraw)
        mform.addRow(self.chk_dot_open)
        self.cmb_dot_col = QComboBox()
        self.cmb_dot_col.addItems(["Head red / tail blue", "Same as line (time)"])
        self.cmb_dot_col.setToolTip("Only matters when Colour = By time")
        self.cmb_dot_col.currentIndexChanged.connect(self.redraw)
        mform.addRow("Dot colour:", self.cmb_dot_col)
        ctrl.addWidget(grp_mode)

        # time conversion
        grp_time = QGroupBox("Time")
        tform = QFormLayout(grp_time)
        self.spn_fps = QDoubleSpinBox()
        self.spn_fps.setRange(0.001, 10000)
        self.spn_fps.setDecimals(3)
        self.spn_fps.setValue(10)
        self.spn_fps.setKeyboardTracking(False)
        self.spn_fps.valueChanged.connect(self.redraw)
        tform.addRow("Frames per second:", self.spn_fps)
        self.cmb_tunit = QComboBox()
        self.cmb_tunit.addItems(["Frames", "Seconds (s)", "Minutes (min)"])
        self.cmb_tunit.setToolTip("Units used for the colour bar, title and first/last labels")
        self.cmb_tunit.currentIndexChanged.connect(self.redraw)
        tform.addRow("Show time as:", self.cmb_tunit)
        self.lbl_time = QLabel("")
        self.lbl_time.setWordWrap(True)
        self.lbl_time.setStyleSheet("color: #333;")
        tform.addRow(self.lbl_time)
        ctrl.addWidget(grp_time)

        # spatial calibration
        grp_cal = QGroupBox("Spatial calibration")
        cform = QFormLayout(grp_cal)
        self.spn_ppcm_x = QDoubleSpinBox()
        self.spn_ppcm_y = QDoubleSpinBox()
        for s in (self.spn_ppcm_x, self.spn_ppcm_y):
            s.setRange(0.0001, 1e6)
            s.setDecimals(4)
            s.setValue(1.0)
            s.setKeyboardTracking(False)
            s.valueChanged.connect(self._cal_changed)
        cform.addRow("x: pixels per cm:", self.spn_ppcm_x)
        cform.addRow("y: pixels per cm:", self.spn_ppcm_y)
        self.chk_ppcm_same = QCheckBox("Use the x factor for y as well")
        self.chk_ppcm_same.toggled.connect(self._cal_changed)
        cform.addRow(self.chk_ppcm_same)
        self.cmb_sunit = QComboBox()
        self.cmb_sunit.addItems(["px", "cm", "mm"])
        self.cmb_sunit.setToolTip("Units for the plotted coordinates and axis labels")
        self.cmb_sunit.currentIndexChanged.connect(self._cal_changed)
        cform.addRow("Show axes in:", self.cmb_sunit)
        self.lbl_cal = QLabel("")
        self.lbl_cal.setStyleSheet("color: #333;")
        self.lbl_cal.setWordWrap(True)
        cform.addRow(self.lbl_cal)
        ctrl.addWidget(grp_cal)

        grp_opt = QGroupBox("Display")
        form = QFormLayout(grp_opt)
        self.cmb_colour = QComboBox()
        self.cmb_colour.addItems(["By body point", "By time (frame)"])
        self.cmb_colour.currentIndexChanged.connect(self.redraw)
        form.addRow("Colour:", self.cmb_colour)

        # single colour for body lines when Colour = By body point
        self.BODY_COLOURS = {"Grey (default)": "#999999", **self.MOM_COLOURS}
        self.body_custom = None
        self.cmb_body_col = QComboBox()
        for name, hexc in self.BODY_COLOURS.items():
            pm = QPixmap(14, 14)
            pm.fill(QColor(hexc))
            self.cmb_body_col.addItem(QIcon(pm), name)
        self.cmb_body_col.addItem("Custom…")
        self.cmb_body_col.setToolTip("Colour of the body lines when Colour = By body point")
        self.cmb_body_col.activated.connect(self._body_colour_chosen)
        form.addRow("Body line colour:", self.cmb_body_col)
        self.spn_body_alpha = QDoubleSpinBox()
        self.spn_body_alpha.setRange(0.05, 1.0)
        self.spn_body_alpha.setSingleStep(0.05)
        self.spn_body_alpha.setDecimals(2)
        self.spn_body_alpha.setValue(0.8)
        self.spn_body_alpha.setKeyboardTracking(False)
        self.spn_body_alpha.setToolTip("Lower = more transparent; helps dense regions read better")
        self.spn_body_alpha.valueChanged.connect(self.redraw)
        form.addRow("Body line opacity:", self.spn_body_alpha)
        self.spn_body_lw = QDoubleSpinBox()
        self.spn_body_lw.setRange(0.05, 10)
        self.spn_body_lw.setSingleStep(0.1)
        self.spn_body_lw.setDecimals(2)
        self.spn_body_lw.setValue(0.6)
        self.spn_body_lw.setKeyboardTracking(False)
        self.spn_body_lw.valueChanged.connect(self.redraw)
        form.addRow("Body line width:", self.spn_body_lw)

        self.spn_start = QSpinBox()
        self.spn_end = QSpinBox()
        for s in (self.spn_start, self.spn_end):
            s.setRange(0, 0)
            s.valueChanged.connect(self.redraw)
        form.addRow("Start frame:", self.spn_start)
        form.addRow("End frame:", self.spn_end)

        self.chk_skel = QCheckBox("Draw body skeleton")
        self.chk_skel.toggled.connect(self.redraw)
        form.addRow(self.chk_skel)
        self.spn_skel = QSpinBox()
        self.spn_skel.setRange(1, 10000)
        self.spn_skel.setValue(100)
        self.spn_skel.valueChanged.connect(self.redraw)
        form.addRow("  every N frames:", self.spn_skel)

        self.chk_markers = QCheckBox("Start / end markers")
        self.chk_markers.setChecked(True)
        self.chk_markers.toggled.connect(self.redraw)
        form.addRow(self.chk_markers)
        self.spn_fl_lw = QDoubleSpinBox()
        self.spn_fl_lw.setRange(0.1, 10)
        self.spn_fl_lw.setSingleStep(0.25)
        self.spn_fl_lw.setDecimals(2)
        self.spn_fl_lw.setValue(1.0)
        self.spn_fl_lw.setKeyboardTracking(False)
        self.spn_fl_lw.setToolTip("Width of the first/last body lines (Body lines mode)")
        self.spn_fl_lw.valueChanged.connect(self.redraw)
        form.addRow("  first/last line width:", self.spn_fl_lw)
        self.spn_fl_ms = QDoubleSpinBox()
        self.spn_fl_ms.setRange(0.5, 30)
        self.spn_fl_ms.setSingleStep(0.5)
        self.spn_fl_ms.setDecimals(1)
        self.spn_fl_ms.setValue(4)
        self.spn_fl_ms.setKeyboardTracking(False)
        self.spn_fl_ms.setToolTip("Size of start/end markers (head dot on first/last body line)")
        self.spn_fl_ms.valueChanged.connect(self.redraw)
        form.addRow("  start/end marker size:", self.spn_fl_ms)

        self.chk_flip = QCheckBox("Image coords (flip y)")
        self.chk_flip.setChecked(True)
        self.chk_flip.toggled.connect(self.redraw)
        form.addRow(self.chk_flip)

        self.cmb_rot = QComboBox()
        self.cmb_rot.addItems(["0°", "90° clockwise", "180°", "90° anticlockwise"])
        self.cmb_rot.setToolTip("Rotates the whole plot (all points) about the centre of the track")
        self.cmb_rot.currentIndexChanged.connect(self.redraw)
        form.addRow("Rotate:", self.cmb_rot)

        self.chk_equal = QCheckBox("Equal aspect")
        self.chk_equal.setChecked(True)
        self.chk_equal.toggled.connect(self.redraw)
        form.addRow(self.chk_equal)

        self.chk_legend = QCheckBox("Legend")
        self.chk_legend.setChecked(True)
        self.chk_legend.toggled.connect(self.redraw)
        form.addRow(self.chk_legend)

        self.chk_box = QCheckBox("Closed box (top/right spines)")
        self.chk_box.setChecked(True)
        self.chk_box.toggled.connect(self.redraw)
        form.addRow(self.chk_box)

        self.chk_axes = QCheckBox("Show axes")
        self.chk_axes.setChecked(True)
        self.chk_axes.toggled.connect(self.redraw)
        form.addRow(self.chk_axes)

        self.spn_lw = QSpinBox()
        self.spn_lw.setRange(1, 10)
        self.spn_lw.setValue(1)
        self.spn_lw.valueChanged.connect(self.redraw)
        form.addRow("Line width:", self.spn_lw)

        self.txt_title = QLineEdit()
        self.txt_title.setPlaceholderText("(file name)")
        self.txt_title.editingFinished.connect(self.redraw)
        form.addRow("Title:", self.txt_title)
        ctrl.addWidget(grp_opt)

        # axis ranges
        grp_rng = self.grp_rng = QGroupBox("Axis range (px)")
        rform = QFormLayout(grp_rng)
        self.chk_auto = QCheckBox("Auto")
        self.chk_auto.setChecked(True)
        self.chk_auto.toggled.connect(self._auto_toggled)
        rform.addRow(self.chk_auto)
        self.rng = {}
        for key in ("x min", "x max", "y min", "y max"):
            s = QDoubleSpinBox()
            s.setRange(-1e6, 1e6)
            s.setDecimals(3)
            s.setSingleStep(10)
            s.setKeyboardTracking(False)
            s.setEnabled(False)
            s.valueChanged.connect(self.redraw)
            rform.addRow(key + ":", s)
            self.rng[key] = s
        btn_view = QPushButton("Use current view")
        btn_view.setToolTip("Copy the current (e.g. zoomed) limits into the boxes")
        btn_view.clicked.connect(self._use_current_view)
        rform.addRow(btn_view)
        ctrl.addWidget(grp_rng)

        btn_pdf = QPushButton("Export PDF…")
        btn_pdf.setStyleSheet("font-weight: bold; padding: 6px;")
        btn_pdf.clicked.connect(self.export_pdf)
        ctrl.addWidget(btn_pdf)
        ctrl.addStretch(1)

        # ---------------- plot ----------------
        right = QVBoxLayout()
        main.addLayout(right, 1)
        self.fig = Figure(figsize=(7, 7), layout="tight")  # re-fits on every resize
        self.canvas = FigureCanvas(self.fig)
        right.addWidget(NavToolbar(self.canvas, self))
        right.addWidget(self.canvas, 1)

        self.lbl_cal.setText("1 px = 10 mm in x and y (set your factors above)")
        if path:
            self.load(path)

    # ---------------- file handling ----------------
    def open_file(self):
        p, _ = QFileDialog.getOpenFileName(
            self, "Open larva file", "", "Excel files (*.xlsx *.xls);;All files (*)")
        if p:
            self.load(p)

    def load(self, path):
        try:
            larvae = load_larva_xlsx(path)
        except Exception as e:
            QMessageBox.critical(self, "Load error", f"Could not read file:\n{e}")
            return
        self.larvae, self.path = larvae, path
        self.cmb_larva.blockSignals(True)
        self.cmb_larva.clear()
        self.cmb_larva.addItems(list(larvae))
        self.cmb_larva.blockSignals(False)
        self.cmb_larva.setVisible(len(larvae) > 1)
        self.lbl_larva.setVisible(len(larvae) > 1)
        self.select_larva(0)

    def select_larva(self, i):
        if i < 0 or not getattr(self, "larvae", None):
            return
        self.larva_id = list(self.larvae)[i]
        pts = self.raw_points = self.points = self.larvae[self.larva_id]
        n = len(next(iter(pts.values())))
        valid = sum(~np.isnan(v).any(axis=1) for v in pts.values()).astype(bool).sum()
        extra = f"\nlarva: {self.larva_id}" if len(self.larvae) > 1 else ""
        self.lbl_file.setText(f"{os.path.basename(self.path)}{extra}\n"
                              f"{n} frames ({valid} tracked), {len(pts)} points: "
                              f"{', '.join(pts)}")

        # rebuild checkboxes
        for cb in self.checks.values():
            cb.deleteLater()
        self.checks = {}
        for name in pts:
            cb = QCheckBox(name)
            cb.setStyleSheet(f"color: {POINT_COLOURS.get(name, 'black')};")
            cb.toggled.connect(self.redraw)
            self.pts_layout.addWidget(cb)
            self.checks[name] = cb
        # defaults: trajectories -> head + tail; body lines -> mom overlay only
        self.mode_sel = {0: {"head", "tail"} & set(pts), 1: {"mom"} & set(pts)}
        self._set_ticks(self.mode_sel[self.cmb_mode.currentIndex()])

        for s in (self.spn_start, self.spn_end):
            s.blockSignals(True)
            s.setRange(0, n - 1)
        self.spn_start.setValue(0)
        self.spn_end.setValue(n - 1)
        for s in (self.spn_start, self.spn_end):
            s.blockSignals(False)
        self.redraw()

    def _set_range_boxes(self, xlim, ylim):
        lo_hi = {"x min": min(xlim), "x max": max(xlim),
                 "y min": min(ylim), "y max": max(ylim)}
        for k, s in self.rng.items():
            s.blockSignals(True)
            s.setValue(round(lo_hi[k], 3))
            s.blockSignals(False)

    def _auto_toggled(self, on):
        for s in self.rng.values():
            s.setEnabled(not on)
        if not on and self.fig.axes:
            # start manual mode from what is currently shown
            ax = self.fig.axes[0]
            self._set_range_boxes(ax.get_xlim(), ax.get_ylim())
        self.redraw()

    def _use_current_view(self):
        if not self.fig.axes:
            return
        ax = self.fig.axes[0]
        self._set_range_boxes(ax.get_xlim(), ax.get_ylim())
        self.chk_auto.blockSignals(True)
        self.chk_auto.setChecked(False)
        self.chk_auto.blockSignals(False)
        for s in self.rng.values():
            s.setEnabled(True)
        self.redraw()

    def _body_colour_chosen(self, i):
        if self.cmb_body_col.itemText(i) == "Custom…":
            c = QColorDialog.getColor(QColor(self.body_custom or "#000000"), self,
                                      "Body line colour")
            if c.isValid():
                self.body_custom = c.name()
                pm = QPixmap(14, 14)
                pm.fill(c)
                self.cmb_body_col.setItemIcon(i, QIcon(pm))
        self.redraw()

    def _body_colour(self):
        name = self.cmb_body_col.currentText()
        if name == "Custom…":
            return self.body_custom or "#000000"
        return self.BODY_COLOURS.get(name, "#999999")

    def _mom_colour_chosen(self, i):
        if self.cmb_mom_col.itemText(i) == "Custom…":
            start = QColor(self.mom_custom or "#000000")
            c = QColorDialog.getColor(start, self, "Mom line colour")
            if c.isValid():
                self.mom_custom = c.name()
                pm = QPixmap(14, 14)
                pm.fill(c)
                self.cmb_mom_col.setItemIcon(i, QIcon(pm))
                self.cmb_mom_col.setItemText(i, "Custom…")
        self.redraw()

    def _pcol(self, name):
        """Plot colour for a body point (mom uses the chosen mom colour)."""
        if name == "mom":
            return self._mom_colour()
        return POINT_COLOURS.get(name, "#000000")

    def _mom_colour(self):
        name = self.cmb_mom_col.currentText()
        if name == "Custom…":
            return self.mom_custom or "#000000"
        return self.MOM_COLOURS.get(name, "#000000")

    def _set_ticks(self, sel):
        for name, cb in self.checks.items():
            cb.blockSignals(True)
            cb.setChecked(sel is None or name in sel)
            cb.blockSignals(False)

    def _mode_changed(self, new):
        # remember ticks of the mode we are leaving, restore those of the new one
        self.mode_sel[self._mode] = set(self.selected())
        self._mode = new
        self._set_ticks(self.mode_sel.get(new, set()))
        self.grp_points.setTitle("Body points" if new == 0
                                 else "Overlay trajectories on body lines")
        self.redraw()

    def apply_preset(self, sel):
        self._set_ticks(sel)
        self.redraw()

    # ---------------- spatial calibration ----------------
    def _scale_factors(self):
        """Multipliers turning pixels into the chosen display unit (x, y)."""
        u = self.cmb_sunit.currentIndex()
        if u == 0:
            return 1.0, 1.0
        fx = self.spn_ppcm_x.value()
        fy = fx if self.chk_ppcm_same.isChecked() else self.spn_ppcm_y.value()
        k = 1.0 if u == 1 else 10.0          # cm or mm
        return k / fx, k / fy

    def _calibrated(self, pts):
        sx, sy = self._scale_factors()
        if sx == 1.0 and sy == 1.0:
            return pts
        return {n: v * np.array([sx, sy]) for n, v in pts.items()}

    def _cal_changed(self, *_):
        self.spn_ppcm_y.setEnabled(not self.chk_ppcm_same.isChecked())
        u = self.cmb_sunit.currentText()
        self.grp_rng.setTitle(f"Axis range ({u})")
        for s in self.rng.values():
            s.setSingleStep({"px": 10, "cm": 0.1, "mm": 1}[u])
        fx = self.spn_ppcm_x.value()
        fy = fx if self.chk_ppcm_same.isChecked() else self.spn_ppcm_y.value()
        self.lbl_cal.setText(f"1 px = {10 / fx:.4g} mm in x, {10 / fy:.4g} mm in y")
        # manual axis limits are in the old unit -> go back to Auto
        if not self.chk_auto.isChecked():
            self.chk_auto.setChecked(True)   # triggers redraw
        else:
            self.redraw()

    # ---------------- rotation ----------------
    def _rotated(self, pts):
        """Rotate all points about the centre of the track's bounding box.
        Directions are as seen on screen, so they respect the y-flip setting."""
        k = self.cmb_rot.currentIndex()
        if k == 0:
            return pts
        allxy = np.vstack(list(pts.values()))
        c = np.array([np.nanmin(allxy[:, 0]) + np.nanmax(allxy[:, 0]),
                      np.nanmin(allxy[:, 1]) + np.nanmax(allxy[:, 1])]) / 2
        # screen y points down when the axis is flipped
        s = 1 if self.chk_flip.isChecked() else -1
        out = {}
        for n, v in pts.items():
            dx, dy = v[:, 0] - c[0], v[:, 1] - c[1]
            if k == 1:      # 90° clockwise on screen
                nx, ny = -s * dy, s * dx
            elif k == 2:    # 180°
                nx, ny = -dx, -dy
            else:           # 90° anticlockwise on screen
                nx, ny = s * dy, -s * dx
            out[n] = np.column_stack([nx + c[0], ny + c[1]])
        return out

    # ---------------- plotting ----------------
    def selected(self):
        return [n for n, cb in self.checks.items() if cb.isChecked()]

    def redraw(self):
        self.fig.clear()
        ax = self.fig.add_subplot(111)
        if not getattr(self, "raw_points", None):
            self.canvas.draw_idle()
            return
        # calibrate first (camera x/y have their own factors), then rotate
        self.points = self._rotated(self._calibrated(self.raw_points))

        a, b = self.spn_start.value(), self.spn_end.value()
        if b < a:
            a, b = b, a
        sl = slice(a, b + 1)
        frames = np.arange(a, b + 1)
        lw = self.spn_lw.value()
        by_time = self.cmb_colour.currentIndex() == 1
        norm = mpl.colors.Normalize(vmin=a, vmax=b)
        cmap = mpl.cm.viridis
        lc_last = None
        body_mode = self.cmb_mode.currentIndex() == 1
        for w in (self.spn_body_step, self.chk_body_spine, self.chk_body_head,
                  self.chk_body_tail, self.chk_body_spdots, self.spn_dot, self.chk_dot_open,
                  self.cmb_dot_col, self.chk_body_momdot,
                  self.spn_ov_lw, self.cmb_ov_ls, self.chk_ov_halo):
            w.setEnabled(body_mode)
        has_mom = "mom" in self.points
        self.chk_body_momdot.setEnabled(body_mode and has_mom)

        if body_mode:
            lc_last = self._draw_body_lines(ax, a, b, lw, by_time, cmap, norm)
            self._draw_skeleton(ax, a, b, by_time, cmap, norm, on_top=True)
            self._finish_axes(ax, a, b, by_time, lc_last, legend_handles=None)
            return

        # skeleton first so trajectories sit on top
        self._draw_skeleton(ax, a, b, by_time, cmap, norm, on_top=False)

        for name in self.selected():
            xy = self.points[name][sl]
            colour = self._pcol(name)
            if by_time:
                segs = np.stack([xy[:-1], xy[1:]], axis=1)
                lc = LineCollection(segs, cmap=cmap, norm=norm, linewidths=lw, zorder=2)
                lc.set_array(frames[:-1])
                ax.add_collection(lc)
                lc_last = lc
                # invisible line for the legend entry
                ax.plot([], [], "-", color="0.3", lw=lw, label=name)
            else:
                ax.plot(xy[:, 0], xy[:, 1], "-", color=colour, lw=lw, label=name, zorder=2)

            if self.chk_markers.isChecked():
                good = np.where(~np.isnan(xy).any(axis=1))[0]
                if len(good):
                    ax.plot(*xy[good[0]], "o", ms=6, mfc="white",
                            mec=colour or "k", mew=1.5, zorder=3)
                    ax.plot(*xy[good[-1]], "s", ms=6, color=colour or "k", zorder=3)

        self._finish_axes(ax, a, b, by_time, lc_last, legend_handles="auto")

    def _draw_skeleton(self, ax, a, b, by_time, cmap, norm, on_top):
        """Head -> spine -> tail every N frames.
        on_top=False: thin, underneath trajectories.
        on_top=True : bold outline drawn over the dense body lines."""
        if not self.chk_skel.isChecked():
            return
        order = [p for p in SKELETON_ORDER if p in self.points]
        if len(order) < 2:
            return
        z = 4 if on_top else 1
        lw = 1.8 if on_top else 0.8
        for f in range(a, b + 1, self.spn_skel.value()):
            xy = np.array([self.points[p][f] for p in order])
            if np.isnan(xy).any():
                continue
            col = cmap(norm(f)) if by_time else "0.55"
            if on_top:
                # dark halo so it stands out from same-coloured body lines
                ax.plot(xy[:, 0], xy[:, 1], "-", color="k", lw=lw + 1.2,
                        solid_capstyle="round", zorder=z)
            ax.plot(xy[:, 0], xy[:, 1], "-", color=col, lw=lw,
                    solid_capstyle="round", alpha=1 if on_top else 0.8, zorder=z)
            if on_top:
                ax.plot(xy[1:-1, 0], xy[1:-1, 1], "o", ms=2.5, mfc="white",
                        mec="k", mew=0.6, zorder=z)
            if order[0] == "head":
                ax.plot(*xy[0], "o", ms=4 if on_top else 2.5,
                        color=col if by_time else POINT_COLOURS["head"],
                        mec="k" if on_top else "none", mew=0.6, zorder=z)

    def _draw_body_lines(self, ax, a, b, lw, by_time, cmap, norm):
        """One line per frame from head to tail; nothing joins frames."""
        if self.chk_body_spine.isChecked():
            order = [p for p in SKELETON_ORDER if p in self.points]
        else:
            order = [p for p in ("head", "tail") if p in self.points]
        if len(order) < 2:
            QMessageBox.warning(self, "Missing points", "Need both head and tail columns.")
            return None
        fr = np.arange(a, b + 1, self.spn_body_step.value())
        segs = np.stack([self.points[p][fr] for p in order], axis=1)  # (F, P, 2)
        ok = ~np.isnan(segs).any(axis=(1, 2))
        segs, fr = segs[ok], fr[ok]
        lc = LineCollection(segs, linewidths=self.spn_body_lw.value(),
                            alpha=self.spn_body_alpha.value(), zorder=2)
        if by_time:
            lc.set_cmap(cmap)
            lc.set_norm(norm)
            lc.set_array(fr)
        else:
            lc.set_color(self._body_colour())
        ax.add_collection(lc)
        # head / tail dots on every body line
        size = self.spn_dot.value() ** 2          # scatter uses area (pt^2)
        open_c = self.chk_dot_open.isChecked()
        dot_sets = []
        if self.chk_body_spdots.isChecked():
            for sp in ("spinepoint_1", "spinepoint_2", "spinepoint_3"):
                if sp in self.points:
                    dot_sets.append(("spine", self.points[sp][fr]))
        if self.chk_body_tail.isChecked():
            dot_sets.append(("tail", segs[:, -1, :]))
        if self.chk_body_momdot.isChecked() and "mom" in self.points:
            dot_sets.append(("mom", self.points["mom"][fr]))
        if self.chk_body_head.isChecked():
            dot_sets.append(("head", segs[:, 0, :]))
        for name, p in dot_sets:
            if not len(p):
                continue
            kw = dict(s=size, zorder=3, linewidths=0.6 if open_c else 0)
            if by_time and self.cmb_dot_col.currentIndex() == 1:
                if open_c:
                    ax.scatter(p[:, 0], p[:, 1], facecolors="white",
                               edgecolors=cmap(norm(fr)), **kw)
                else:
                    ax.scatter(p[:, 0], p[:, 1], c=fr, cmap=cmap, norm=norm, **kw)
            else:
                col = self._pcol(name)
                if open_c:
                    ax.scatter(p[:, 0], p[:, 1], facecolors="white", edgecolors=col, **kw)
                else:
                    ax.scatter(p[:, 0], p[:, 1], color=col, **kw)
        # trajectories of the ticked body points, drawn over the body lines
        mlw = self.spn_ov_lw.value()
        ls = self.LINE_STYLES[self.cmb_ov_ls.currentText()]
        for name in self.selected():
            m = self.points[name][a:b + 1]
            lcol = self._pcol(name)
            if self.chk_ov_halo.isChecked():
                # edge keeps it readable over dense body lines; dark edge for light colours
                r, g, b_ = mpl.colors.to_rgb(lcol)
                edge = "black" if (0.299 * r + 0.587 * g + 0.114 * b_) > 0.7 else "white"
                ax.plot(m[:, 0], m[:, 1], color=edge, ls=ls, lw=mlw + 1.6,
                        zorder=3.4, solid_capstyle="round", dash_capstyle="round")
            ax.plot(m[:, 0], m[:, 1], color=lcol, ls=ls, lw=mlw,
                    zorder=3.5, solid_capstyle="round", dash_capstyle="round",
                    label=f"{name} trajectory")
        if self.chk_markers.isChecked() and len(segs):
            # first and last body line highlighted, head end marked
            flw, fms = self.spn_fl_lw.value(), self.spn_fl_ms.value()
            ax.plot(segs[0][:, 0], segs[0][:, 1], "-", color="k", lw=flw, zorder=4,
                    label=f"first frame ({self._tstr(fr[0])})")
            ax.plot(*segs[0][0], "o", ms=fms, mfc="white", mec="k",
                    mew=max(0.5, flw * 0.75), zorder=5)
            ax.plot(segs[-1][:, 0], segs[-1][:, 1], "--", color="k", lw=flw, zorder=4,
                    label=f"last frame ({self._tstr(fr[-1])})")
            ax.plot(*segs[-1][0], "o", ms=fms, color="k", zorder=5)
        return lc if by_time else None

    # ---------------- time conversion ----------------
    def _tstr(self, frame, unit=True):
        """Frame number formatted in the chosen time unit."""
        k = self.cmb_tunit.currentIndex()
        if k == 0:
            return f"{int(frame)}"
        t = frame / self.spn_fps.value()
        if k == 2:
            t /= 60
        u = " s" if k == 1 else " min"
        return f"{t:.4g}{u if unit else ''}"

    def _update_time_label(self, a, b):
        fps = self.spn_fps.value()
        t0, t1 = a / fps, b / fps
        dur = (b - a) / fps
        self.lbl_time.setText(
            f"Frames {a}–{b} = {t0:.2f}–{t1:.2f} s\n"
            f"Duration: {dur:.2f} s ({dur / 60:.2f} min)")

    def _finish_axes(self, ax, a, b, by_time, lc_last, legend_handles):
        manual = not self.chk_auto.isChecked()
        if manual:
            x0, x1 = self.rng["x min"].value(), self.rng["x max"].value()
            y0, y1 = self.rng["y min"].value(), self.rng["y max"].value()
            if x1 <= x0 or y1 <= y0:
                manual = False   # invalid range -> fall back to auto
        if manual:
            ax.set_xlim(x0, x1)
            ax.set_ylim((y1, y0) if self.chk_flip.isChecked() else (y0, y1))
            if self.chk_equal.isChecked():
                # keep the exact limits; shape the box instead
                ax.set_aspect("equal", adjustable="box")
        else:
            ax.autoscale_view()
            if self.chk_flip.isChecked():
                if not ax.yaxis_inverted():
                    ax.invert_yaxis()
            if self.chk_equal.isChecked():
                ax.set_aspect("equal", adjustable="datalim")

        if self.chk_axes.isChecked():
            u = self.cmb_sunit.currentText()
            ax.set_xlabel(f"x ({u})")
            ax.set_ylabel(f"y ({u})")
            for s in ("top", "right"):
                ax.spines[s].set_visible(self.chk_box.isChecked())
        else:
            ax.set_axis_off()

        title = self.txt_title.text().strip() or os.path.splitext(os.path.basename(self.path))[0]
        if self.cmb_tunit.currentIndex() == 0:
            ax.set_title(f"{title}  (frames {a}–{b})", fontsize=11)
        else:
            ax.set_title(f"{title}  ({self._tstr(a, unit=False)}–{self._tstr(b)})", fontsize=11)
        self._update_time_label(a, b)

        if by_time and lc_last is not None:
            cb = self.fig.colorbar(lc_last, ax=ax, fraction=0.04, pad=0.02)
            k = self.cmb_tunit.currentIndex()
            if k == 0:
                cb.set_label("frame")
            else:
                div = self.spn_fps.value() * (60 if k == 2 else 1)
                # choose round tick values in time units, then place them in frames
                tt = mpl.ticker.MaxNLocator(nbins=7).tick_values(a / div, b / div)
                tt = tt[(tt >= a / div - 1e-9) & (tt <= b / div + 1e-9)]
                cb.set_ticks(tt * div)
                cb.set_ticklabels([f"{t:g}" for t in tt])
                cb.set_label("time (s)" if k == 1 else "time (min)")

        if self.chk_legend.isChecked() and legend_handles is None:
            h, l = ax.get_legend_handles_labels()
            open_c = self.chk_dot_open.isChecked()
            fixed = self.cmb_dot_col.currentIndex() == 0 or not by_time
            for on, name in ((self.chk_body_head.isChecked(), "head"),
                             (self.chk_body_tail.isChecked(), "tail"),
                             (self.chk_body_spdots.isChecked(), "spine"),
                             (self.chk_body_momdot.isChecked() and "mom" in self.points, "mom")):
                if on:
                    col = self._pcol(name) if fixed else "0.4"
                    h.append(mpl.lines.Line2D(
                        [], [], marker="o", ls="", ms=min(self.spn_dot.value() * 1.3, 8),
                        mfc="white" if open_c else col, mec=col))
                    l.append(name)
            if self.chk_markers.isChecked() and not self.chk_body_head.isChecked():
                h.append(mpl.lines.Line2D([], [], marker="o", ls="", color="0.4", ms=4))
                l.append("head end")
            if h:
                ax.legend(h, l, frameon=False, fontsize=8, loc="best")
        elif self.chk_legend.isChecked() and self.selected():
            h, l = ax.get_legend_handles_labels()
            if self.chk_markers.isChecked():
                h += [mpl.lines.Line2D([], [], marker="o", ls="", mfc="white", mec="k"),
                      mpl.lines.Line2D([], [], marker="s", ls="", color="k")]
                l += ["start", "end"]
            ax.legend(h, l, frameon=False, fontsize=8, loc="best")

        self.fig.tight_layout()
        self.canvas.draw_idle()

    # ---------------- export ----------------
    def export_pdf(self):
        if not self.points:
            QMessageBox.information(self, "Nothing to export", "Load a file first.")
            return
        stem = os.path.splitext(os.path.basename(self.path))[0]
        if self.cmb_mode.currentIndex() == 1:
            tag = "bodylines" + ("_spine" if self.chk_body_spine.isChecked() else "_head-tail")
        else:
            tag = "trajectory_" + ("-".join(self.selected()) or "none")
        default = os.path.join(os.path.dirname(self.path), f"{stem}_{tag}.pdf")
        p, _ = QFileDialog.getSaveFileName(self, "Export PDF", default, "PDF (*.pdf)")
        if not p:
            return
        if not p.lower().endswith(".pdf"):
            p += ".pdf"
        self.fig.savefig(p, format="pdf", bbox_inches="tight")
        QMessageBox.information(self, "Exported", f"Saved:\n{p}")


def main():
    app = QApplication(sys.argv)
    w = TrajectoryViewer(sys.argv[1] if len(sys.argv) > 1 else None)
    w.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
