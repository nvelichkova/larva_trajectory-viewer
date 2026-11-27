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

from PyQt5.QtCore import Qt, QSize
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QHBoxLayout, QVBoxLayout, QGroupBox,
    QCheckBox, QPushButton, QFileDialog, QLabel, QSpinBox, QComboBox,
    QFormLayout, QMessageBox, QLineEdit, QDoubleSpinBox, QScrollArea, QColorDialog,
    QGridLayout, QSlider, QSplitter, QAbstractSpinBox,
)
from PyQt5.QtGui import QColor, QIcon, QPixmap, QImage
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
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        scroll.setMinimumWidth(220)
        # splitter: drag the divider to make the control panel wider/narrower
        self.splitter = QSplitter(Qt.Horizontal)
        self.splitter.addWidget(scroll)
        self.splitter.setChildrenCollapsible(False)
        self.splitter.setHandleWidth(7)
        self.splitter.setStyleSheet("QSplitter::handle { background: #d0d0d0; }"
                                    "QSplitter::handle:hover { background: #a0a0a0; }")
        main.addWidget(self.splitter)
        self._panel, self._scroll = panel, scroll

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

        # ---- frame range (which part of the recording to plot) ----
        grp_fr = QGroupBox("Frame range")
        fgrid = QGridLayout(grp_fr)
        self.spn_start = QSpinBox()
        self.spn_end = QSpinBox()
        self.sld_start = QSlider(Qt.Horizontal)
        self.sld_end = QSlider(Qt.Horizontal)
        self.spn_tstart = QDoubleSpinBox()
        self.spn_tend = QDoubleSpinBox()
        for s in (self.spn_start, self.spn_end):
            s.setRange(0, 0)
            s.setKeyboardTracking(False)
        for s in (self.spn_tstart, self.spn_tend):
            s.setRange(0, 0)
            s.setDecimals(2)
            s.setSuffix(" s")
            s.setKeyboardTracking(False)
        for s in (self.sld_start, self.sld_end):
            s.setRange(0, 0)
            # update numbers while dragging, redraw when released
            s.sliderReleased.connect(self.redraw)
        self.spn_start.valueChanged.connect(lambda v: self._set_range(start=v))
        self.spn_end.valueChanged.connect(lambda v: self._set_range(end=v))
        self.sld_start.valueChanged.connect(lambda v: self._set_range(start=v, slider=True))
        self.sld_end.valueChanged.connect(lambda v: self._set_range(end=v, slider=True))
        self.spn_tstart.valueChanged.connect(
            lambda t: self._set_range(start=int(round(t * self.spn_fps.value()))))
        self.spn_tend.valueChanged.connect(
            lambda t: self._set_range(end=int(round(t * self.spn_fps.value()))))
        fgrid.addWidget(QLabel("Start:"), 0, 0)
        fgrid.addWidget(self.spn_start, 0, 1)
        fgrid.addWidget(self.spn_tstart, 0, 2)
        fgrid.addWidget(self.sld_start, 1, 0, 1, 3)
        fgrid.addWidget(QLabel("End:"), 2, 0)
        fgrid.addWidget(self.spn_end, 2, 1)
        fgrid.addWidget(self.spn_tend, 2, 2)
        fgrid.addWidget(self.sld_end, 3, 0, 1, 3)
        btn_all = QPushButton("All frames")
        btn_all.setToolTip("Show the whole recording (within any trim)")
        btn_all.clicked.connect(self._range_all)
        fgrid.addWidget(btn_all, 4, 0, 1, 3)
        self.lbl_range = QLabel("")
        self.lbl_range.setStyleSheet("color: #333;")
        self.lbl_range.setWordWrap(True)
        fgrid.addWidget(self.lbl_range, 5, 0, 1, 3)
        # trim: redefine what counts as "the recording"
        btn_trim = QPushButton("Set as recording")
        btn_trim.setToolTip(
            "Make the current Start–End the recording: frames outside it are dropped.\n"
            "'Whole recording' colours and 'Recording start' time then refer to this,\n"
            "and later frame selections stay inside it and keep its colours.")
        btn_trim.clicked.connect(self._trim_to_range)
        btn_untrim = QPushButton("Full file")
        btn_untrim.setToolTip("Undo the trim: the recording is the whole file again")
        btn_untrim.clicked.connect(self._untrim)
        fgrid.addWidget(btn_trim, 6, 0, 1, 2)
        fgrid.addWidget(btn_untrim, 6, 2)
        self.lbl_trim = QLabel("")
        self.lbl_trim.setStyleSheet("color: #8a4b00;")
        self.lbl_trim.setWordWrap(True)
        fgrid.addWidget(self.lbl_trim, 7, 0, 1, 3)
        fgrid.setColumnStretch(1, 1)
        fgrid.setColumnStretch(2, 1)
        self.spn_start.setToolTip("First frame to plot")
        self.spn_end.setToolTip("Last frame to plot")
        self.spn_tstart.setToolTip("Start as time (uses Frames per second in the Time box)")
        self.spn_tend.setToolTip("End as time (uses Frames per second in the Time box)")
        ctrl.addWidget(grp_fr)

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
        # re-sync the time boxes of the frame range, then redraw
        self.spn_fps.valueChanged.connect(lambda *_: self._set_range())
        tform.addRow("Frames per second:", self.spn_fps)
        self.cmb_tunit = QComboBox()
        self.cmb_tunit.addItems(["Frames", "Seconds (s)", "Minutes (min)"])
        self.cmb_tunit.setToolTip("Units used for the colour bar, title and first/last labels")
        self.cmb_tunit.currentIndexChanged.connect(self.redraw)
        tform.addRow("Show time as:", self.cmb_tunit)
        self.cmb_tzero = QComboBox()
        self.cmb_tzero.addItems(["Start frame (count up from 0)",
                                 "End frame (count down to 0)",
                                 "Recording start (frame 0, or trim start)",
                                 "Recording end (count down to 0)"])
        self.cmb_tzero.setToolTip(
            "Start frame: the current Start frame is time 0 — after cutting frames, the\n"
            "new first frame is 0.\n"
            "End frame: the end of the plotted range is 0 and earlier frames are negative\n"
            "(e.g. −120 s … 0), so everything leads up to the last frame.\n"
            "Recording start: the first frame of the recording is 0 — frame 0 of the file,\n"
            "or the trim start after 'Set as recording'.\n"
            "Recording end: the last frame of the recording (or trim end) is 0; a section\n"
            "keeps its place on that countdown, e.g. −11311 … −5311 s.")
        self.cmb_tzero.currentIndexChanged.connect(self.redraw)
        tform.addRow("Time zero at:", self.cmb_tzero)
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

        # colour map used when Colour = By time (light -> dark options first)
        self.CMAPS = [
            ("Light→dark: Yellow-Green-Blue", "YlGnBu"),
            ("Light→dark: Yellow-Orange-Red", "YlOrRd"),
            ("Light→dark: Magma", "magma_r"),
            ("Light→dark: Inferno", "inferno_r"),
            ("Light→dark: Plasma", "plasma_r"),
            ("Light→dark: Viridis", "viridis_r"),
            ("Light→dark: Cividis", "cividis_r"),
            ("Light→dark: Blues", "Blues"),
            ("Light→dark: Purples", "Purples"),
            ("Light→dark: Greens", "Greens"),
            ("Light→dark: Greys", "Greys"),
            ("Light→dark: Pink-Purple", "RdPu"),
            ("Light→dark: Blue-Purple", "BuPu"),
            ("Dark→light: Viridis (original)", "viridis"),
            ("Dark→light: Magma", "magma"),
            ("Rainbow: Turbo", "turbo"),
        ]
        self.cmb_cmap = QComboBox()
        for label, name in self.CMAPS:
            self.cmb_cmap.addItem(self._cmap_icon(name), label)
        self.cmb_cmap.setIconSize(QSize(60, 12))
        self.cmb_cmap.setCurrentIndex(0)
        self.cmb_cmap.setToolTip("Colour map for Colour = By time (first frame → last frame)")
        self.cmb_cmap.currentIndexChanged.connect(self._cmap_changed)
        form.addRow("Time colour map:", self.cmb_cmap)
        self.chk_cmap_rev = QCheckBox("Reverse colour map")
        self.chk_cmap_rev.toggled.connect(self._cmap_changed)
        form.addRow(self.chk_cmap_rev)
        self.spn_cmap_skip = QSpinBox()
        self.spn_cmap_skip.setRange(0, 60)
        self.spn_cmap_skip.setSuffix(" %")
        self.spn_cmap_skip.setValue(15)
        self.spn_cmap_skip.setKeyboardTracking(False)
        self.spn_cmap_skip.setToolTip("Cut off the palest end of the map so early frames "
                                      "stay visible on a white background")
        self.spn_cmap_skip.valueChanged.connect(self._cmap_changed)
        form.addRow("Skip palest:", self.spn_cmap_skip)
        self.cmb_cscale = QComboBox()
        self.cmb_cscale.addItems(["Start frame → end frame (restarts when cropped)",
                                  "Whole recording (fixed colours)"])
        self.cmb_cscale.setToolTip(
            "Start frame → end frame: the palette (and colour bar) begins at the current\n"
            "Start frame, so after cutting frames the new first frame gets the first colour.\n"
            "Whole recording: every frame keeps the colour it has in the full plot of the\n"
            "recording (the whole file, or the trimmed part after 'Set as recording');\n"
            "the colour bar shows the recording with lines marking the plotted part.")
        self.cmb_cscale.currentIndexChanged.connect(self.redraw)
        form.addRow("Colour scale:", self.cmb_cscale)
        self.chk_cbar_full = QCheckBox("Colour bar shows whole recording")
        self.chk_cbar_full.setToolTip(
            "Only used with Colour scale = Whole recording.\n"
            "Off (default): the colour bar covers just the plotted frames, keeping their\n"
            "whole-recording colours, with ticks relative to Time zero.\n"
            "On: the colour bar covers the whole recording, with lines marking the\n"
            "plotted part.")
        self.chk_cbar_full.toggled.connect(self.redraw)
        form.addRow(self.chk_cbar_full)

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
        right_w = QWidget()
        right = QVBoxLayout(right_w)
        right.setContentsMargins(0, 0, 0, 0)
        self.splitter.addWidget(right_w)
        self.splitter.setStretchFactor(0, 0)
        self.splitter.setStretchFactor(1, 1)
        self.fig = Figure(figsize=(7, 7), layout="tight")  # re-fits on every resize
        self.canvas = FigureCanvas(self.fig)
        right.addWidget(NavToolbar(self.canvas, self))
        right.addWidget(self.canvas, 1)

        self.lbl_cal.setText("1 px = 10 mm in x and y (set your factors above)")
        self._compact_panel()
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

        self.n_frames = n
        self.trim = (0, n - 1)          # the 'recording' (can be trimmed)
        self._update_trim_label()
        for s in (self.spn_start, self.spn_end, self.sld_start, self.sld_end):
            s.blockSignals(True)
            s.setRange(0, n - 1)
            s.blockSignals(False)
        self._set_range(start=0, end=n - 1)

    # ---------------- panel layout ----------------
    def _compact_panel(self):
        """Let the control panel work at narrow widths / high DPI scaling."""
        panel = self._panel
        for form in panel.findChildren(QFormLayout):
            # long label + field: put the field on the next line when narrow
            form.setRowWrapPolicy(QFormLayout.WrapLongRows)
            form.setFieldGrowthPolicy(QFormLayout.AllNonFixedFieldsGrow)
        for cmb in panel.findChildren(QComboBox):
            # don't force the panel as wide as the longest menu entry
            cmb.setSizeAdjustPolicy(QComboBox.AdjustToMinimumContentsLengthWithIcon)
            cmb.setMinimumContentsLength(8)
            cmb.view().setMinimumWidth(cmb.view().sizeHintForColumn(0) + 30)
        for w in panel.findChildren(QAbstractSpinBox):
            w.setMinimumWidth(70)
        for lbl in panel.findChildren(QLabel):
            lbl.setWordWrap(True)
        # starting width: what the panel wants, within sensible limits
        want = min(max(panel.sizeHint().width() + 25, 320), 460)
        self.splitter.setSizes([want, max(self.width() - want, 400)])

    # ---------------- time colour map ----------------
    @staticmethod
    def _cmap_icon(name, w=60, h=12):
        from PyQt5.QtGui import QImage, QColor, QIcon, QPixmap  # local: always available
        cm = mpl.colormaps[name]
        img = QImage(w, h, QImage.Format_RGB32)
        for x in range(w):
            r, g, b, _ = cm(x / (w - 1))
            c = QColor(int(r * 255), int(g * 255), int(b * 255))
            for y in range(h):
                img.setPixelColor(x, y, c)
        return QIcon(QPixmap.fromImage(img))

    def _time_cmap(self):
        """Selected map, optionally reversed, with the palest end trimmed."""
        name = self.CMAPS[self.cmb_cmap.currentIndex()][1]
        cm = mpl.colormaps[name]
        if self.chk_cmap_rev.isChecked():
            cm = cm.reversed()
        skip = self.spn_cmap_skip.value() / 100
        if skip <= 0:
            return cm
        # find which end is lighter and trim that end
        lum = lambda c: 0.299 * c[0] + 0.587 * c[1] + 0.114 * c[2]
        if lum(cm(0.0)) >= lum(cm(1.0)):
            vals = np.linspace(skip, 1, 256)
        else:
            vals = np.linspace(0, 1 - skip, 256)
        return mpl.colors.ListedColormap(cm(vals), name=f"{cm.name}_trim")

    def _cmap_changed(self, *_):
        # switching to a colour map implies time colouring
        if self.cmb_colour.currentIndex() != 1 and self.sender() is self.cmb_cmap:
            self.cmb_colour.setCurrentIndex(1)   # triggers redraw
        else:
            self.redraw()

    # ---------------- frame range ----------------
    def _set_range(self, start=None, end=None, slider=False):
        """Keep frame boxes, time boxes and sliders in sync; start <= end."""
        a = self.spn_start.value() if start is None else start
        b = self.spn_end.value() if end is None else end
        n = getattr(self, "n_frames", 1)
        t0, t1 = getattr(self, "trim", (0, n - 1))
        a = max(t0, min(a, t1))
        b = max(t0, min(b, t1))
        if a > b:                       # push the other end along
            if start is not None:
                b = a
            else:
                a = b
        fps = self.spn_fps.value()
        widgets = (self.spn_start, self.spn_end, self.sld_start, self.sld_end,
                   self.spn_tstart, self.spn_tend)
        for w in widgets:
            w.blockSignals(True)
        self.spn_start.setValue(a)
        self.spn_end.setValue(b)
        self.sld_start.setValue(a)
        self.sld_end.setValue(b)
        for s in (self.spn_start, self.spn_end, self.sld_start, self.sld_end):
            s.setRange(t0, t1)
        for s in (self.spn_tstart, self.spn_tend):
            s.setRange(t0 / fps, t1 / fps)
        self.spn_tstart.setValue(a / fps)
        self.spn_tend.setValue(b / fps)
        for w in widgets:
            w.blockSignals(False)
        self.lbl_range.setText(f"{b - a + 1} of {t1 - t0 + 1} frames\n"
                               f"{(b - a) / fps:.2f} s ({(b - a) / fps / 60:.2f} min)")
        dragging = slider and (self.sld_start.isSliderDown() or self.sld_end.isSliderDown())
        if not dragging:
            self.redraw()

    def _range_all(self):
        t0, t1 = getattr(self, "trim", (0, getattr(self, "n_frames", 1) - 1))
        self._set_range(start=t0, end=t1)

    def _trim_to_range(self):
        self.trim = (self.spn_start.value(), self.spn_end.value())
        self._update_trim_label()
        self._set_range()

    def _untrim(self):
        self.trim = (0, self.n_frames - 1)
        self._update_trim_label()
        self._set_range()

    def _update_trim_label(self):
        t0, t1 = self.trim
        if (t0, t1) == (0, self.n_frames - 1):
            self.lbl_trim.setText("Recording = full file")
        else:
            self.lbl_trim.setText(f"Recording trimmed to frames {t0}–{t1} "
                                  f"(of 0–{self.n_frames - 1})")

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
        # frame that counts as time 0 (first frame of file, or end of plotted range)
        # 0 start frame, 1 end frame, 2 recording start, 3 recording end
        self._t_mode = self.cmb_tzero.currentIndex()
        self._t_ref = {0: a, 1: b, 2: self.trim[0], 3: self.trim[1]}[self._t_mode]
        sl = slice(a, b + 1)
        frames = np.arange(a, b + 1)
        lw = self.spn_lw.value()
        by_time = self.cmb_colour.currentIndex() == 1
        if self.cmb_cscale.currentIndex() == 1:
            # fixed to the whole file, so cropped plots share the same colours
            t0, t1 = self.trim
            norm = mpl.colors.Normalize(vmin=t0, vmax=max(t1, t0 + 1))
        else:
            norm = mpl.colors.Normalize(vmin=a, vmax=b)
        cmap = self._time_cmap()
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
        rel = frame - getattr(self, "_t_ref", 0)
        if k == 0:
            return f"{int(rel)}".replace("-", "−")
        t = rel / self.spn_fps.value()
        if k == 2:
            t /= 60
        u = " s" if k == 1 else " min"
        return f"{self._num(t)}{u if unit else ''}"

    @staticmethod
    def _num(t):
        """Readable number without scientific notation (−17310, −12.5, −0.333)."""
        a = abs(t)
        if a >= 100:
            txt = f"{t:.0f}"
        elif a >= 10:
            txt = f"{t:.1f}"
        else:
            txt = f"{t:.3g}"
        if "." in txt:
            txt = txt.rstrip("0").rstrip(".")
        if txt in ("-0", ""):
            txt = "0"
        return txt.replace("-", "−")

    def _update_time_label(self, a, b):
        fps = self.spn_fps.value()
        t0, t1 = a / fps, b / fps
        dur = (b - a) / fps
        txt = (f"Frames {a}–{b} = {t0:.2f}–{t1:.2f} s\n"
               f"Duration: {dur:.2f} s ({dur / 60:.2f} min)")
        mode = getattr(self, "_t_mode", 0)
        if mode == 0 and a > 0:
            txt += f"\nTime 0 = start frame {a}"
        elif mode == 1:
            txt += f"\nTime 0 = end frame {b}; first plotted frame = −{dur:.2f} s"
        elif mode == 2 and self.trim[0] > 0:
            txt += f"\nTime 0 = recording start, frame {self.trim[0]}"
        elif mode == 3:
            r = self.trim[1]
            txt += (f"\nTime 0 = recording end, frame {r}; plotted "
                    f"{self._num((a - r) / fps)} to {self._num((b - r) / fps)} s")
        self.lbl_time.setText(txt)

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
        sep = " to " if self._t_mode in (1, 3) else "–"
        if self.cmb_tunit.currentIndex() == 0:
            rng = f"frames {self._tstr(a)}{sep}{self._tstr(b)}"
            if self._t_ref:     # relative numbering: also give the real frame
                rng += f"; 0 = frame {self._t_ref}"
        else:
            rng = f"{self._tstr(a, unit=False)}{sep}{self._tstr(b)}"
        ax.set_title(f"{title}  ({rng})", fontsize=11)
        self._update_time_label(a, b)

        if by_time and lc_last is not None:
            whole = self.cmb_cscale.currentIndex() == 1
            cropped = a > self.trim[0] or b < self.trim[1]
            if whole and cropped and not self.chk_cbar_full.isChecked() and b > a:
                # colour bar for the plotted frames only, but with the colours those
                # frames have on the whole-recording scale
                fr_ = np.linspace(a, b, 256)
                sub = mpl.colors.ListedColormap(lc_last.cmap(lc_last.norm(fr_)))
                sm = mpl.cm.ScalarMappable(norm=mpl.colors.Normalize(a, b), cmap=sub)
                sm.set_array([])
                cb = self.fig.colorbar(sm, ax=ax, fraction=0.04, pad=0.02)
                cb_lo, cb_hi = a, b
            else:
                cb = self.fig.colorbar(lc_last, ax=ax, fraction=0.04, pad=0.02)
                cb_lo, cb_hi = lc_last.norm.vmin, lc_last.norm.vmax
                if whole and cropped:
                    # show which part of the recording is plotted
                    for f in (a, b):
                        cb.ax.axhline(f, color="k", lw=1.2)
            k = self.cmb_tunit.currentIndex()
            ref = self._t_ref
            div = 1.0 if k == 0 else self.spn_fps.value() * (60 if k == 2 else 1)
            # choose round tick values in the shown unit (relative to time 0),
            # then place them at the matching frames
            lo, hi = cb_lo, cb_hi   # range of the colour bar
            rlo, rhi = (lo - ref) / div, (hi - ref) / div
            loc = mpl.ticker.MaxNLocator(nbins=7, integer=(k == 0))
            tt = loc.tick_values(rlo, rhi)
            tt = tt[(tt >= rlo - 1e-9) & (tt <= rhi + 1e-9)]
            cb.set_ticks(tt * div + ref)
            cb.set_ticklabels([self._num(t) for t in tt])
            unit = {0: "frame", 1: "time (s)", 2: "time (min)"}[k]
            note = {0: ", 0 = start frame", 1: ", 0 = last frame",
                    2: ", 0 = recording start", 3: ", 0 = recording end"}.get(self._t_mode, "")
            cb.set_label(unit + (note if ref or self._t_mode in (1, 3) else ""))

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
