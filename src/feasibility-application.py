from PyQt5 import QtWidgets, QtGui, QtCore
import sys
import json
import csv
import numpy as np
from datetime import datetime
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors as rl_colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, HRFlowable
from reportlab.lib.units import cm
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
from matplotlib.backends.backend_qt5agg import NavigationToolbar2QT as NavigationToolbar

# ── Stylesheet ────────────────────────────────────────────────────────────────
APP_STYLE = """
QMainWindow { background: #f0f2f5; }
QTabWidget::pane { border:1px solid #c8cdd3; border-radius:4px; background:white; }
QTabBar::tab {
    background:#e1e4e8; border:1px solid #c8cdd3; padding:8px 18px; margin-right:2px;
    border-top-left-radius:4px; border-top-right-radius:4px; font-weight:bold; color:#555; font-size:12px;
}
QTabBar::tab:selected { background:white; color:#1a73e8; border-bottom-color:white; }
QTabBar::tab:hover:!selected { background:#d0d4da; }
QGroupBox {
    font-weight:bold; border:2px solid #d0d4da; border-radius:6px;
    margin-top:12px; padding-top:8px; background:#fafbfc;
}
QGroupBox::title { subcontrol-origin:margin; left:10px; padding:0 5px; color:#1a73e8; }
QPushButton {
    background:#1a73e8; color:white; border:none; padding:7px 16px;
    border-radius:4px; font-weight:bold; min-width:80px; font-size:12px;
}
QPushButton:hover { background:#1558b0; }
QPushButton:pressed { background:#0d47a1; }
QPushButton#dangerBtn  { background:#d32f2f; }
QPushButton#dangerBtn:hover  { background:#b71c1c; }
QPushButton#successBtn { background:#388e3c; }
QPushButton#successBtn:hover { background:#2e7d32; }
QPushButton#calcBtn {
    background:qlineargradient(x1:0,y1:0,x2:1,y2:0,stop:0 #1a73e8,stop:1 #0d47a1);
    font-size:14px; padding:10px 24px; border-radius:6px;
}
QLineEdit, QDoubleSpinBox, QComboBox {
    border:1px solid #c8cdd3; border-radius:4px; padding:5px 8px;
    background:white; font-size:13px;
}
QLineEdit:focus, QDoubleSpinBox:focus { border:2px solid #1a73e8; }
QTableWidget {
    border:1px solid #d0d4da; border-radius:4px; gridline-color:#e8eaed;
    font-size:12px; alternate-background-color:#f8f9fa;
}
QTableWidget::item:selected { background:#e3f2fd; color:#0d47a1; }
QHeaderView::section {
    background:#f1f3f4; border:none; border-bottom:2px solid #d0d4da;
    padding:6px; font-weight:bold; color:#444; font-size:12px;
}
QLabel { color:#333; }
QStatusBar { background:#f1f3f4; border-top:1px solid #d0d4da; color:#666; font-size:12px; }
QMenuBar { background:#f1f3f4; border-bottom:1px solid #d0d4da; font-size:13px; }
QMenuBar::item:selected { background:#e3f2fd; color:#1a73e8; }
QMenu::item { padding:6px 24px; }
QMenu::item:selected { background:#e3f2fd; color:#1a73e8; }
QScrollArea { border:none; }
"""


class ImproveWindow(QtWidgets.QMainWindow):
    def __init__(self, master):
        super().__init__()
        self.master = master
        self.setWindowTitle("Economic Feasibility Analyser – Improve Project")
        self.setGeometry(100, 80, 1280, 860)
        self.setMinimumSize(1000, 700)

        self.exchange_rates    = {"USD_to_EGP": 49.41, "EUR_to_EGP": 57.26}
        self.fixed_costs_data  = []
        self.running_costs_data = []
        self.savings_data      = []
        self.economic_data     = None

        self.setStyleSheet(APP_STYLE)
        self._init_menu()
        self._init_statusbar()

        self.tabs = QtWidgets.QTabWidget()
        self.setCentralWidget(self.tabs)

        self.init_analysis_tab()
        self.init_upgrade_tab()
        self.init_running_tab()
        self.init_savings_tab()
        self.init_results_tab()
        self.init_graphs_tab()

    # ── Menu & Status ─────────────────────────────────────────────────────────
    def _init_menu(self):
        mb = self.menuBar()
        fm = mb.addMenu("&File")
        self._act(fm, "New Project",    self.new_project,  "Ctrl+N")
        self._act(fm, "Open Project…",  self.load_project, "Ctrl+O")
        self._act(fm, "Save Project…",  self.save_project, "Ctrl+S")
        fm.addSeparator()
        self._act(fm, "Export PDF…",    self.export_pdf,   "Ctrl+P")
        self._act(fm, "Export CSV…",    self.export_csv,   "Ctrl+E")
        fm.addSeparator()
        self._act(fm, "Exit",           self.close,        "Ctrl+Q")
        hm = mb.addMenu("&Help")
        self._act(hm, "About",          self.show_about)

    def _act(self, menu, text, slot, shortcut=None):
        a = QtWidgets.QAction(text, self)
        if shortcut:
            a.setShortcut(shortcut)
        a.triggered.connect(slot)
        menu.addAction(a)

    def _init_statusbar(self):
        self.status_bar = self.statusBar()
        self.status_bar.showMessage("Ready  •  Fill in project data then click Calculate")

    # ── Tab 1: Project Basics ─────────────────────────────────────────────────
    def init_analysis_tab(self):
        self.analysis_tab = QtWidgets.QWidget()
        outer = QtWidgets.QVBoxLayout()
        scroll = QtWidgets.QScrollArea()
        scroll.setWidgetResizable(True)
        ctr = QtWidgets.QWidget()
        lay = QtWidgets.QVBoxLayout(ctr)
        lay.setSpacing(14)

        # Project info
        g = QtWidgets.QGroupBox("Project Information")
        f = QtWidgets.QFormLayout(); f.setSpacing(10)
        self.project_name = QtWidgets.QLineEdit()
        self.project_name.setPlaceholderText("e.g. Pump Efficiency Upgrade – Phase 2")
        self.project_desc = QtWidgets.QLineEdit()
        self.project_desc.setPlaceholderText("Brief description or location")
        f.addRow("Project Name:", self.project_name)
        f.addRow("Description:",  self.project_desc)
        g.setLayout(f); lay.addWidget(g)

        # Exchange rates
        g2 = QtWidgets.QGroupBox("Exchange Rates")
        f2 = QtWidgets.QFormLayout(); f2.setSpacing(10)
        dv = QtGui.QDoubleValidator(0.01, 999999, 4)
        self.usd_input = QtWidgets.QLineEdit("49.41"); self.usd_input.setValidator(dv)
        self.usd_input.setToolTip("How many EGP equals 1 USD")
        self.usd_input.textChanged.connect(self._on_fx_changed)
        self.eur_input = QtWidgets.QLineEdit("57.26"); self.eur_input.setValidator(dv)
        self.eur_input.setToolTip("How many EGP equals 1 EUR")
        self.eur_input.textChanged.connect(self._on_fx_changed)
        f2.addRow("1 USD = ? EGP:", self.usd_input)
        f2.addRow("1 EUR = ? EGP:", self.eur_input)
        g2.setLayout(f2); lay.addWidget(g2)

        # Current system
        g3 = QtWidgets.QGroupBox("Current System")
        f3 = QtWidgets.QFormLayout(); f3.setSpacing(10)
        self.curr_value = QtWidgets.QLineEdit()
        self.curr_value.setPlaceholderText("Replacement value in USD")
        self.curr_value.setValidator(QtGui.QDoubleValidator(0, 1e12, 2))
        self.curr_perf = QtWidgets.QLineEdit()
        self.curr_perf.setPlaceholderText("Current annual revenue / output in USD")
        self.curr_perf.setValidator(QtGui.QDoubleValidator(-1e12, 1e12, 2))
        f3.addRow("Current System Value ($):", self.curr_value)
        f3.addRow("Current Annual Performance ($):", self.curr_perf)
        g3.setLayout(f3); lay.addWidget(g3)

        # Improvement
        g4 = QtWidgets.QGroupBox("Improvement Parameters")
        f4 = QtWidgets.QFormLayout(); f4.setSpacing(10)
        self.budget = QtWidgets.QLineEdit()
        self.budget.setPlaceholderText("Capital budget in USD")
        self.budget.setValidator(QtGui.QDoubleValidator(0, 1e12, 2))
        self.impr_perf = QtWidgets.QLineEdit()
        self.impr_perf.setPlaceholderText("Expected annual performance after improvement in USD")
        self.impr_perf.setValidator(QtGui.QDoubleValidator(-1e12, 1e12, 2))
        f4.addRow("Improvement Budget ($):", self.budget)
        f4.addRow("Expected Performance After Improvement ($):", self.impr_perf)
        g4.setLayout(f4); lay.addWidget(g4)

        # Analysis params
        g5 = QtWidgets.QGroupBox("Analysis Parameters")
        f5 = QtWidgets.QFormLayout(); f5.setSpacing(10)
        self.years_input = QtWidgets.QLineEdit("5")
        self.years_input.setValidator(QtGui.QIntValidator(1, 50))
        self.years_input.setToolTip("Analysis horizon: 1–50 years")
        self.discount_input = QtWidgets.QLineEdit("24.50")
        self.discount_input.setValidator(QtGui.QDoubleValidator(0, 100, 2))
        self.discount_input.setToolTip("WACC or required rate of return (%)")
        f5.addRow("Analysis Period (years):", self.years_input)
        f5.addRow("Discount Rate (%):", self.discount_input)
        g5.setLayout(f5); lay.addWidget(g5)

        lay.addStretch()
        scroll.setWidget(ctr)
        outer.addWidget(scroll)
        self.analysis_tab.setLayout(outer)
        self.tabs.addTab(self.analysis_tab, "Project Basics")

    def _on_fx_changed(self):
        try:
            self.exchange_rates["USD_to_EGP"] = float(self.usd_input.text())
            self.exchange_rates["EUR_to_EGP"] = float(self.eur_input.text())
            self.update_fixed_totals()
            self.update_running_totals()
            self.update_savings_totals()
        except (ValueError, ZeroDivisionError):
            pass

    # ── Tab 2: Upgrade Costs ──────────────────────────────────────────────────
    def init_upgrade_tab(self):
        self.upgrade_tab = QtWidgets.QWidget()
        lay = QtWidgets.QVBoxLayout(); lay.setSpacing(10)

        fg = QtWidgets.QGroupBox("Add Upgrade / CapEx Item")
        fl = QtWidgets.QHBoxLayout(); fl.setSpacing(8)
        self.fixed_name = QtWidgets.QLineEdit(); self.fixed_name.setPlaceholderText("Item name")
        self.fixed_qty  = QtWidgets.QLineEdit("1")
        self.fixed_qty.setValidator(QtGui.QDoubleValidator(0, 1e9, 4)); self.fixed_qty.setFixedWidth(70)
        self.fixed_unit_cost = QtWidgets.QLineEdit(); self.fixed_unit_cost.setPlaceholderText("Unit cost")
        self.fixed_unit_cost.setValidator(QtGui.QDoubleValidator(0, 1e12, 2))
        self.fixed_currency = QtWidgets.QComboBox(); self.fixed_currency.addItems(["USD", "EGP", "EUR"])
        ab = QtWidgets.QPushButton("Add"); ab.setObjectName("successBtn"); ab.clicked.connect(self.add_upgrade)
        for lbl, w in [("Item:", self.fixed_name), ("Qty:", self.fixed_qty),
                       ("Unit Cost:", self.fixed_unit_cost), ("Currency:", self.fixed_currency)]:
            fl.addWidget(QtWidgets.QLabel(lbl)); fl.addWidget(w)
        fl.addWidget(ab)
        fg.setLayout(fl); lay.addWidget(fg)

        self.fixed_table = QtWidgets.QTableWidget(0, 5)
        self.fixed_table.setHorizontalHeaderLabels(["Item", "Qty", "Unit Cost", "Currency", "Total"])
        self.fixed_table.horizontalHeader().setStretchLastSection(True)
        self.fixed_table.setAlternatingRowColors(True)
        self.fixed_table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        lay.addWidget(self.fixed_table)

        foot = QtWidgets.QHBoxLayout()
        self.fixed_total_label = QtWidgets.QLabel("Total Upgrade Cost (USD): $0.00")
        self.fixed_total_label.setStyleSheet("font-weight:bold; color:#1a73e8; font-size:13px;")
        db = QtWidgets.QPushButton("Delete Selected"); db.setObjectName("dangerBtn"); db.clicked.connect(self.del_upgrade)
        foot.addWidget(self.fixed_total_label); foot.addStretch(); foot.addWidget(db)
        lay.addLayout(foot)

        self.upgrade_tab.setLayout(lay)
        self.tabs.addTab(self.upgrade_tab, "Upgrade Costs")

    # ── Tab 3: Operating Costs ────────────────────────────────────────────────
    def init_running_tab(self):
        self.running_tab = QtWidgets.QWidget()
        lay = QtWidgets.QVBoxLayout(); lay.setSpacing(10)

        pg = QtWidgets.QGroupBox("Standard Operating Costs (EGP/day)")
        pl = QtWidgets.QVBoxLayout()
        self.predefined_costs = [
            {"name": "Petrol 92 Consumption",  "qty": 1,  "unit_cost": 13.75, "unit": "liter/hour",   "hours": 16},
            {"name": "Petrol Gas Consumption", "qty": 1,  "unit_cost": 11.5,  "unit": "liter/hour",   "hours": 16},
            {"name": "Power Consumption",      "qty": 2,  "unit_cost": 1.25,  "unit": "KW/hour",      "hours": 16},
            {"name": "Air Consumption",        "qty": 1,  "unit_cost": 0.21,  "unit": "m3/hour",      "hours": 16},
            {"name": "Gas Consumption",        "qty": 1,  "unit_cost": 2.7,   "unit": "m3/hour",      "hours": 16},
            {"name": "Needed Labor",           "qty": 1,  "unit_cost": 8500,  "unit": "labour/month", "hours": 8},
            {"name": "Needed Area",            "qty": 24, "unit_cost": 1.0,   "unit": "m2/month",     "hours": 8},
        ]
        self.fixed_costs_table = QtWidgets.QTableWidget(len(self.predefined_costs), 6)
        self.fixed_costs_table.setHorizontalHeaderLabels(
            ["Item", "Qty", "Cost/Unit", "Unit", "Hrs/Day", "Daily Total (EGP)"])
        self.fixed_costs_table.horizontalHeader().setStretchLastSection(True)
        self.fixed_costs_table.setAlternatingRowColors(True)
        for i, item in enumerate(self.predefined_costs):
            self.fixed_costs_table.setItem(i, 0, QtWidgets.QTableWidgetItem(item["name"]))
            self.fixed_costs_table.setItem(i, 3, QtWidgets.QTableWidgetItem(item["unit"]))
            for col, key, hi in [(1, "qty", 99999), (2, "unit_cost", 99999), (4, "hours", 24)]:
                sb = QtWidgets.QDoubleSpinBox()
                sb.setRange(0, hi); sb.setValue(item[key]); sb.setDecimals(2)
                sb.valueChanged.connect(self.update_running_totals)
                self.fixed_costs_table.setCellWidget(i, col, sb)
            daily = item["qty"] * item["unit_cost"] * item["hours"]
            lbl = QtWidgets.QLabel(f"{daily:,.2f}"); lbl.setAlignment(QtCore.Qt.AlignCenter)
            self.fixed_costs_table.setCellWidget(i, 5, lbl)
        pl.addWidget(self.fixed_costs_table)
        pg.setLayout(pl); lay.addWidget(pg)

        cg = QtWidgets.QGroupBox("Additional Custom Operating Costs")
        cl = QtWidgets.QVBoxLayout()
        fl2 = QtWidgets.QHBoxLayout(); fl2.setSpacing(8)
        self.run_name = QtWidgets.QLineEdit(); self.run_name.setPlaceholderText("Item name")
        self.run_qty  = QtWidgets.QLineEdit("1")
        self.run_qty.setValidator(QtGui.QDoubleValidator(0, 1e9, 4)); self.run_qty.setFixedWidth(70)
        self.run_unit_cost = QtWidgets.QLineEdit(); self.run_unit_cost.setPlaceholderText("Unit cost")
        self.run_unit_cost.setValidator(QtGui.QDoubleValidator(0, 1e12, 2))
        self.run_hours = QtWidgets.QLineEdit("8")
        self.run_hours.setValidator(QtGui.QDoubleValidator(0, 24, 2)); self.run_hours.setFixedWidth(60)
        self.run_currency = QtWidgets.QComboBox(); self.run_currency.addItems(["USD", "EGP", "EUR"])
        ab2 = QtWidgets.QPushButton("Add"); ab2.setObjectName("successBtn"); ab2.clicked.connect(self.add_running)
        for lbl, w in [("Item:", self.run_name), ("Qty:", self.run_qty),
                       ("Unit Cost:", self.run_unit_cost), ("Hrs/Day:", self.run_hours),
                       ("Currency:", self.run_currency)]:
            fl2.addWidget(QtWidgets.QLabel(lbl)); fl2.addWidget(w)
        fl2.addWidget(ab2)

        self.running_table = QtWidgets.QTableWidget(0, 6)
        self.running_table.setHorizontalHeaderLabels(
            ["Item", "Qty", "Unit Cost", "Hrs/Day", "Currency", "Daily Total"])
        self.running_table.setAlternatingRowColors(True)
        self.running_table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        cl.addLayout(fl2); cl.addWidget(self.running_table)

        foot2 = QtWidgets.QHBoxLayout()
        self.running_total_label = QtWidgets.QLabel("Total Daily Operating Cost (USD): $0.00")
        self.running_total_label.setStyleSheet("font-weight:bold; color:#1a73e8; font-size:13px;")
        db2 = QtWidgets.QPushButton("Delete Selected"); db2.setObjectName("dangerBtn"); db2.clicked.connect(self.del_running)
        foot2.addWidget(self.running_total_label); foot2.addStretch(); foot2.addWidget(db2)
        cl.addLayout(foot2)
        cg.setLayout(cl); lay.addWidget(cg)

        self.running_tab.setLayout(lay)
        self.tabs.addTab(self.running_tab, "Operating Costs")

    # ── Tab 4: Savings ────────────────────────────────────────────────────────
    def init_savings_tab(self):
        self.savings_tab = QtWidgets.QWidget()
        lay = QtWidgets.QVBoxLayout(); lay.setSpacing(10)

        pg = QtWidgets.QGroupBox("Standard Savings (EGP/day) — set qty to 0 if not applicable")
        pl = QtWidgets.QVBoxLayout()
        self.predefined_savings = [
            {"name": "Petrol 92 Saving",  "qty": 1,  "unit_cost": 13.75, "unit": "liter/hour",   "hours": 16},
            {"name": "Petrol Gas Saving", "qty": 1,  "unit_cost": 11.5,  "unit": "liter/hour",   "hours": 16},
            {"name": "Power Saving",      "qty": 2,  "unit_cost": 1.25,  "unit": "KW/hour",      "hours": 16},
            {"name": "Air Saving",        "qty": 1,  "unit_cost": 0.21,  "unit": "m3/hour",      "hours": 16},
            {"name": "Gas Saving",        "qty": 1,  "unit_cost": 2.7,   "unit": "m3/hour",      "hours": 16},
            {"name": "Labor Saving",      "qty": 1,  "unit_cost": 8500,  "unit": "labour/month", "hours": 8},
            {"name": "Space Saving",      "qty": 24, "unit_cost": 1.0,   "unit": "m2/month",     "hours": 8},
        ]
        self.fixed_savings_table = QtWidgets.QTableWidget(len(self.predefined_savings), 6)
        self.fixed_savings_table.setHorizontalHeaderLabels(
            ["Item", "Qty", "Cost/Unit", "Unit", "Hrs/Day", "Daily Saving (EGP)"])
        self.fixed_savings_table.horizontalHeader().setStretchLastSection(True)
        self.fixed_savings_table.setAlternatingRowColors(True)
        for i, item in enumerate(self.predefined_savings):
            self.fixed_savings_table.setItem(i, 0, QtWidgets.QTableWidgetItem(item["name"]))
            self.fixed_savings_table.setItem(i, 3, QtWidgets.QTableWidgetItem(item["unit"]))
            for col, key, hi in [(1, "qty", 99999), (2, "unit_cost", 99999), (4, "hours", 24)]:
                sb = QtWidgets.QDoubleSpinBox()
                sb.setRange(0, hi); sb.setValue(item[key]); sb.setDecimals(2)
                sb.valueChanged.connect(self.update_savings_totals)
                self.fixed_savings_table.setCellWidget(i, col, sb)
            daily = item["qty"] * item["unit_cost"] * item["hours"]
            lbl = QtWidgets.QLabel(f"{daily:,.2f}"); lbl.setAlignment(QtCore.Qt.AlignCenter)
            self.fixed_savings_table.setCellWidget(i, 5, lbl)
        pl.addWidget(self.fixed_savings_table)
        pg.setLayout(pl); lay.addWidget(pg)

        cg = QtWidgets.QGroupBox("Additional Custom Annual Savings")
        cl = QtWidgets.QVBoxLayout()
        fl3 = QtWidgets.QHBoxLayout(); fl3.setSpacing(8)
        self.save_desc   = QtWidgets.QLineEdit(); self.save_desc.setPlaceholderText("Description")
        self.save_amount = QtWidgets.QLineEdit(); self.save_amount.setPlaceholderText("Annual amount")
        self.save_amount.setValidator(QtGui.QDoubleValidator(0, 1e12, 2))
        self.save_currency = QtWidgets.QComboBox(); self.save_currency.addItems(["USD", "EGP", "EUR"])
        ab3 = QtWidgets.QPushButton("Add"); ab3.setObjectName("successBtn"); ab3.clicked.connect(self.add_saving)
        for lbl, w in [("Description:", self.save_desc),
                       ("Annual Amount:", self.save_amount), ("Currency:", self.save_currency)]:
            fl3.addWidget(QtWidgets.QLabel(lbl)); fl3.addWidget(w)
        fl3.addWidget(ab3)

        self.savings_table = QtWidgets.QTableWidget(0, 3)
        self.savings_table.setHorizontalHeaderLabels(["Description", "Annual Amount", "Currency"])
        self.savings_table.setAlternatingRowColors(True)
        self.savings_table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        cl.addLayout(fl3); cl.addWidget(self.savings_table)

        foot3 = QtWidgets.QHBoxLayout()
        self.savings_total_label = QtWidgets.QLabel("Total Annual Savings (USD): $0.00")
        self.savings_total_label.setStyleSheet("font-weight:bold; color:#388e3c; font-size:13px;")
        db3 = QtWidgets.QPushButton("Delete Selected"); db3.setObjectName("dangerBtn"); db3.clicked.connect(self.del_saving)
        foot3.addWidget(self.savings_total_label); foot3.addStretch(); foot3.addWidget(db3)
        cl.addLayout(foot3)
        cg.setLayout(cl); lay.addWidget(cg)

        self.savings_tab.setLayout(lay)
        self.tabs.addTab(self.savings_tab, "Savings")

    # ── Tab 5: Results ────────────────────────────────────────────────────────
    def init_results_tab(self):
        self.results_tab = QtWidgets.QWidget()
        lay = QtWidgets.QVBoxLayout(); lay.setSpacing(12)

        self.calc_btn = QtWidgets.QPushButton("  Calculate Improvement Metrics")
        self.calc_btn.setObjectName("calcBtn")
        self.calc_btn.clicked.connect(self.calculate_metrics)
        lay.addWidget(self.calc_btn)

        # Metric cards
        mg = QtWidgets.QGroupBox("Financial Metrics")
        ml = QtWidgets.QGridLayout(); ml.setSpacing(10)
        self.metric_labels = {}
        cards = [
            ("total_investment",   "Total Investment",   "#1565c0"),
            ("annual_running_cost","Annual Running Cost","#c62828"),
            ("annual_savings",     "Annual Savings",     "#2e7d32"),
            ("net_annual_benefit", "Net Annual Benefit", "#4527a0"),
            ("roi",                "ROI",                "#e65100"),
            ("npv",                "NPV",                "#00695c"),
            ("irr",                "IRR",                "#1a237e"),
            ("payback_period",     "Payback Period",     "#37474f"),
        ]
        defaults = ["$0.00","$0.00","$0.00","$0.00","0.00%","$0.00","0.00%","N/A"]
        for idx, ((key, name, color), default) in enumerate(zip(cards, defaults)):
            card = QtWidgets.QFrame()
            card.setStyleSheet(f"QFrame{{background:white;border:1px solid #e0e0e0;"
                               f"border-radius:8px;border-left:5px solid {color};}}")
            cl2 = QtWidgets.QVBoxLayout(card); cl2.setSpacing(2); cl2.setContentsMargins(10,8,10,8)
            tl = QtWidgets.QLabel(name); tl.setStyleSheet("color:#666;font-size:11px;font-weight:bold;")
            vl = QtWidgets.QLabel(default); vl.setStyleSheet(f"color:{color};font-size:16px;font-weight:bold;")
            cl2.addWidget(tl); cl2.addWidget(vl)
            self.metric_labels[key] = vl
            ml.addWidget(card, idx // 4, idx % 4)
        mg.setLayout(ml); lay.addWidget(mg)

        # Cash flow table with cumulative column
        cfg = QtWidgets.QGroupBox("Cash Flow Analysis")
        cfl = QtWidgets.QVBoxLayout()
        self.cashflow_table = QtWidgets.QTableWidget(0, 6)
        self.cashflow_table.setHorizontalHeaderLabels(
            ["Year","Investment ($)","Running Costs ($)","Savings & Benefits ($)",
             "Net Cash Flow ($)","Cumulative ($)"])
        self.cashflow_table.horizontalHeader().setStretchLastSection(True)
        self.cashflow_table.setAlternatingRowColors(True)
        cfl.addWidget(self.cashflow_table)
        cfg.setLayout(cfl); lay.addWidget(cfg)

        # Recommendation
        self.recommendation = QtWidgets.QLabel("Run calculation to see recommendation.")
        self.recommendation.setWordWrap(True)
        self.recommendation.setStyleSheet(
            "font-weight:bold;font-size:13px;padding:12px;"
            "background:#f8f9fa;border-radius:6px;border:1px solid #e0e0e0;")
        lay.addWidget(self.recommendation)

        # Export buttons
        er = QtWidgets.QHBoxLayout()
        pb = QtWidgets.QPushButton("Export PDF Report"); pb.setObjectName("successBtn"); pb.clicked.connect(self.export_pdf)
        cb = QtWidgets.QPushButton("Export CSV"); cb.clicked.connect(self.export_csv)
        er.addStretch(); er.addWidget(pb); er.addWidget(cb)
        lay.addLayout(er)

        self.results_tab.setLayout(lay)
        self.tabs.addTab(self.results_tab, "Results")

    # ── Tab 6: Graphs ─────────────────────────────────────────────────────────
    def init_graphs_tab(self):
        self.graphs_tab = QtWidgets.QWidget()
        lay = QtWidgets.QVBoxLayout()
        self.figure = Figure(figsize=(12, 9), dpi=100)
        self.canvas = FigureCanvas(self.figure)
        self.toolbar = NavigationToolbar(self.canvas, self)
        lay.addWidget(self.toolbar); lay.addWidget(self.canvas)
        self.graphs_tab.setLayout(lay)
        self.tabs.addTab(self.graphs_tab, "Financial Graphs")

    # ── Upgrade helpers ───────────────────────────────────────────────────────
    def add_upgrade(self):
        try:
            name = self.fixed_name.text().strip()
            if not name:
                QtWidgets.QMessageBox.warning(self, "Input Error", "Please enter an item name."); return
            qty  = float(self.fixed_qty.text())
            cost = float(self.fixed_unit_cost.text())
            cur  = self.fixed_currency.currentText()
            tot  = qty * cost
            self.fixed_costs_data.append({"product":name,"qty":qty,"unit_cost":cost,"total":tot,"currency":cur})
            r = self.fixed_table.rowCount(); self.fixed_table.insertRow(r)
            for c, v in enumerate([name, f"{qty:g}", f"{cost:,.2f}", cur, f"{tot:,.2f}"]):
                self.fixed_table.setItem(r, c, QtWidgets.QTableWidgetItem(v))
            self.update_fixed_totals()
            self.fixed_name.clear(); self.fixed_qty.setText("1"); self.fixed_unit_cost.clear()
        except ValueError:
            QtWidgets.QMessageBox.warning(self, "Input Error", "Enter valid numbers for quantity and unit cost.")

    def del_upgrade(self):
        rows = {i.row() for i in self.fixed_table.selectedItems()}
        if not rows or not self._confirm_del(len(rows)): return
        for r in sorted(rows, reverse=True):
            self.fixed_table.removeRow(r)
            if r < len(self.fixed_costs_data): del self.fixed_costs_data[r]
        self.update_fixed_totals()

    def update_fixed_totals(self):
        u2e = self.exchange_rates["USD_to_EGP"]; e2e = self.exchange_rates["EUR_to_EGP"]
        t = sum(self._to_usd(i["total"], i["currency"]) for i in self.fixed_costs_data)
        self.fixed_total_label.setText(f"Total Upgrade Cost (USD): ${t:,.2f}  ≈  EGP {t*u2e:,.0f}")

    # ── Running cost helpers ──────────────────────────────────────────────────
    def add_running(self):
        try:
            name = self.run_name.text().strip()
            if not name:
                QtWidgets.QMessageBox.warning(self, "Input Error", "Please enter an item name."); return
            qty   = float(self.run_qty.text())
            cost  = float(self.run_unit_cost.text())
            hrs   = float(self.run_hours.text())
            cur   = self.run_currency.currentText()
            tot   = qty * cost * hrs  # daily
            self.running_costs_data.append(
                {"product":name,"qty":qty,"unit_cost":cost,"hours_per_day":hrs,"total":tot,"currency":cur})
            r = self.running_table.rowCount(); self.running_table.insertRow(r)
            for c, v in enumerate([name, f"{qty:g}", f"{cost:,.2f}", f"{hrs:g}", cur, f"{tot:,.2f}"]):
                self.running_table.setItem(r, c, QtWidgets.QTableWidgetItem(v))
            self.update_running_totals()
            self.run_name.clear(); self.run_qty.setText("1"); self.run_unit_cost.clear(); self.run_hours.setText("8")
        except ValueError:
            QtWidgets.QMessageBox.warning(self, "Input Error", "Enter valid numbers for running cost.")

    def del_running(self):
        rows = {i.row() for i in self.running_table.selectedItems()}
        if not rows or not self._confirm_del(len(rows)): return
        for r in sorted(rows, reverse=True):
            self.running_table.removeRow(r)
            if r < len(self.running_costs_data): del self.running_costs_data[r]
        self.update_running_totals()

    def update_running_totals(self):
        u2e = self.exchange_rates["USD_to_EGP"]
        total = 0.0
        for i in range(self.fixed_costs_table.rowCount()):
            qw = self.fixed_costs_table.cellWidget(i,1); cw = self.fixed_costs_table.cellWidget(i,2)
            hw = self.fixed_costs_table.cellWidget(i,4); lw = self.fixed_costs_table.cellWidget(i,5)
            if qw and cw and hw:
                d = qw.value() * cw.value() * hw.value()
                lw.setText(f"{d:,.2f}"); total += d / u2e
        for item in self.running_costs_data:
            total += self._to_usd(item["total"], item["currency"])
        self.running_total_label.setText(
            f"Total Daily Operating Cost (USD): ${total:,.2f}  ≈  EGP {total*u2e:,.0f}"
            f"  |  Annual ≈ ${total*365:,.0f}")

    # ── Savings helpers ───────────────────────────────────────────────────────
    def add_saving(self):
        try:
            desc = self.save_desc.text().strip()
            amt  = float(self.save_amount.text())
            cur  = self.save_currency.currentText()
            self.savings_data.append({"description":desc,"amount":amt,"currency":cur})
            r = self.savings_table.rowCount(); self.savings_table.insertRow(r)
            for c, v in enumerate([desc, f"{amt:,.2f}", cur]):
                self.savings_table.setItem(r, c, QtWidgets.QTableWidgetItem(v))
            self.update_savings_totals()
            self.save_desc.clear(); self.save_amount.clear()
        except ValueError:
            QtWidgets.QMessageBox.warning(self, "Input Error", "Enter a valid savings amount.")

    def del_saving(self):
        rows = {i.row() for i in self.savings_table.selectedItems()}
        if not rows or not self._confirm_del(len(rows)): return
        for r in sorted(rows, reverse=True):
            self.savings_table.removeRow(r)
            if r < len(self.savings_data): del self.savings_data[r]
        self.update_savings_totals()

    def update_savings_totals(self):
        u2e = self.exchange_rates["USD_to_EGP"]
        total = 0.0
        for i in range(self.fixed_savings_table.rowCount()):
            qw = self.fixed_savings_table.cellWidget(i,1); cw = self.fixed_savings_table.cellWidget(i,2)
            hw = self.fixed_savings_table.cellWidget(i,4); lw = self.fixed_savings_table.cellWidget(i,5)
            if qw and cw and hw:
                d = qw.value() * cw.value() * hw.value()
                lw.setText(f"{d:,.2f}"); total += d / u2e * 365
        for item in self.savings_data:
            total += self._to_usd(item["amount"], item["currency"])
        self.savings_total_label.setText(
            f"Total Annual Savings (USD): ${total:,.2f}  ≈  EGP {total*u2e:,.0f}")

    # ── Data readers ──────────────────────────────────────────────────────────
    def get_fixed_costs_data(self):
        out = []
        for i in range(self.fixed_costs_table.rowCount()):
            n = self.fixed_costs_table.item(i,0).text()
            q = self.fixed_costs_table.cellWidget(i,1).value()
            c = self.fixed_costs_table.cellWidget(i,2).value()
            u = self.fixed_costs_table.item(i,3).text()
            h = self.fixed_costs_table.cellWidget(i,4).value()
            out.append({"product":n,"qty":q,"unit_cost":c,"unit":u,
                        "hours_per_day":h,"total":q*c*h,"currency":"EGP"})
        return out

    def get_fixed_savings_data(self):
        out = []
        for i in range(self.fixed_savings_table.rowCount()):
            n = self.fixed_savings_table.item(i,0).text()
            q = self.fixed_savings_table.cellWidget(i,1).value()
            c = self.fixed_savings_table.cellWidget(i,2).value()
            u = self.fixed_savings_table.item(i,3).text()
            h = self.fixed_savings_table.cellWidget(i,4).value()
            out.append({"product":n,"qty":q,"unit_cost":c,"unit":u,
                        "hours_per_day":h,"total":q*c*h,"currency":"EGP"})
        return out

    def _to_usd(self, amount, currency):
        r = self.exchange_rates
        if currency == "EGP": return amount / r["USD_to_EGP"]
        if currency == "EUR": return amount * r["EUR_to_EGP"] / r["USD_to_EGP"]
        return amount

    def _confirm_del(self, n):
        return QtWidgets.QMessageBox.question(
            self, "Confirm Delete", f"Delete {n} selected row(s)?",
            QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No) == QtWidgets.QMessageBox.Yes

    # ── Calculation ───────────────────────────────────────────────────────────
    def calculate_metrics(self):
        try:
            u2e = float(self.usd_input.text()); e2e = float(self.eur_input.text())
            self.exchange_rates["USD_to_EGP"] = u2e; self.exchange_rates["EUR_to_EGP"] = e2e

            curr_val  = float(self.curr_value.text())  if self.curr_value.text()  else 0.0
            inv       = float(self.budget.text())      if self.budget.text()      else 0.0
            curr_perf = float(self.curr_perf.text())   if self.curr_perf.text()   else 0.0
            impr_perf = float(self.impr_perf.text())   if self.impr_perf.text()   else 0.0
            yrs       = int(self.years_input.text())
            disc      = float(self.discount_input.text()) / 100.0

            # Upgrade costs
            fixed_total_usd = sum(self._to_usd(i["total"], i["currency"]) for i in self.fixed_costs_data)

            # Running costs — both custom and predefined are DAILY; annualise once
            custom_daily_usd = sum(self._to_usd(i["total"], i["currency"]) for i in self.running_costs_data)
            fvc = self.get_fixed_costs_data()
            predef_daily_usd = sum(i["total"] / u2e for i in fvc)   # EGP/day → USD/day
            annual_running_cost = (custom_daily_usd + predef_daily_usd) * 365

            # Savings — custom is annual, predefined is daily → annualise
            custom_annual_usd = sum(self._to_usd(i["amount"], i["currency"]) for i in self.savings_data)
            fvs = self.get_fixed_savings_data()
            predef_annual_usd = sum(i["total"] / u2e * 365 for i in fvs)
            annual_savings = custom_annual_usd + predef_annual_usd

            total_investment   = inv + fixed_total_usd
            perf_delta         = impr_perf - curr_perf
            net_annual_benefit = perf_delta - annual_running_cost + annual_savings

            roi = ((net_annual_benefit * yrs - total_investment) / total_investment * 100
                   if total_investment > 0 else 0.0)

            cash_flows = [-total_investment] + [net_annual_benefit] * yrs
            npv  = sum(cf / (1 + disc) ** t for t, cf in enumerate(cash_flows))
            irr  = self._calc_irr(cash_flows)
            payback = self._calc_payback(cash_flows, total_investment, net_annual_benefit, yrs)

            # Update metric cards
            self.metric_labels["total_investment"].setText(f"${total_investment:,.2f}")
            self.metric_labels["annual_running_cost"].setText(f"${annual_running_cost:,.2f}")
            self.metric_labels["annual_savings"].setText(f"${annual_savings:,.2f}")
            self.metric_labels["net_annual_benefit"].setText(f"${net_annual_benefit:,.2f}")
            self.metric_labels["roi"].setText(f"{roi:.2f}%")
            self.metric_labels["npv"].setText(f"${npv:,.2f}")
            self.metric_labels["irr"].setText(f"{irr:.2f}%" if isinstance(irr, float) else str(irr))
            self.metric_labels["payback_period"].setText(
                f"{payback:.2f} years" if isinstance(payback, float) else str(payback))

            # Colour NPV / net benefit green or red
            for key, val in [("npv", npv), ("net_annual_benefit", net_annual_benefit)]:
                c = "#2e7d32" if val >= 0 else "#c62828"
                self.metric_labels[key].setStyleSheet(f"color:{c};font-size:16px;font-weight:bold;")

            # Cash flow table (6 columns including Cumulative)
            self.cashflow_table.setRowCount(0)
            cumul = 0.0
            for yr in range(yrs + 1):
                if yr == 0:
                    inv_v, run_v, sav_v, net_v = total_investment, 0.0, 0.0, -total_investment
                else:
                    inv_v, run_v = 0.0, annual_running_cost
                    sav_v = annual_savings + perf_delta
                    net_v = net_annual_benefit
                cumul += net_v
                ri = self.cashflow_table.rowCount(); self.cashflow_table.insertRow(ri)
                for ci, (v, is_money) in enumerate(
                        [(yr, False), (inv_v, True), (run_v, True),
                         (sav_v, True), (net_v, True), (cumul, True)]):
                    it = QtWidgets.QTableWidgetItem(f"{v:,.2f}" if is_money else str(v))
                    it.setTextAlignment(QtCore.Qt.AlignRight | QtCore.Qt.AlignVCenter)
                    if ci == 4:
                        it.setForeground(QtGui.QColor("#2e7d32" if net_v >= 0 else "#c62828"))
                    if ci == 5:
                        it.setForeground(QtGui.QColor("#2e7d32" if cumul >= 0 else "#c62828"))
                    self.cashflow_table.setItem(ri, ci, it)
            self.cashflow_table.resizeColumnsToContents()

            # Recommendation
            if npv > 0:
                rec = "RECOMMEND IMPROVEMENT — Project has positive NPV and is financially viable"
                if isinstance(irr, float) and irr > disc * 100:
                    rec += f" | IRR {irr:.2f}% > hurdle rate {disc*100:.1f}%"
                if isinstance(payback, float):
                    rec += f" | Payback in {payback:.1f} years"
                self.recommendation.setText("✅  " + rec)
                self.recommendation.setStyleSheet(
                    "font-weight:bold;font-size:13px;padding:12px;color:#1b5e20;"
                    "background:#e8f5e9;border-radius:6px;border:1px solid #a5d6a7;")
            else:
                rec = "DO NOT IMPROVE — Project does not meet financial criteria"
                if isinstance(payback, float) and payback > yrs:
                    rec += f" | Payback {payback:.1f} yrs > analysis period {yrs} yrs"
                self.recommendation.setText("⚠️  " + rec)
                self.recommendation.setStyleSheet(
                    "font-weight:bold;font-size:13px;padding:12px;color:#b71c1c;"
                    "background:#ffebee;border-radius:6px;border:1px solid #ef9a9a;")

            self.update_graphs(cash_flows, yrs, npv, irr, payback,
                               total_investment, annual_running_cost, annual_savings)

            self.economic_data = {
                "project_type": "Improve",
                "project_name": self.project_name.text(),
                "project_desc": self.project_desc.text(),
                "calc_date": datetime.now().strftime("%Y-%m-%d %H:%M"),
                "exchange_rates": dict(self.exchange_rates),
                "current_value": curr_val, "investment": inv,
                "current_performance": curr_perf, "improved_performance": impr_perf,
                "years": yrs, "discount_rate": disc * 100,
                "fixed_costs": list(self.fixed_costs_data),
                "running_costs": list(self.running_costs_data),
                "savings": list(self.savings_data),
                "fixed_variables_costs": fvc, "fixed_variables_savings": fvs,
                "total_investment": total_investment, "annual_running_cost": annual_running_cost,
                "annual_savings": annual_savings, "net_annual_benefit": net_annual_benefit,
                "estimated_roi": roi, "npv": npv,
                "irr": irr if isinstance(irr, float) else 0.0,
                "payback_period": payback if isinstance(payback, float) else 0.0,
                "recommendation": rec, "cash_flows": cash_flows,
            }
            if self.master is not None:
                self.master.economic_data = self.economic_data

            now = datetime.now().strftime("%H:%M:%S")
            self.status_bar.showMessage(
                f"Calculation complete  {now}  |  NPV: ${npv:,.0f}  |  "
                f"ROI: {roi:.1f}%  |  {'Viable' if npv > 0 else 'Not viable'}")

        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Calculation Error", str(e))
            import traceback; traceback.print_exc()

    @staticmethod
    def _calc_irr(cash_flows):
        try:
            rates = np.linspace(0.001, 5.0, 5000)
            npvs  = np.array([sum(cf / (1 + r)**t for t, cf in enumerate(cash_flows)) for r in rates])
            sc = np.where(np.diff(np.sign(npvs)))[0]
            if len(sc) == 0: return "N/A"
            i = sc[0]
            r0, r1, v0, v1 = rates[i], rates[i+1], npvs[i], npvs[i+1]
            return round((r0 - v0 * (r1 - r0) / (v1 - v0)) * 100, 4)
        except Exception:
            return "N/A"

    @staticmethod
    def _calc_payback(cash_flows, total_investment, net_annual_benefit, yrs):
        if net_annual_benefit <= 0: return "N/A"
        cumul = 0.0
        for yr in range(1, yrs + 1):
            cumul += net_annual_benefit
            if cumul >= total_investment:
                prior = cumul - net_annual_benefit
                return round(yr - 1 + (total_investment - prior) / net_annual_benefit, 4)
        return "N/A"

    # ── Graphs ────────────────────────────────────────────────────────────────
    def update_graphs(self, cash_flows, years, npv, irr, payback,
                      total_investment, annual_running_cost, annual_savings):
        self.figure.clear()
        self.figure.patch.set_facecolor("#f8f9fa")
        gs = self.figure.add_gridspec(2, 2, hspace=0.42, wspace=0.35)
        yrs_list = list(range(years + 1))

        import matplotlib.ticker as mticker
        fmt_k = mticker.FuncFormatter(lambda v, _: f"${v/1e3:.0f}k" if abs(v) >= 1000 else f"${v:.0f}")

        # 1 – Annual Cash Flows
        ax1 = self.figure.add_subplot(gs[0, 0])
        colors_bar = ["#c62828" if cf < 0 else "#2e7d32" for cf in cash_flows]
        ax1.bar(yrs_list, cash_flows, color=colors_bar, edgecolor="white", linewidth=0.5)
        ax1.axhline(0, color="#333", linewidth=0.8)
        ax1.set_title("Annual Cash Flows", fontweight="bold", fontsize=11)
        ax1.set_xlabel("Year"); ax1.set_ylabel("USD")
        ax1.yaxis.set_major_formatter(fmt_k)
        ax1.grid(axis="y", alpha=0.3); ax1.set_facecolor("white")

        # 2 – NPV Sensitivity
        ax2 = self.figure.add_subplot(gs[0, 1])
        irr_val = irr if isinstance(irr, float) else None
        max_r = (irr_val / 100 + 0.3) if irr_val else 0.6
        rates = np.linspace(0.001, max_r, 200)
        npvs  = [sum(cf / (1 + r)**t for t, cf in enumerate(cash_flows)) for r in rates]
        ax2.plot(rates*100, npvs, color="#1a73e8", linewidth=2)
        ax2.fill_between(rates*100, npvs, 0,
            where=[v >= 0 for v in npvs], alpha=0.15, color="#2e7d32")
        ax2.fill_between(rates*100, npvs, 0,
            where=[v < 0 for v in npvs], alpha=0.15, color="#c62828")
        if irr_val:
            ax2.axvline(irr_val, color="#c62828", linestyle="--", lw=1.5, label=f"IRR={irr_val:.1f}%")
            ax2.legend(fontsize=9)
        ax2.axhline(0, color="#333", linewidth=0.8)
        ax2.set_title("NPV Sensitivity", fontweight="bold", fontsize=11)
        ax2.set_xlabel("Discount Rate (%)"); ax2.set_ylabel("NPV ($)")
        ax2.yaxis.set_major_formatter(fmt_k)
        ax2.grid(alpha=0.3); ax2.set_facecolor("white")

        # 3 – Cumulative Cash Flow
        ax3 = self.figure.add_subplot(gs[1, 0])
        cumulative = list(np.cumsum(cash_flows))
        ax3.plot(yrs_list, cumulative, color="#1a73e8", linewidth=2, zorder=3)
        scatter_c = ["#2e7d32" if v >= 0 else "#c62828" for v in cumulative]
        ax3.scatter(yrs_list, cumulative, color=scatter_c, zorder=4, s=45)
        ax3.fill_between(yrs_list, cumulative, 0,
            where=[v >= 0 for v in cumulative], alpha=0.12, color="#2e7d32")
        ax3.fill_between(yrs_list, cumulative, 0,
            where=[v < 0 for v in cumulative], alpha=0.12, color="#c62828")
        if isinstance(payback, float):
            ax3.axvline(payback, color="#e65100", linestyle="--", lw=1.5, label=f"Payback={payback:.1f}yr")
            ax3.legend(fontsize=9)
        ax3.axhline(0, color="#333", linewidth=0.8)
        ax3.set_title("Cumulative Cash Flow", fontweight="bold", fontsize=11)
        ax3.set_xlabel("Year"); ax3.set_ylabel("Cumulative ($)")
        ax3.yaxis.set_major_formatter(fmt_k)
        ax3.grid(alpha=0.3); ax3.set_facecolor("white")

        # 4 – Investment Breakdown pie
        ax4 = self.figure.add_subplot(gs[1, 1])
        lbls, vals, pie_c = [], [], []
        for lbl, val, col in [
            ("Capital Investment", total_investment,   "#1a73e8"),
            ("Annual OpEx",        annual_running_cost,"#c62828"),
            ("Annual Savings",     annual_savings,     "#2e7d32"),
        ]:
            if val > 0:
                lbls.append(lbl); vals.append(val); pie_c.append(col)
        if vals:
            _, _, autotexts = ax4.pie(
                vals, labels=lbls, colors=pie_c, autopct="%1.1f%%",
                startangle=140, pctdistance=0.75,
                wedgeprops=dict(edgecolor="white", linewidth=1.5))
            for at in autotexts: at.set_fontsize(9)
        ax4.set_title("Cost & Benefit Breakdown", fontweight="bold", fontsize=11)

        self.canvas.draw()
        self.tabs.setCurrentWidget(self.graphs_tab)

    # ── Export PDF ────────────────────────────────────────────────────────────
    def export_pdf(self):
        if not self.economic_data:
            QtWidgets.QMessageBox.warning(self, "No Data", "Please run the calculation first."); return
        path, _ = QtWidgets.QFileDialog.getSaveFileName(
            self, "Save PDF Report", "feasibility_report.pdf", "PDF Files (*.pdf)")
        if not path: return
        try:
            d = self.economic_data
            doc = SimpleDocTemplate(path, pagesize=A4,
                leftMargin=2*cm, rightMargin=2*cm, topMargin=2.5*cm, bottomMargin=2*cm)
            styles = getSampleStyleSheet()
            title_s = ParagraphStyle("T", parent=styles["Title"], fontSize=18,
                textColor=rl_colors.HexColor("#1a73e8"), spaceAfter=6)
            h1_s = ParagraphStyle("H1", parent=styles["Heading1"], fontSize=13,
                textColor=rl_colors.HexColor("#1a73e8"), spaceBefore=14, spaceAfter=4)
            body_s = styles["Normal"]; body_s.fontSize = 10

            story = [
                Paragraph("Economic Feasibility Report", title_s),
                Paragraph(f"<b>Project:</b> {d.get('project_name','—')}", body_s),
                Paragraph(f"<b>Description:</b> {d.get('project_desc','—')}", body_s),
                Paragraph(f"<b>Generated:</b> {d.get('calc_date','—')}", body_s),
                HRFlowable(width="100%", thickness=1, color=rl_colors.HexColor("#1a73e8")),
                Spacer(1, 0.3*cm),
                Paragraph("1. Project Parameters", h1_s),
                self._pdf_table([
                    ["Parameter", "Value"],
                    ["Current System Value",        f"${d['current_value']:,.2f}"],
                    ["Current Annual Performance",  f"${d['current_performance']:,.2f}"],
                    ["Improvement Budget (CapEx)",  f"${d['investment']:,.2f}"],
                    ["Expected Annual Performance", f"${d['improved_performance']:,.2f}"],
                    ["Analysis Period",             f"{d['years']} years"],
                    ["Discount Rate",               f"{d['discount_rate']:.2f}%"],
                    ["USD → EGP",                   str(d['exchange_rates']['USD_to_EGP'])],
                    ["EUR → EGP",                   str(d['exchange_rates']['EUR_to_EGP'])],
                ]),
                Spacer(1, 0.4*cm),
                Paragraph("2. Financial Metrics", h1_s),
                self._pdf_table([
                    ["Metric", "Value"],
                    ["Total Investment",    f"${d['total_investment']:,.2f}"],
                    ["Annual Running Cost", f"${d['annual_running_cost']:,.2f}"],
                    ["Annual Savings",      f"${d['annual_savings']:,.2f}"],
                    ["Net Annual Benefit",  f"${d['net_annual_benefit']:,.2f}"],
                    ["ROI",                 f"{d['estimated_roi']:.2f}%"],
                    ["NPV",                 f"${d['npv']:,.2f}"],
                    ["IRR",                 f"{d['irr']:.2f}%" if d['irr'] else "N/A"],
                    ["Payback Period",      f"{d['payback_period']:.2f} years" if d['payback_period'] else "N/A"],
                ]),
                Spacer(1, 0.4*cm),
                Paragraph("3. Cash Flow Analysis", h1_s),
            ]

            cf_rows = [["Year","Investment ($)","Running Costs ($)","Savings ($)","Net CF ($)","Cumulative ($)"]]
            cumul = 0.0
            net_ann = d["net_annual_benefit"]; perf_d = d["improved_performance"] - d["current_performance"]
            for yr in range(d["years"] + 1):
                if yr == 0:
                    iv, rv, sv, nv = d["total_investment"], 0, 0, -d["total_investment"]
                else:
                    iv, rv = 0, d["annual_running_cost"]; sv = d["annual_savings"] + perf_d; nv = net_ann
                cumul += nv
                cf_rows.append([str(yr), f"{iv:,.0f}", f"{rv:,.0f}", f"{sv:,.0f}", f"{nv:,.0f}", f"{cumul:,.0f}"])
            story.append(self._pdf_table(cf_rows))
            story.append(Spacer(1, 0.4*cm))
            story.append(Paragraph("4. Recommendation", h1_s))
            rc = rl_colors.HexColor("#1b5e20" if d["npv"] > 0 else "#b71c1c")
            story.append(Paragraph(d["recommendation"],
                ParagraphStyle("Rec", parent=body_s, textColor=rc, fontSize=11, fontName="Helvetica-Bold")))

            doc.build(story)
            self.status_bar.showMessage(f"PDF exported: {path}")
            QtWidgets.QMessageBox.information(self, "Export Complete", f"PDF report saved to:\n{path}")
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "PDF Error", str(e))

    @staticmethod
    def _pdf_table(data):
        t = Table(data, repeatRows=1)
        t.setStyle(TableStyle([
            ("BACKGROUND",  (0,0),(-1,0), rl_colors.HexColor("#1a73e8")),
            ("TEXTCOLOR",   (0,0),(-1,0), rl_colors.white),
            ("FONTNAME",    (0,0),(-1,0), "Helvetica-Bold"),
            ("FONTSIZE",    (0,0),(-1,-1),9),
            ("ALIGN",       (0,0),(-1,-1),"CENTER"),
            ("ALIGN",       (0,1),(0,-1), "LEFT"),
            ("ROWBACKGROUNDS",(0,1),(-1,-1),[rl_colors.white,rl_colors.HexColor("#f1f3f4")]),
            ("GRID",        (0,0),(-1,-1),0.5,rl_colors.HexColor("#d0d4da")),
            ("TOPPADDING",  (0,0),(-1,-1),5),
            ("BOTTOMPADDING",(0,0),(-1,-1),5),
        ]))
        return t

    # ── Export CSV ────────────────────────────────────────────────────────────
    def export_csv(self):
        if not self.economic_data:
            QtWidgets.QMessageBox.warning(self, "No Data", "Please run the calculation first."); return
        path, _ = QtWidgets.QFileDialog.getSaveFileName(
            self, "Save CSV", "cashflow.csv", "CSV Files (*.csv)")
        if not path: return
        try:
            d = self.economic_data
            with open(path, "w", newline="", encoding="utf-8") as f:
                w = csv.writer(f)
                w.writerow(["Economic Feasibility – Cash Flow"])
                w.writerow(["Project", d.get("project_name","—")])
                w.writerow(["Generated", d.get("calc_date","—")])
                w.writerow([])
                w.writerow(["Year","Investment","Running Costs","Savings & Benefits","Net CF","Cumulative"])
                cumul = 0.0
                net_ann = d["net_annual_benefit"]; perf_d = d["improved_performance"] - d["current_performance"]
                for yr in range(d["years"] + 1):
                    if yr == 0:
                        iv, rv, sv, nv = d["total_investment"], 0, 0, -d["total_investment"]
                    else:
                        iv, rv = 0, d["annual_running_cost"]; sv = d["annual_savings"]+perf_d; nv = net_ann
                    cumul += nv
                    w.writerow([yr, f"{iv:.2f}", f"{rv:.2f}", f"{sv:.2f}", f"{nv:.2f}", f"{cumul:.2f}"])
                w.writerow([]); w.writerow(["Key Metrics"])
                for k, v in [("NPV", f"${d['npv']:,.2f}"),
                              ("IRR", f"{d['irr']:.2f}%" if d["irr"] else "N/A"),
                              ("ROI", f"{d['estimated_roi']:.2f}%"),
                              ("Payback", f"{d['payback_period']:.2f} yrs" if d["payback_period"] else "N/A")]:
                    w.writerow([k, v])
            self.status_bar.showMessage(f"CSV exported: {path}")
            QtWidgets.QMessageBox.information(self, "Export Complete", f"CSV saved to:\n{path}")
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "CSV Error", str(e))

    # ── Save / Load ───────────────────────────────────────────────────────────
    def save_project(self):
        path, _ = QtWidgets.QFileDialog.getSaveFileName(
            self, "Save Project", "project.json", "JSON Files (*.json)")
        if not path: return
        try:
            state = {
                "project_name": self.project_name.text(), "project_desc": self.project_desc.text(),
                "usd_to_egp": self.usd_input.text(), "eur_to_egp": self.eur_input.text(),
                "curr_value": self.curr_value.text(), "curr_perf": self.curr_perf.text(),
                "budget": self.budget.text(), "impr_perf": self.impr_perf.text(),
                "years": self.years_input.text(), "discount": self.discount_input.text(),
                "fixed_costs": self.fixed_costs_data, "running_costs": self.running_costs_data,
                "savings": self.savings_data,
                "predef_costs":   self._read_predef(self.fixed_costs_table),
                "predef_savings": self._read_predef(self.fixed_savings_table),
            }
            with open(path, "w", encoding="utf-8") as f: json.dump(state, f, indent=2)
            self.status_bar.showMessage(f"Project saved: {path}")
            QtWidgets.QMessageBox.information(self, "Saved", f"Project saved to:\n{path}")
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Save Error", str(e))

    def load_project(self):
        path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self, "Open Project", "", "JSON Files (*.json)")
        if not path: return
        try:
            with open(path, encoding="utf-8") as f: s = json.load(f)
            self.project_name.setText(s.get("project_name",""))
            self.project_desc.setText(s.get("project_desc",""))
            self.usd_input.setText(s.get("usd_to_egp","49.41"))
            self.eur_input.setText(s.get("eur_to_egp","57.26"))
            self.curr_value.setText(s.get("curr_value",""))
            self.curr_perf.setText(s.get("curr_perf",""))
            self.budget.setText(s.get("budget",""))
            self.impr_perf.setText(s.get("impr_perf",""))
            self.years_input.setText(s.get("years","5"))
            self.discount_input.setText(s.get("discount","24.50"))

            # Restore custom tables
            self._restore_custom(self.fixed_table, s.get("fixed_costs",[]), self.fixed_costs_data,
                ["product","qty","unit_cost","currency","total"])
            self._restore_custom(self.running_table, s.get("running_costs",[]), self.running_costs_data,
                ["product","qty","unit_cost","hours_per_day","currency","total"])
            self._restore_custom(self.savings_table, s.get("savings",[]), self.savings_data,
                ["description","amount","currency"])

            for i, row in enumerate(s.get("predef_costs",[])):
                if i < self.fixed_costs_table.rowCount():
                    for col, val in zip([1,2,4], row):
                        w = self.fixed_costs_table.cellWidget(i,col)
                        if w: w.setValue(val)
            for i, row in enumerate(s.get("predef_savings",[])):
                if i < self.fixed_savings_table.rowCount():
                    for col, val in zip([1,2,4], row):
                        w = self.fixed_savings_table.cellWidget(i,col)
                        if w: w.setValue(val)

            self._on_fx_changed()
            self.status_bar.showMessage(f"Project loaded: {path}")
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Load Error", str(e))

    def new_project(self):
        if QtWidgets.QMessageBox.question(
            self, "New Project", "Clear all data and start fresh?",
            QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No) != QtWidgets.QMessageBox.Yes: return
        for w in [self.project_name, self.project_desc, self.curr_value, self.curr_perf,
                  self.budget, self.impr_perf]:
            w.clear()
        self.years_input.setText("5"); self.discount_input.setText("24.50")
        self.fixed_costs_data.clear(); self.running_costs_data.clear(); self.savings_data.clear()
        self.fixed_table.setRowCount(0); self.running_table.setRowCount(0)
        self.savings_table.setRowCount(0); self.cashflow_table.setRowCount(0)
        self.recommendation.setText("Run calculation to see recommendation.")
        self.recommendation.setStyleSheet(
            "font-weight:bold;font-size:13px;padding:12px;background:#f8f9fa;"
            "border-radius:6px;border:1px solid #e0e0e0;")
        for lbl in self.metric_labels.values(): lbl.setText("—")
        self.figure.clear(); self.canvas.draw()
        self.economic_data = None
        self._on_fx_changed()
        self.status_bar.showMessage("New project started")

    def _read_predef(self, table):
        rows = []
        for i in range(table.rowCount()):
            q = table.cellWidget(i,1); c = table.cellWidget(i,2); h = table.cellWidget(i,4)
            if q and c and h: rows.append([q.value(), c.value(), h.value()])
        return rows

    def _restore_custom(self, table, data, store, keys):
        table.setRowCount(0); store.clear(); store.extend(data)
        for item in data:
            r = table.rowCount(); table.insertRow(r)
            for ci, key in enumerate(keys):
                v = item.get(key,"")
                text = f"{v:,.2f}" if isinstance(v, float) else str(v)
                table.setItem(r, ci, QtWidgets.QTableWidgetItem(text))

    def show_about(self):
        QtWidgets.QMessageBox.about(self, "About – Economic Feasibility Analyser",
            "<h2 style='color:#1a73e8'>Economic Feasibility Analyser</h2>"
            "<p>Version 2.0 &nbsp;|&nbsp; Improve Project Edition</p>"
            "<p>Full financial modelling for industrial improvement projects.</p>"
            "<ul>"
            "<li>NPV, IRR, ROI, Payback Period</li>"
            "<li>Multi-currency: USD / EGP / EUR</li>"
            "<li>Predefined + custom operating costs and savings</li>"
            "<li>PDF report export (ReportLab)</li>"
            "<li>CSV cash-flow export</li>"
            "<li>JSON project save &amp; load</li>"
            "<li>Interactive financial charts (matplotlib)</li>"
            "</ul>")


def main():
    app = QtWidgets.QApplication(sys.argv)
    app.setStyle("Fusion")
    window = ImproveWindow(None)
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()