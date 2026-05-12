from PyQt5 import QtWidgets, QtGui, QtCore
import sys
import os
import numpy as np
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
import matplotlib.pyplot as plt
from matplotlib.backends.backend_qt5agg import NavigationToolbar2QT as NavigationToolbar

class ImproveWindow(QtWidgets.QMainWindow):
    def __init__(self, master):
        super().__init__()
        self.master = master
        self.setWindowTitle("Economic Feasibility - Improve Project")
        self.setGeometry(150, 150, 1200, 800)
        self.exchange_rates = {"USD_to_EGP": 49.41, "EUR_to_EGP": 57.26}
        self.fixed_costs_data = []
        self.running_costs_data = []
        self.savings_data = []

        self.tabs = QtWidgets.QTabWidget()
        self.setCentralWidget(self.tabs)

        self.init_analysis_tab()
        self.init_upgrade_tab()
        self.init_running_tab()
        self.init_savings_tab()
        self.init_results_tab()
        self.init_graphs_tab()  # Add new tab for graphs

    def init_analysis_tab(self):
        self.analysis_tab = QtWidgets.QWidget()
        layout = QtWidgets.QFormLayout()

        # Exchange rates
        self.usd_input = QtWidgets.QLineEdit("49.41")
        self.eur_input = QtWidgets.QLineEdit("57.26")
        
        # Current system values
        self.curr_value = QtWidgets.QLineEdit()
        self.curr_value.setPlaceholderText("Current system value in USD")
        
        # Improvement parameters
        self.budget = QtWidgets.QLineEdit()
        self.budget.setPlaceholderText("Budget for improvement in USD")
        self.curr_perf = QtWidgets.QLineEdit()
        self.curr_perf.setPlaceholderText("Current annual performance in USD")
        self.impr_perf = QtWidgets.QLineEdit()
        self.impr_perf.setPlaceholderText("Expected performance after improvement in USD")
        
        # Analysis parameters
        self.years_input = QtWidgets.QLineEdit("5")
        self.discount_input = QtWidgets.QLineEdit("24.50")

        layout.addRow("USD to EGP:", self.usd_input)
        layout.addRow("EUR to EGP:", self.eur_input)
        layout.addRow(QtWidgets.QLabel("<b>Current System</b>"))
        layout.addRow("Current System Value ($):", self.curr_value)
        layout.addRow("Current Annual Performance ($):", self.curr_perf)
        layout.addRow(QtWidgets.QLabel("<b>Improvement Parameters</b>"))
        layout.addRow("Improvement Budget ($):", self.budget)
        layout.addRow("Expected Performance After Improvement ($):", self.impr_perf)
        layout.addRow(QtWidgets.QLabel("<b>Analysis Parameters</b>"))
        layout.addRow("Analysis Period (years):", self.years_input)
        layout.addRow("Discount Rate (%):", self.discount_input)

        self.analysis_tab.setLayout(layout)
        self.tabs.addTab(self.analysis_tab, "Project Basics")

    def init_upgrade_tab(self):
        self.upgrade_tab = QtWidgets.QWidget()
        layout = QtWidgets.QVBoxLayout()

        # Form for adding upgrade items
        form_layout = QtWidgets.QHBoxLayout()
        self.fixed_name = QtWidgets.QLineEdit()
        self.fixed_name.setPlaceholderText("Item name")
        self.fixed_qty = QtWidgets.QLineEdit("1")
        self.fixed_qty.setValidator(QtGui.QDoubleValidator())
        self.fixed_unit_cost = QtWidgets.QLineEdit()
        self.fixed_unit_cost.setValidator(QtGui.QDoubleValidator())
        self.fixed_currency = QtWidgets.QComboBox()
        self.fixed_currency.addItems(["USD", "EGP", "EUR"])
        add_btn = QtWidgets.QPushButton("Add Upgrade Cost")
        add_btn.clicked.connect(self.add_upgrade)

        form_layout.addWidget(QtWidgets.QLabel("Item:"))
        form_layout.addWidget(self.fixed_name)
        form_layout.addWidget(QtWidgets.QLabel("Qty:"))
        form_layout.addWidget(self.fixed_qty)
        form_layout.addWidget(QtWidgets.QLabel("Unit Cost:"))
        form_layout.addWidget(self.fixed_unit_cost)
        form_layout.addWidget(QtWidgets.QLabel("Currency:"))
        form_layout.addWidget(self.fixed_currency)
        form_layout.addWidget(add_btn)

        # Upgrade costs table
        self.fixed_table = QtWidgets.QTableWidget(0, 5)
        self.fixed_table.setHorizontalHeaderLabels(["Item", "Qty", "Unit Cost", "Currency", "Total"])
        self.fixed_table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        
        layout.addLayout(form_layout)
        layout.addWidget(self.fixed_table)

        # Summary
        self.fixed_total_label = QtWidgets.QLabel("Total Upgrade Cost (USD): $0.00")
        self.fixed_total_label.setStyleSheet("font-weight: bold;")
        layout.addWidget(self.fixed_total_label)

        # Delete button
        del_btn = QtWidgets.QPushButton("Delete Selected")
        del_btn.clicked.connect(self.del_upgrade)
        layout.addWidget(del_btn)

        self.upgrade_tab.setLayout(layout)
        self.tabs.addTab(self.upgrade_tab, "Upgrade Costs")

    def init_running_tab(self):
        self.running_tab = QtWidgets.QWidget()
        layout = QtWidgets.QVBoxLayout()

        # Predefined running costs
        predefined_group = QtWidgets.QGroupBox("Standard Operating Costs (EGP)")
        predefined_layout = QtWidgets.QVBoxLayout()
        
        self.predefined_costs = [
            {"name": "Petrol 92 Consumption", "qty": 1, "unit_cost": 13.75, "unit": "liter/hour", "hours": 16},
            {"name": "Petrol gas Consumption", "qty": 1, "unit_cost": 11.5, "unit": "liter/hour", "hours": 16},
            {"name": "Power Consumption", "qty": 2, "unit_cost": 1.25, "unit": "KW/Hour", "hours": 16},
            {"name": "Air Consumption", "qty": 1, "unit_cost": 0.21, "unit": "m3/Hour", "hours": 16},
            {"name": "Gas Consumption", "qty": 1, "unit_cost": 2.7, "unit": "m3/Hour", "hours": 16},
            {"name": "Needed Labor", "qty": 1, "unit_cost": 8500, "unit": "labour/month", "hours": 8},
            {"name": "Needed area", "qty": 24, "unit_cost": 1, "unit": "m2/month", "hours": 8}
        ]

        self.fixed_costs_table = QtWidgets.QTableWidget(len(self.predefined_costs), 6)
        self.fixed_costs_table.setHorizontalHeaderLabels(["Product Name", "Qty", "Cost (per unit)", "Unit", "Hours/Day", "Daily Total (EGP)"])
        self.fixed_costs_table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)

        for i, item in enumerate(self.predefined_costs):
            self.fixed_costs_table.setItem(i, 0, QtWidgets.QTableWidgetItem(item["name"]))
            
            qty_spinbox = QtWidgets.QDoubleSpinBox()
            qty_spinbox.setRange(0, 99999)
            qty_spinbox.setValue(item["qty"])
            qty_spinbox.valueChanged.connect(self.update_running_totals)
            self.fixed_costs_table.setCellWidget(i, 1, qty_spinbox)
            
            cost_spinbox = QtWidgets.QDoubleSpinBox()
            cost_spinbox.setRange(0, 99999)
            cost_spinbox.setValue(item["unit_cost"])
            cost_spinbox.valueChanged.connect(self.update_running_totals)
            self.fixed_costs_table.setCellWidget(i, 2, cost_spinbox)
            
            self.fixed_costs_table.setItem(i, 3, QtWidgets.QTableWidgetItem(item["unit"]))
            
            hours_spinbox = QtWidgets.QDoubleSpinBox()
            hours_spinbox.setRange(0, 24)
            hours_spinbox.setValue(item["hours"])
            hours_spinbox.valueChanged.connect(self.update_running_totals)
            self.fixed_costs_table.setCellWidget(i, 4, hours_spinbox)
            
            daily_total = item["qty"] * item["unit_cost"] * item["hours"]
            total_label = QtWidgets.QLabel(f"{daily_total:.2f}")
            total_label.setAlignment(QtCore.Qt.AlignCenter)
            self.fixed_costs_table.setCellWidget(i, 5, total_label)

        predefined_layout.addWidget(self.fixed_costs_table)
        predefined_group.setLayout(predefined_layout)
        layout.addWidget(predefined_group)

        # Additional running costs
        custom_group = QtWidgets.QGroupBox("Additional Operating Costs")
        custom_layout = QtWidgets.QVBoxLayout()
        
        form_layout = QtWidgets.QHBoxLayout()
        self.run_name = QtWidgets.QLineEdit()
        self.run_qty = QtWidgets.QLineEdit("1")
        self.run_qty.setValidator(QtGui.QDoubleValidator())
        self.run_unit_cost = QtWidgets.QLineEdit()
        self.run_unit_cost.setValidator(QtGui.QDoubleValidator())
        self.run_hours = QtWidgets.QLineEdit("8")
        self.run_hours.setValidator(QtGui.QDoubleValidator())
        self.run_currency = QtWidgets.QComboBox()
        self.run_currency.addItems(["USD", "EGP", "EUR"])
        add_btn = QtWidgets.QPushButton("Add Operating Cost")
        add_btn.clicked.connect(self.add_running)

        form_layout.addWidget(QtWidgets.QLabel("Item:"))
        form_layout.addWidget(self.run_name)
        form_layout.addWidget(QtWidgets.QLabel("Qty:"))
        form_layout.addWidget(self.run_qty)
        form_layout.addWidget(QtWidgets.QLabel("Unit Cost:"))
        form_layout.addWidget(self.run_unit_cost)
        form_layout.addWidget(QtWidgets.QLabel("Hours/Day:"))
        form_layout.addWidget(self.run_hours)
        form_layout.addWidget(QtWidgets.QLabel("Currency:"))
        form_layout.addWidget(self.run_currency)
        form_layout.addWidget(add_btn)

        self.running_table = QtWidgets.QTableWidget(0, 6)
        self.running_table.setHorizontalHeaderLabels(["Item", "Qty", "Unit Cost", "Hours/Day", "Currency", "Daily Total"])
        self.running_table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        
        custom_layout.addLayout(form_layout)
        custom_layout.addWidget(self.running_table)
        
        self.running_total_label = QtWidgets.QLabel("Total Daily Operating Cost (USD): $0.00")
        self.running_total_label.setStyleSheet("font-weight: bold;")
        custom_layout.addWidget(self.running_total_label)

        del_btn = QtWidgets.QPushButton("Delete Selected")
        del_btn.clicked.connect(self.del_running)
        custom_layout.addWidget(del_btn)

        custom_group.setLayout(custom_layout)
        layout.addWidget(custom_group)

        self.running_tab.setLayout(layout)
        self.tabs.addTab(self.running_tab, "Operating Costs")

    def init_savings_tab(self):
        self.savings_tab = QtWidgets.QWidget()
        layout = QtWidgets.QVBoxLayout()

        # Predefined savings
        predefined_group = QtWidgets.QGroupBox("Standard Savings (EGP)")
        predefined_layout = QtWidgets.QVBoxLayout()
        
        self.predefined_savings = [
            {"name": "Petrol 92 Consumption", "qty": 1, "unit_cost": 13.75, "unit": "liter/hour", "hours": 16},
            {"name": "Petrol gas Consumption", "qty": 1, "unit_cost": 11.5, "unit": "liter/hour", "hours": 16},
            {"name": "Power Consumption", "qty": 2, "unit_cost": 1.25, "unit": "KW/Hour", "hours": 16},
            {"name": "Air Consumption", "qty": 1, "unit_cost": 0.21, "unit": "m3/Hour", "hours": 16},
            {"name": "Gas Consumption", "qty": 1, "unit_cost": 2.7, "unit": "m3/Hour", "hours": 16},
            {"name": "Saving Labor", "qty": 1, "unit_cost": 8500, "unit": "labour/month", "hours": 8},
            {"name": "Saving area", "qty": 24, "unit_cost": 1, "unit": "m2/month", "hours": 8}
        ]

        self.fixed_savings_table = QtWidgets.QTableWidget(len(self.predefined_savings), 6)
        self.fixed_savings_table.setHorizontalHeaderLabels(["Product Name", "Qty", "Cost (per unit)", "Unit", "Hours/Day", "Daily Savings (EGP)"])
        self.fixed_savings_table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)

        for i, item in enumerate(self.predefined_savings):
            self.fixed_savings_table.setItem(i, 0, QtWidgets.QTableWidgetItem(item["name"]))
            
            qty_spinbox = QtWidgets.QDoubleSpinBox()
            qty_spinbox.setRange(0, 99999)
            qty_spinbox.setValue(item["qty"])
            qty_spinbox.valueChanged.connect(self.update_savings_totals)
            self.fixed_savings_table.setCellWidget(i, 1, qty_spinbox)
            
            cost_spinbox = QtWidgets.QDoubleSpinBox()
            cost_spinbox.setRange(0, 99999)
            cost_spinbox.setValue(item["unit_cost"])
            cost_spinbox.valueChanged.connect(self.update_savings_totals)
            self.fixed_savings_table.setCellWidget(i, 2, cost_spinbox)
            
            self.fixed_savings_table.setItem(i, 3, QtWidgets.QTableWidgetItem(item["unit"]))
            
            hours_spinbox = QtWidgets.QDoubleSpinBox()
            hours_spinbox.setRange(0, 24)
            hours_spinbox.setValue(item["hours"])
            hours_spinbox.valueChanged.connect(self.update_savings_totals)
            self.fixed_savings_table.setCellWidget(i, 4, hours_spinbox)
            
            daily_total = item["qty"] * item["unit_cost"] * item["hours"]
            total_label = QtWidgets.QLabel(f"{daily_total:.2f}")
            total_label.setAlignment(QtCore.Qt.AlignCenter)
            self.fixed_savings_table.setCellWidget(i, 5, total_label)

        predefined_layout.addWidget(self.fixed_savings_table)
        predefined_group.setLayout(predefined_layout)
        layout.addWidget(predefined_group)

        # Additional savings
        custom_group = QtWidgets.QGroupBox("Additional Savings")
        custom_layout = QtWidgets.QVBoxLayout()
        
        form_layout = QtWidgets.QHBoxLayout()
        self.save_desc = QtWidgets.QLineEdit()
        self.save_amount = QtWidgets.QLineEdit()
        self.save_amount.setValidator(QtGui.QDoubleValidator())
        self.save_currency = QtWidgets.QComboBox()
        self.save_currency.addItems(["USD", "EGP", "EUR"])
        add_btn = QtWidgets.QPushButton("Add Annual Saving")
        add_btn.clicked.connect(self.add_saving)

        form_layout.addWidget(QtWidgets.QLabel("Description:"))
        form_layout.addWidget(self.save_desc)
        form_layout.addWidget(QtWidgets.QLabel("Annual Savings:"))
        form_layout.addWidget(self.save_amount)
        form_layout.addWidget(QtWidgets.QLabel("Currency:"))
        form_layout.addWidget(self.save_currency)
        form_layout.addWidget(add_btn)

        self.savings_table = QtWidgets.QTableWidget(0, 3)
        self.savings_table.setHorizontalHeaderLabels(["Description", "Amount", "Currency"])
        self.savings_table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        
        custom_layout.addLayout(form_layout)
        custom_layout.addWidget(self.savings_table)
        
        self.savings_total_label = QtWidgets.QLabel("Total Annual Savings (USD): $0.00")
        self.savings_total_label.setStyleSheet("font-weight: bold;")
        custom_layout.addWidget(self.savings_total_label)

        del_btn = QtWidgets.QPushButton("Delete Selected")
        del_btn.clicked.connect(self.del_saving)
        custom_layout.addWidget(del_btn)

        custom_group.setLayout(custom_layout)
        layout.addWidget(custom_group)

        self.savings_tab.setLayout(layout)
        self.tabs.addTab(self.savings_tab, "Savings")

    def init_results_tab(self):
        self.results_tab = QtWidgets.QWidget()
        layout = QtWidgets.QVBoxLayout()

        # Calculate button
        self.calc_btn = QtWidgets.QPushButton("Calculate Improvement Metrics")
        self.calc_btn.setStyleSheet("font-weight: bold; font-size: 14px;")
        self.calc_btn.clicked.connect(self.calculate_metrics)
        layout.addWidget(self.calc_btn)

        # Key metrics
        metrics_group = QtWidgets.QGroupBox("Financial Metrics")
        metrics_layout = QtWidgets.QFormLayout()
        
        self.metric_labels = {
            "total_investment": QtWidgets.QLabel("$0.00"),
            "annual_running_cost": QtWidgets.QLabel("$0.00"),
            "annual_savings": QtWidgets.QLabel("$0.00"),
            "net_annual_benefit": QtWidgets.QLabel("$0.00"),
            "roi": QtWidgets.QLabel("0.00%"),
            "npv": QtWidgets.QLabel("$0.00"),
            "irr": QtWidgets.QLabel("0.00%"),
            "payback_period": QtWidgets.QLabel("0.00 years")
        }

        for key, label in self.metric_labels.items():
            metrics_layout.addRow(f"{key.replace('_', ' ').title()}:", label)

        metrics_group.setLayout(metrics_layout)
        layout.addWidget(metrics_group)

        # Cash flow table
        self.cashflow_table = QtWidgets.QTableWidget(0, 5)
        self.cashflow_table.setHorizontalHeaderLabels([
            "Year", 
            "Investment", 
            "Running Costs", 
            "Savings & Benefits",
            "Net Cash Flow"
        ])
        layout.addWidget(self.cashflow_table)

        # Recommendation
        self.recommendation = QtWidgets.QLabel()
        self.recommendation.setWordWrap(True)
        self.recommendation.setStyleSheet("font-weight: bold; font-size: 14px; padding: 10px;")
        layout.addWidget(self.recommendation)

        self.results_tab.setLayout(layout)
        self.tabs.addTab(self.results_tab, "Results & Recommendation")

    def init_graphs_tab(self):
        """Initialize a separate tab just for graphs"""
        self.graphs_tab = QtWidgets.QWidget()
        layout = QtWidgets.QVBoxLayout()
        
        # Create matplotlib Figure and Canvas objects
        self.figure = Figure(figsize=(10, 8), dpi=100)
        self.canvas = FigureCanvas(self.figure)
        
        # Add navigation toolbar
        self.toolbar = NavigationToolbar(self.canvas, self)
        
        # Add widgets to layout
        layout.addWidget(self.toolbar)
        layout.addWidget(self.canvas)
        
        # Set the layout for the graphs tab
        self.graphs_tab.setLayout(layout)
        self.tabs.addTab(self.graphs_tab, "Financial Graphs")

    def update_graphs(self, cash_flows, years, npv, irr, payback):
        """Update the graphs with new data"""
        # Clear the figure
        self.figure.clear()
        
        # Create subplots with adjusted layout
        gs = self.figure.add_gridspec(2, 2, height_ratios=[1, 1], width_ratios=[1, 1])
        
        # Plot 1: Annual Cash Flows
        ax1 = self.figure.add_subplot(gs[0, 0])
        years_list = list(range(years + 1))
        ax1.bar(years_list, cash_flows, color=['red' if cf < 0 else 'green' for cf in cash_flows])
        ax1.set_title('Annual Cash Flows')
        ax1.set_xlabel('Year')
        ax1.set_ylabel('Amount ($)')
        ax1.axhline(0, color='black', linewidth=0.5)
        ax1.grid(True)
        
        # Plot 2: NPV Sensitivity to Discount Rate
        ax2 = self.figure.add_subplot(gs[0, 1])
        if isinstance(irr, (int, float)):
            rates = np.linspace(0, irr/100 + 0.2, 20)
            npvs = []
            for r in rates:
                npvs.append(sum(cf / ((1 + r) ** i) for i, cf in enumerate(cash_flows)))
            
            ax2.plot(rates*100, npvs, 'b-')
            ax2.axvline(irr, color='r', linestyle='--', label=f'IRR: {irr:.1f}%')
            ax2.axhline(0, color='black', linewidth=0.5)
            ax2.set_title('NPV Sensitivity to Discount Rate')
            ax2.set_xlabel('Discount Rate (%)')
            ax2.set_ylabel('NPV ($)')
            ax2.legend()
            ax2.grid(True)
        
        # Plot 3: Cumulative Cash Flow
        ax3 = self.figure.add_subplot(gs[1, :])
        cumulative = np.cumsum(cash_flows)
        ax3.plot(years_list, cumulative, 'b-o')
        if isinstance(payback, (int, float)):
            ax3.axvline(payback, color='r', linestyle='--', label=f'Payback: {payback:.1f} years')
        ax3.axhline(0, color='black', linewidth=0.5)
        ax3.set_title('Cumulative Cash Flow')
        ax3.set_xlabel('Year')
        ax3.set_ylabel('Cumulative Amount ($)')
        ax3.legend()
        ax3.grid(True)
        
        # Adjust layout and refresh
        self.figure.tight_layout()
        self.canvas.draw()
        
    def add_upgrade(self):
        try:
            item = self.fixed_name.text()
            qty = float(self.fixed_qty.text())
            unit_cost = float(self.fixed_unit_cost.text())
            currency = self.fixed_currency.currentText()
            total = qty * unit_cost
            
            self.fixed_costs_data.append({
                "product": item,
                "qty": qty,
                "unit_cost": unit_cost,
                "total": total,
                "currency": currency
            })
            
            row = self.fixed_table.rowCount()
            self.fixed_table.insertRow(row)
            self.fixed_table.setItem(row, 0, QtWidgets.QTableWidgetItem(item))
            self.fixed_table.setItem(row, 1, QtWidgets.QTableWidgetItem(str(qty)))
            self.fixed_table.setItem(row, 2, QtWidgets.QTableWidgetItem(f"{unit_cost:.2f}"))
            self.fixed_table.setItem(row, 3, QtWidgets.QTableWidgetItem(currency))
            self.fixed_table.setItem(row, 4, QtWidgets.QTableWidgetItem(f"{total:.2f}"))
            
            self.update_fixed_totals()
            self.fixed_name.clear()
            self.fixed_qty.clear()
            self.fixed_unit_cost.clear()
        except ValueError:
            QtWidgets.QMessageBox.warning(self, "Error", "Please enter valid numbers for quantity and unit cost")

    def del_upgrade(self):
        selected = self.fixed_table.selectedItems()
        if not selected:
            return
            
        rows = {item.row() for item in selected}
        for row in sorted(rows, reverse=True):
            self.fixed_table.removeRow(row)
            if row < len(self.fixed_costs_data):
                del self.fixed_costs_data[row]
        self.update_fixed_totals()

    def update_fixed_totals(self):
        total_usd = 0
        usd_to_egp = self.exchange_rates["USD_to_EGP"]
        eur_to_egp = self.exchange_rates["EUR_to_EGP"]
        
        for item in self.fixed_costs_data:
            total = item["total"]
            if item["currency"] == "EGP":
                total /= usd_to_egp
            elif item["currency"] == "EUR":
                total *= (eur_to_egp / usd_to_egp)
            total_usd += total
        
        self.fixed_total_label.setText(f"Total Upgrade Cost (USD): ${total_usd:,.2f}")

    def add_running(self):
        try:
            item = self.run_name.text()
            qty = float(self.run_qty.text())
            unit_cost = float(self.run_unit_cost.text())
            hours = float(self.run_hours.text())
            currency = self.run_currency.currentText()
            total = qty * unit_cost * hours
            
            self.running_costs_data.append({
                "product": item,
                "qty": qty,
                "unit_cost": unit_cost,
                "hours_per_day": hours,
                "total": total,
                "currency": currency
            })
            
            row = self.running_table.rowCount()
            self.running_table.insertRow(row)
            self.running_table.setItem(row, 0, QtWidgets.QTableWidgetItem(item))
            self.running_table.setItem(row, 1, QtWidgets.QTableWidgetItem(str(qty)))
            self.running_table.setItem(row, 2, QtWidgets.QTableWidgetItem(f"{unit_cost:.2f}"))
            self.running_table.setItem(row, 3, QtWidgets.QTableWidgetItem(str(hours)))
            self.running_table.setItem(row, 4, QtWidgets.QTableWidgetItem(currency))
            self.running_table.setItem(row, 5, QtWidgets.QTableWidgetItem(f"{total:.2f}"))
            
            self.update_running_totals()
            self.run_name.clear()
            self.run_qty.clear()
            self.run_unit_cost.clear()
            self.run_hours.clear()
        except ValueError:
            QtWidgets.QMessageBox.warning(self, "Error", "Please enter valid numbers for running cost")

    def del_running(self):
        selected = self.running_table.selectedItems()
        if not selected:
            return
            
        rows = {item.row() for item in selected}
        for row in sorted(rows, reverse=True):
            self.running_table.removeRow(row)
            if row < len(self.running_costs_data):
                del self.running_costs_data[row]
        self.update_running_totals()

    def update_running_totals(self):
        total_costs_usd = 0
        usd_to_egp = self.exchange_rates["USD_to_EGP"]
        eur_to_egp = self.exchange_rates["EUR_to_EGP"]

        # Update predefined costs table and calculate totals
        for i in range(self.fixed_costs_table.rowCount()):
            qty_widget = self.fixed_costs_table.cellWidget(i, 1)
            cost_widget = self.fixed_costs_table.cellWidget(i, 2)
            hours_widget = self.fixed_costs_table.cellWidget(i, 4)
            total_widget = self.fixed_costs_table.cellWidget(i, 5)

            if qty_widget and cost_widget and hours_widget:
                qty = qty_widget.value()
                cost = cost_widget.value()
                hours = hours_widget.value()
                daily_total_egp = qty * cost * hours
                total_widget.setText(f"{daily_total_egp:.2f}")
                total_costs_usd += daily_total_egp / usd_to_egp

        # Add custom running costs
        for item in self.running_costs_data:
            total = item["total"]
            if item["currency"] == "EGP":
                total = total / usd_to_egp
            elif item["currency"] == "EUR":
                total = total * (eur_to_egp / usd_to_egp)
            total_costs_usd += total

        self.running_total_label.setText(f"Total Daily Operating Cost (USD): ${total_costs_usd:,.2f}")

    def add_saving(self):
        try:
            desc = self.save_desc.text()
            amount = float(self.save_amount.text())
            currency = self.save_currency.currentText()
            
            self.savings_data.append({
                "description": desc,
                "amount": amount,
                "currency": currency
            })
            
            row = self.savings_table.rowCount()
            self.savings_table.insertRow(row)
            self.savings_table.setItem(row, 0, QtWidgets.QTableWidgetItem(desc))
            self.savings_table.setItem(row, 1, QtWidgets.QTableWidgetItem(f"{amount:,.2f}"))
            self.savings_table.setItem(row, 2, QtWidgets.QTableWidgetItem(currency))
            
            self.update_savings_totals()
            self.save_desc.clear()
            self.save_amount.clear()
        except ValueError:
            QtWidgets.QMessageBox.warning(self, "Error", "Please enter a valid savings amount")

    def del_saving(self):
        selected = self.savings_table.selectedItems()
        if not selected:
            return
            
        rows = {item.row() for item in selected}
        for row in sorted(rows, reverse=True):
            self.savings_table.removeRow(row)
            if row < len(self.savings_data):
                del self.savings_data[row]
        self.update_savings_totals()

    def update_savings_totals(self):
        total_savings_usd = 0
        usd_to_egp = self.exchange_rates["USD_to_EGP"]
        eur_to_egp = self.exchange_rates["EUR_to_EGP"]

        # Update predefined savings table and calculate totals
        for i in range(self.fixed_savings_table.rowCount()):
            qty_widget = self.fixed_savings_table.cellWidget(i, 1)
            cost_widget = self.fixed_savings_table.cellWidget(i, 2)
            hours_widget = self.fixed_savings_table.cellWidget(i, 4)
            total_widget = self.fixed_savings_table.cellWidget(i, 5)

            if qty_widget and cost_widget and hours_widget:
                qty = qty_widget.value()
                cost = cost_widget.value()
                hours = hours_widget.value()
                daily_total_egp = qty * cost * hours
                total_widget.setText(f"{daily_total_egp:.2f}")
                total_savings_usd += daily_total_egp / usd_to_egp * 365  # Convert daily to annual

        # Add custom savings
        for item in self.savings_data:
            amount = item["amount"]
            if item["currency"] == "EGP":
                amount = amount / usd_to_egp
            elif item["currency"] == "EUR":
                amount = amount * (eur_to_egp / usd_to_egp)
            total_savings_usd += amount

        self.savings_total_label.setText(f"Total Annual Savings (USD): ${total_savings_usd:,.2f}")

    def get_fixed_costs_data(self):
        costs = []
        for i in range(self.fixed_costs_table.rowCount()):
            name = self.fixed_costs_table.item(i, 0).text()
            qty = self.fixed_costs_table.cellWidget(i, 1).value()
            unit_cost = self.fixed_costs_table.cellWidget(i, 2).value()
            unit = self.fixed_costs_table.item(i, 3).text()
            hours = self.fixed_costs_table.cellWidget(i, 4).value()
            daily_total = qty * unit_cost * hours

            item_data = {
                "product": name,
                "qty": qty,
                "unit_cost": unit_cost,
                "unit": unit,
                "hours_per_day": hours,
                "total": daily_total,
                "currency": "EGP"
            }
            costs.append(item_data)
        return costs

    def get_fixed_savings_data(self):
        savings = []
        for i in range(self.fixed_savings_table.rowCount()):
            name = self.fixed_savings_table.item(i, 0).text()
            qty = self.fixed_savings_table.cellWidget(i, 1).value()
            unit_cost = self.fixed_savings_table.cellWidget(i, 2).value()
            unit = self.fixed_savings_table.item(i, 3).text()
            hours = self.fixed_savings_table.cellWidget(i, 4).value()
            daily_total = qty * unit_cost * hours

            item_data = {
                "product": name,
                "qty": qty,
                "unit_cost": unit_cost,
                "unit": unit,
                "hours_per_day": hours,
                "total": daily_total,
                "currency": "EGP"
            }
            savings.append(item_data)
        return savings

    def calculate_metrics(self):
        try:
            # Get basic parameters
            self.exchange_rates["USD_to_EGP"] = float(self.usd_input.text())
            self.exchange_rates["EUR_to_EGP"] = float(self.eur_input.text())
            current_value = float(self.curr_value.text()) if self.curr_value.text() else 0
            inv = float(self.budget.text()) if self.budget.text() else 0
            current_perf = float(self.curr_perf.text()) if self.curr_perf.text() else 0
            improved_perf = float(self.impr_perf.text()) if self.impr_perf.text() else 0
            yrs = int(self.years_input.text())
            discount_rate = float(self.discount_input.text()) / 100
            usd_to_egp = self.exchange_rates["USD_to_EGP"]
            eur_to_egp = self.exchange_rates["EUR_to_EGP"]

            # Calculate total upgrade costs in USD
            fixed_total_usd = sum(
                item["total"] / usd_to_egp if item["currency"] == "EGP" else
                item["total"] * (eur_to_egp / usd_to_egp) if item["currency"] == "EUR" else
                item["total"]
                for item in self.fixed_costs_data
            )

            # Calculate total running costs in USD
            running_total_usd = sum(
                item["total"] / usd_to_egp if item["currency"] == "EGP" else
                item["total"] * (eur_to_egp / usd_to_egp) if item["currency"] == "EUR" else
                item["total"]
                for item in self.running_costs_data
            )

            # Calculate total savings in USD
            savings_total_usd = sum(
                item["amount"] / usd_to_egp if item["currency"] == "EGP" else
                item["amount"] * (eur_to_egp / usd_to_egp) if item["currency"] == "EUR" else
                item["amount"]
                for item in self.savings_data
            )

            # Get fixed variable costs and savings
            fixed_var_costs = self.get_fixed_costs_data()
            fixed_var_savings = self.get_fixed_savings_data()
            fixed_var_cost_total_usd = sum(item["total"] for item in fixed_var_costs) / usd_to_egp * 365  # Daily to annual
            fixed_var_savings_total_usd = sum(item["total"] for item in fixed_var_savings) / usd_to_egp * 365

            # Calculate key metrics
            total_investment = inv + fixed_total_usd
            annual_running_cost = (running_total_usd + fixed_var_cost_total_usd) * 365
            annual_savings = savings_total_usd + fixed_var_savings_total_usd
            net_annual_benefit = (improved_perf - current_perf) - annual_running_cost + annual_savings

            # ROI calculation
            roi = (net_annual_benefit * yrs - total_investment) / total_investment * 100 if total_investment > 0 else 0

            # Cash flows for NPV and IRR
            cash_flows = [-total_investment] + [net_annual_benefit] * yrs

            # NPV calculation
            npv = sum(cf / ((1 + discount_rate) ** i) for i, cf in enumerate(cash_flows))

            # IRR calculation
            irr = "N/A"
            try:
                rates = np.linspace(0.01, 1.0, 1000)
                npv_values = []
                for rate in rates:
                    npv_val = sum(cf / ((1 + rate) ** i) for i, cf in enumerate(cash_flows))
                    npv_values.append(abs(npv_val))
                min_idx = np.argmin(npv_values)
                irr = rates[min_idx] * 100
                if irr > 100:
                    irr = "N/A"
            except:
                pass

            # Payback period calculation
            cumulative_cash_flow = 0
            payback = "N/A"
            for year in range(1, yrs + 1):
                cumulative_cash_flow += net_annual_benefit
                if cumulative_cash_flow >= total_investment:
                    payback = year - 1 + (total_investment - (cumulative_cash_flow - net_annual_benefit)) / net_annual_benefit
                    break

            # Update results display
            self.metric_labels["total_investment"].setText(f"${total_investment:,.2f}")
            self.metric_labels["annual_running_cost"].setText(f"${annual_running_cost:,.2f}")
            self.metric_labels["annual_savings"].setText(f"${annual_savings:,.2f}")
            self.metric_labels["net_annual_benefit"].setText(f"${net_annual_benefit:,.2f}")
            self.metric_labels["roi"].setText(f"{roi:.2f}%")
            self.metric_labels["npv"].setText(f"${npv:,.2f}")
            self.metric_labels["irr"].setText(f"{irr:.2f}%" if isinstance(irr, (int, float)) else irr)
            self.metric_labels["payback_period"].setText(f"{payback:.2f} years" if isinstance(payback, (int, float)) else payback)

            # Update cash flow table
            self.cashflow_table.setRowCount(0)
            cumulative = 0
            for year in range(yrs + 1):
                if year == 0:
                    investment = total_investment
                    running = 0
                    savings = 0
                    net = -total_investment
                else:
                    investment = 0
                    running = annual_running_cost
                    savings = annual_savings + (improved_perf - current_perf)
                    net = net_annual_benefit
                
                cumulative += net
                
                row = self.cashflow_table.rowCount()
                self.cashflow_table.insertRow(row)
                self.cashflow_table.setItem(row, 0, QtWidgets.QTableWidgetItem(str(year)))
                self.cashflow_table.setItem(row, 1, QtWidgets.QTableWidgetItem(f"{investment:,.2f}"))
                self.cashflow_table.setItem(row, 2, QtWidgets.QTableWidgetItem(f"{running:,.2f}"))
                self.cashflow_table.setItem(row, 3, QtWidgets.QTableWidgetItem(f"{savings:,.2f}"))
                self.cashflow_table.setItem(row, 4, QtWidgets.QTableWidgetItem(f"{net:,.2f}"))

            # Generate recommendation
            recommendation = ""
            if npv > 0:
                recommendation = "✅ RECOMMEND IMPROVEMENT - Project has positive NPV and is financially viable"
                if isinstance(irr, (int, float)) and irr > discount_rate * 100:
                    recommendation += f" with IRR of {irr:.2f}% (above discount rate)"
            else:
                recommendation = "⚠️ DO NOT IMPROVE - Project does not meet financial criteria"
                if isinstance(payback, (int, float)) and payback > yrs:
                    recommendation += f" (payback period of {payback:.1f} years exceeds analysis period)"

            self.recommendation.setText(recommendation)

            # Update graphs
            self.update_graphs(cash_flows, yrs, npv, irr if isinstance(irr, (int, float)) else 0, 
                             payback if isinstance(payback, (int, float)) else 0)

            # Store results for PDF export
            self.master.economic_data = {
                "project_type": "Improve",
                "exchange_rates": self.exchange_rates,
                "current_value": current_value,
                "investment": inv,
                "current_performance": current_perf,
                "improved_performance": improved_perf,
                "years": yrs,
                "discount_rate": discount_rate * 100,
                "fixed_costs": self.fixed_costs_data,
                "running_costs": self.running_costs_data,
                "savings": self.savings_data,
                "fixed_variables_costs": fixed_var_costs,
                "fixed_variables_savings": fixed_var_savings,
                "total_investment": total_investment,
                "annual_running_cost": annual_running_cost,
                "annual_savings": annual_savings,
                "net_annual_benefit": net_annual_benefit,
                "estimated_roi": roi,
                "npv": npv,
                "irr": irr if isinstance(irr, float) else 0,
                "payback_period": payback if isinstance(payback, float) else 0,
                "recommendation": recommendation
            }

        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Error", f"Calculation error: {str(e)}")

def main():
    app = QtWidgets.QApplication(sys.argv)
    window = ImproveWindow(None)  # Passing None as master since we're using it standalone
    window.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()