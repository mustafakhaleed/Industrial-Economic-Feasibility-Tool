# Industrial Economic Feasibility Tool (PyQt5)

A professional desktop application designed for industrial engineers and project managers to evaluate the financial viability of "Improvement" projects. The tool computes complex financial metrics and generates comprehensive visual reports.

## 🚀 Key Features
* **Multi-Tab Analysis:** Dedicated sections for Analysis, Upgrades, Running Costs, and Savings.
* **Core Financial Metrics:**
    * **NPV (Net Present Value):** Accounts for time-value of money.
    * **IRR (Internal Rate of Return):** Calculated via iterative numerical methods.
    * **ROI & Payback Period:** Quick efficiency and break-even analysis.
* **Data Visualization:** Integrated **Matplotlib** graphs showing cost vs. savings trends and NPV projections.
* **Professional PDF Export:** Generates a formatted summary report using **ReportLab**.
* **Currency Conversion:** Built-in exchange rates for USD, EUR, and EGP.

## 📐 Mathematical Models
The application automates the following engineering economics formulas:

### Net Present Value (NPV)

$$NPV = \sum_{t=1}^{n} \frac{R_t}{(1+i)^t} - C_{initial}$$

### Internal Rate of Return (IRR)
The tool solves for the discount rate where:
$$0 = NPV = \sum_{t=1}^{n} \frac{CashFlow_t}{(1+IRR)^t}$$

## 🛠️ Tech Stack
* **UI Framework:** PyQt5 (Python)
* **Mathematical Logic:** NumPy
* **Visualization:** Matplotlib
* **Reporting:** ReportLab (PDF Generation)

## 📂 Installation & Usage
1. **Clone the repository:**
   ```bash
   git clone [https://github.com/mustafakhaleed/Industrial-Economic-Feasibility-Tool.git](https://github.com/mustafakhaleed/Industrial-Economic-Feasibility-Tool.git)
   cd Industrial-Economic-Feasibility-Tool
