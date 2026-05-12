# Industrial Economic Feasibility Tool (PyQt5)

A professional desktop application designed for industrial engineers to evaluate the financial viability of project improvements. The tool compares current vs. improved performance to compute net benefits, visualizes data trends, and generates PDF reports.

## 🚀 Key Features
* **Financial Analysis:** Calculates **NPV (Net Present Value)**, **IRR (Internal Rate of Return)**, **ROI**, and **Payback Period**.
* **Improvement Logic:** Built-in calculation modules to compare existing machinery/processes against proposed upgrades.
* **Data Visualization:** Integrated **Matplotlib** graphs showing cumulative cash flow and NPV projections over time.
* **PDF Export:** Generates a formatted, professional summary report using **ReportLab**.
* **Multi-Currency Support:** Built-in exchange rates for USD, EUR, and EGP.

## 📐 Mathematical Models
The application automates core engineering economics formulas:

### Net Present Value (NPV)
$$NPV = \sum_{t=1}^{n} \frac{R_t}{(1+i)^t} - C_{initial}$$



### Internal Rate of Return (IRR)
The tool utilizes an iterative solver to find the discount rate ($i$) that makes the NPV equal to zero.

## 🛠️ Tech Stack
* **GUI Framework:** PyQt5
* **Calculations:** NumPy
* **Visualization:** Matplotlib
* **Reporting:** ReportLab

## 📂 Installation & Usage

1. **Clone the repository:**
   ```bash
   git clone [https://github.com/mustafakhaleed/Industrial-Economic-Feasibility-Tool.git](https://github.com/mustafakhaleed/Industrial-Economic-Feasibility-Tool.git)
   cd Industrial-Economic-Feasibility-Tool
