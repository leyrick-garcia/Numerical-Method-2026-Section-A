"""
Lab 03: Real-World Data Linear Regression
Numerical Methods, Section 3D

Dataset:  Annual mean atmospheric CO2 concentration at Mauna Loa Observatory
Source:   NOAA Global Monitoring Laboratory (Keeling Curve)
          https://gml.noaa.gov/ccgg/trends/  (data file: co2_annmean_mlo.csv)

x = Year
y = Annual mean CO2 concentration (ppm)

The regression is performed "by hand" using the least-squares normal
equations for a straight line, exactly as derived in class:

    a1 = ( n*sum(x*y) - sum(x)*sum(y) ) / ( n*sum(x^2) - (sum(x))^2 )
    a0 = ybar - a1*xbar

No numpy.polyfit / scipy curve-fitting routine is used for the fit itself;
numpy is only used for basic array/sum bookkeeping.
"""

import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

# ---------------------------------------------------------------
# 1. The data (16 paired observations, 2009-2024)
# ---------------------------------------------------------------
year = np.array([2009, 2010, 2011, 2012, 2013, 2014, 2015, 2016,
                  2017, 2018, 2019, 2020, 2021, 2022, 2023, 2024])

co2_ppm = np.array([387.64, 390.10, 391.85, 394.06, 396.74, 398.81,
                     401.01, 404.41, 406.76, 408.72, 411.65, 414.21,
                     416.41, 418.53, 421.08, 424.61])

x = year.astype(float)
y = co2_ppm.astype(float)
n = len(x)

# ---------------------------------------------------------------
# 2. Least-squares straight-line fit:  y = a0 + a1*x
# ---------------------------------------------------------------
sum_x  = np.sum(x)
sum_y  = np.sum(y)
sum_xy = np.sum(x * y)
sum_x2 = np.sum(x ** 2)

a1 = (n * sum_xy - sum_x * sum_y) / (n * sum_x2 - sum_x ** 2)
a0 = (sum_y - a1 * sum_x) / n

print(f"Number of observations, n = {n}")
print(f"Slope,     a1 = {a1:.4f} ppm/year")
print(f"Intercept, a0 = {a0:.4f} ppm")
print(f"Regression equation:  y = {a0:.3f} + {a1:.4f} x")

# ---------------------------------------------------------------
# 3. Goodness of fit: Sr (SSE), St (SST), r^2, standard error sy/x
# ---------------------------------------------------------------
y_pred = a0 + a1 * x
residuals = y - y_pred

St = np.sum((y - np.mean(y)) ** 2)     # total sum of squares
Sr = np.sum(residuals ** 2)            # sum of squares of residuals (SSE)
r2 = 1 - Sr / St                       # coefficient of determination
syx = np.sqrt(Sr / (n - 2))            # standard error of the estimate

print(f"\nSt  (total sum of squares)      = {St:.4f}")
print(f"Sr  (sum of squares of resid.)  = {Sr:.4f}")
print(f"r^2 (coefficient of determin.)  = {r2:.6f}")
print(f"r   (correlation coefficient)   = {np.sqrt(r2):.6f}")
print(f"s_y/x (standard error)          = {syx:.4f} ppm")

# ---------------------------------------------------------------
# 4. Prediction for a year NOT in the dataset
# ---------------------------------------------------------------
x_new = 2030
y_new = a0 + a1 * x_new
print(f"\nPrediction: for x = {x_new}, predicted y = {y_new:.2f} ppm")

# ---------------------------------------------------------------
# 5. Plot 1: data + fitted line
# ---------------------------------------------------------------
plt.figure(figsize=(7, 4.5))
plt.scatter(x, y, color="#1f4e79", label="Observed annual mean CO2", zorder=3)
x_line = np.linspace(x.min() - 1, x.max() + 1, 100)
plt.plot(x_line, a0 + a1 * x_line, color="#c8963e", linewidth=2,
          label=f"Fit: y = {a0:.2f} + {a1:.3f}x")
plt.xlabel("Year")
plt.ylabel("Annual mean CO2 concentration (ppm)")
plt.title("Mauna Loa Annual Mean Atmospheric CO2, 2009-2024")
plt.legend()
plt.grid(alpha=0.3)
plt.tight_layout()
plt.savefig("fit_plot.png", dpi=200)
plt.close()

# ---------------------------------------------------------------
# 6. Plot 2: residual plot
# ---------------------------------------------------------------
plt.figure(figsize=(7, 4.5))
plt.axhline(0, color="black", linewidth=1)
plt.scatter(x, residuals, color="#1f4e79", zorder=3)
plt.xlabel("Year")
plt.ylabel("Residual, y - y_predicted (ppm)")
plt.title("Residual Plot")
plt.grid(alpha=0.3)
plt.tight_layout()
plt.savefig("residual_plot.png", dpi=200)
plt.close()

print("\nSaved fit_plot.png and residual_plot.png")
