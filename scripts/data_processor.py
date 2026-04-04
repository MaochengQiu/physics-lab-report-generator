'''
Core data processing script for the physics-lab-report-generator skill.

Handles:
- Mean and standard deviation calculation with expanded formulas.
- Grubbs' test for outlier detection (bad value test).
- Uncertainty calculations (Type A, Type B).
- Physics-specific rounding rules.
- Linear fitting and plotting.
'''

import math
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from scipy import stats

def calculate_mean_expanded(data, unit=""):
    """Calculates the arithmetic mean and returns the expanded formula string and the mean value."""
    n = len(data)
    if n == 0:
        return 0, ""
    sum_data = sum(data)
    mean = sum_data / n
    # Correctly construct the string for the formula
    sum_str = "+".join(map(str, data))
    expanded_str = f"\\bar{{x}} = \\frac{{1}}{{{n}}}({sum_str}) = \\frac{{{sum_data:.3f}}}{{{n}}} = {mean:.3f} {unit}"
    return mean, expanded_str

def calculate_std_dev_expanded(data, mean, unit=""):
    """Calculates the sample standard deviation and returns the expanded formula string and the value."""
    n = len(data)
    if n < 2:
        return 0, "Not enough data points to calculate standard deviation."
    
    residuals_sq_sum = sum([(x - mean)**2 for x in data])
    std_dev = np.sqrt(residuals_sq_sum / (n - 1))
    
    expanded_str = f"\\sigma = \\sqrt{{\\frac{{{residuals_sq_sum:.4f}}}{{{n-1}}}}} = {std_dev:.4f} {unit}"
    return std_dev, expanded_str

def bad_value_test(data, unit=""):
    """
    Performs Grubbs' test for outliers (bad values) iteratively.
    Returns the processed data, final mean, final standard deviation, and the test log string.
    """
    processed_data = list(data)
    test_log = []
    iteration = 0
    
    while True:
        iteration += 1
        n = len(processed_data)
        if n < 3: # Grubbs' test requires at least 3 data points
            test_log.append(f"Round {iteration}: Not enough data points ({n}) for Grubbs' test.")
            break

        current_mean = np.mean(processed_data)
        current_std_dev = np.std(processed_data, ddof=1)
        
        test_log.append(f"\nRound {iteration} Test:")
        test_log.append(f"  Current Mean \\bar{{x}} = {current_mean:.4f} {unit}, Std Dev \\sigma = {current_std_dev:.4f} {unit}")

        # Calculate G values for min and max
        g_max = (np.max(processed_data) - current_mean) / current_std_dev
        g_min = (current_mean - np.min(processed_data)) / current_std_dev
        g_calculated = max(g_max, g_min)
        
        # Get critical G value (two-sided test)
        alpha = 0.05 # Significance level
        t_critical = stats.t.ppf(alpha / (2 * n), n - 2)
        g_critical = ((n - 1) / np.sqrt(n)) * np.sqrt(t_critical**2 / (n - 2 + t_critical**2))

        test_log.append(f"  G calculated = {g_calculated:.4f}, G critical (alpha=0.05) = {g_critical:.4f}")

        if g_calculated > g_critical:
            outlier_value = np.max(processed_data) if g_max > g_min else np.min(processed_data)
            test_log.append(f"  Outlier detected: {outlier_value:.2f} is a bad value (G_calc > G_crit). It has been removed.")
            processed_data.remove(outlier_value)
            test_log.append(f"  Data after removal: {processed_data}")
        else:
            test_log.append("No outliers detected. Test finished.")
            break

    final_mean = np.mean(processed_data) if processed_data else 0
    final_std_dev = np.std(processed_data, ddof=1) if len(processed_data) > 1 else 0
    
    if not processed_data:
        test_log.append("All data points were removed.")
        
    return processed_data, final_mean, final_std_dev, "\n".join(test_log)

def calculate_uncertainties_expanded(data, std_dev, delta_instrument, unit=""):
    """Calculates uncertainties (Type A, Type B, Combined) and returns the expanded formula string and the final uncertainty."""
    n = len(data)
    delta_xa = std_dev / np.sqrt(n) if n > 0 else 0
    delta_xb = delta_instrument / np.sqrt(3) if delta_instrument is not None else 0
    
    delta_x = np.sqrt(delta_xa**2 + delta_xb**2)
    
    uncertainty_log = []
    uncertainty_log.append(f"\\Delta X_A = \\frac{{\\sigma}}{{\\sqrt{{n}}}} = \\frac{{{std_dev:.4f}}}{{\\sqrt{{{n}}}}} = {delta_xa:.4f} {unit}")
    if delta_instrument is not None:
        uncertainty_log.append(f"\\Delta X_B = \\frac{{\\Delta_{{\\text{{inst}}}}}}{{\\sqrt{{3}}}} = \\frac{{{delta_instrument:.2f}}}{{\\sqrt{{3}}}} = {delta_xb:.4f} {unit}")
    else:
        uncertainty_log.append("Instrument uncertainty not provided, \\Delta X_B is 0.")
        
    uncertainty_log.append(f"\\Delta X = \\sqrt{{\\Delta X_A^2 + \\Delta X_B^2}} = \\sqrt{{{delta_xa**2:.8f} + {delta_xb**2:.8f}}} = {delta_x:.4f} {unit}")
    
    return delta_x, "\n".join(uncertainty_log)

def round_physics(value, uncertainty, unit=""):
    """Performs physics-specific rounding rules."""
    if uncertainty <= 0 or not math.isfinite(uncertainty):
        return f"({value:.3f} {unit})", uncertainty, "Invalid uncertainty value, no rounding performed."

    # Find position of the first significant digit of the uncertainty
    first_digit_pos = -int(math.floor(math.log10(uncertainty)))
    
    # Round uncertainty up to one significant digit
    rounded_uncertainty = math.ceil(uncertainty * (10**first_digit_pos)) / (10**first_digit_pos)
    
    # Round the mean to the same decimal place as the rounded uncertainty
    rounded_mean = round(value, first_digit_pos)
    
    format_str = f".{first_digit_pos}f"
    
    result_str = f"({rounded_mean:{format_str}} \\pm {rounded_uncertainty:{format_str}}) {unit}"
    
    rounding_log = (
        f"Original Mean: {value:.4f}, Original Uncertainty: {uncertainty:.4f}\n"
        f"Uncertainty \\Delta X rounded up to one significant digit: {rounded_uncertainty:{format_str}}\n"
        f"Mean \\bar{{x}} rounded to match \\Delta X decimal place: {rounded_mean:{format_str}}\n"
        f"Final Result: {result_str}"
    )
                   
    return result_str, rounded_uncertainty, rounding_log

def linear_fit_and_plot(x, y, xlabel='X', ylabel='Y', title='Linear Fit', save_path='fit_plot.png'):
    """Performs linear regression and creates a plot."""
    x = np.array(x)
    y = np.array(y)
    slope, intercept, r_value, p_value, std_err = stats.linregress(x, y)
    
    plt.figure(figsize=(10, 6))
    plt.scatter(x, y, color='blue', label='Original Data')
    fit_label = f'Fit: y={slope:.4f}x + {intercept:.4f}\n$R^2$={r_value**2:.4f}'
    plt.plot(x, slope*x + intercept, color='red', label=fit_label)
    
    plt.xlabel(xlabel)
    plt.ylabel(ylabel)
    plt.title(title)
    plt.legend()
    plt.grid(True)
    plt.savefig(save_path)
    plt.close()
    
    return {
        'slope': slope,
        'intercept': intercept,
        'r_squared': r_value**2,
        'std_err': std_err,
        'plot_path': save_path
    }

if __name__ == "__main__":
    print("--- Physics Data Processor Script Execution ---")
    
    # --- Sample Data --- #
    sample_data = [15.20, 15.18, 15.05, 15.15, 15.10, 15.22, 15.08, 15.17, 16.5]
    instrument_uncertainty = 0.02
    unit = "cm"
    
    print(f"\nInitial Data: {sample_data}")

    # --- Bad Value Test --- #
    print("\n--- 1. Bad Value Test (Grubbs' Test) ---")
    processed_data, final_mean_after_test, final_std_dev_after_test, bad_test_log = bad_value_test(sample_data, unit)
    print(bad_test_log)
    print(f"\nFinal Processed Data: {processed_data}")
    
    # --- Mean and Std Dev Calculation --- #
    print("\n--- 2. Mean and Standard Deviation Calculation ---")
    mean, mean_expanded_str = calculate_mean_expanded(processed_data, unit)
    print(mean_expanded_str)
    std_dev, std_dev_expanded_str = calculate_std_dev_expanded(processed_data, mean, unit)
    print(std_dev_expanded_str)
    
    # --- Uncertainty Calculation --- #
    print("\n--- 3. Uncertainty Calculation ---")
    delta_x, uncertainty_expanded_str = calculate_uncertainties_expanded(processed_data, std_dev, instrument_uncertainty, unit)
    print(uncertainty_expanded_str)
    
    # --- Rounding and Final Result --- #
    print("\n--- 4. Rounding and Final Result ---")
    result_str, rounded_unc, rounding_log = round_physics(mean, delta_x, unit)
    print(rounding_log)
    
    # --- Linear Fit Example --- #
    print("\n--- 5. Linear Fit and Plot ---")
    x_data = np.array([1, 2, 3, 4, 5, 6])
    y_data = np.array([2.1, 3.9, 6.2, 8.1, 9.8, 12.3])
    fit_results = linear_fit_and_plot(
        x_data, y_data, 
        xlabel='Time (s)', 
        ylabel='Position (m)', 
        title='Position vs. Time', 
        save_path='position_time_fit.png'
    )
    print("Linear fit results:")
    for key, value in fit_results.items():
        print(f"  {key}: {value}")
    print(f"Plot saved to: {fit_results['plot_path']}")
    print("\n--- Script Finished ---")
