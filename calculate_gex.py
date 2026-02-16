import pandas as pd
import numpy as np
from scipy.stats import norm
import matplotlib
matplotlib.use('Agg')  # Use non-interactive backend
import matplotlib.pyplot as plt
from datetime import datetime

# ===== INPUT CONSTANTS =====
VALUATION_DATE = "2026-01-05"
EXPIRATION_DATE = "2026-03-20"
SPOT = 25600
IV = 0.1767
MULT = 20
T = 0.202739726  # Time to expiration in years

# ===== END CONSTANTS =====

def d1(S, K, r=0, T=0.25, sigma=0.2):
    """Calculate d1 in Black-Scholes formula"""
    return (np.log(S / K) + (r + 0.5 * sigma ** 2) * T) / (sigma * np.sqrt(T))

def d2(d1_val, sigma=0.2, T=0.25):
    """Calculate d2 in Black-Scholes formula"""
    return d1_val - sigma * np.sqrt(T)

def calculate_gamma(S, K, r=0, T=0.25, sigma=0.2):
    """Calculate gamma for an option using Black-Scholes model"""
    if T <= 0:
        return 0
    d1_val = d1(S, K, r, T, sigma)
    gamma = norm.pdf(d1_val) / (S * sigma * np.sqrt(T))
    return gamma

def calculate_signed_gex(row, spot, iv, mult, t):
    """
    Calculate signed GEX for a single option row.
    
    GEX = Gamma × OI × Spot² × (1/100) × Multiplier
    Sign: Calls are positive, Puts are negative
    """
    option_type = row['OptionType']
    strike = str(row['Strike']).replace(',', '')
    strike = float(strike)
    
    oi = str(row['OI']).replace(',', '')
    oi = float(oi)
    
    # Skip if OI is 0
    if oi == 0:
        return 0
    
    # Calculate gamma
    gamma = calculate_gamma(spot, strike, r=0, T=t, sigma=iv)
    
    # Calculate GEX: Gamma × OI × Spot × (1/100) × Multiplier
    # Note: GEX is typically normalized per 1% move in spot
    gex = gamma * oi * spot * (1/100) * mult
    
    # Apply sign: Calls positive, Puts negative
    if option_type == 'Put':
        gex = -gex
    
    return gex

def main():
    """Main function to calculate and graph GEX"""
    
    # Read the CSV file
    try:
        df = pd.read_csv('data.csv')
    except FileNotFoundError:
        print("Error: data.csv not found. Please run parse_options_data.py first.")
        return
    
    # Calculate GEX for each row
    df['GEX'] = df.apply(lambda row: calculate_signed_gex(row, SPOT, IV, MULT, T), axis=1)
    
    # Group by strike and sum GEX (to get net GEX at each strike)
    gex_by_strike = df.groupby('Strike')['GEX'].sum().sort_index()
    
    # Print summary
    print("=" * 80)
    print(f"GEX Analysis")
    print("=" * 80)
    print(f"Valuation Date: {VALUATION_DATE}")
    print(f"Expiration Date: {EXPIRATION_DATE}")
    print(f"Spot: {SPOT}")
    print(f"IV: {IV}")
    print(f"Multiplier: {MULT}")
    print(f"Time to Expiration (Years): {T}")
    print("=" * 80)
    print(f"\nTotal Signed GEX (sum of all options): {df['GEX'].sum():,.2f}")
    print(f"Call GEX: {df[df['OptionType'] == 'Call']['GEX'].sum():,.2f}")
    print(f"Put GEX: {df[df['OptionType'] == 'Put']['GEX'].sum():,.2f}")
    print("\n" + "=" * 80)
    
    # Save detailed GEX data to CSV
    df_output = df[['OptionType', 'Strike', 'OI', 'GEX']].copy()
    df_output.to_csv('gex_data.csv', index=False)
    print(f"\nDetailed GEX data saved to gex_data.csv")
    
    # Create visualization
    fig, axes = plt.subplots(2, 1, figsize=(14, 10))
    
    # Plot 1: Signed GEX by strike
    ax1 = axes[0]
    colors = ['green' if x >= 0 else 'red' for x in gex_by_strike.values]
    ax1.bar(gex_by_strike.index, gex_by_strike.values, color=colors, alpha=0.7, edgecolor='black')
    ax1.axhline(y=0, color='black', linestyle='-', linewidth=0.5)
    ax1.axvline(x=SPOT, color='blue', linestyle='--', linewidth=2, label=f'Spot: {SPOT}')
    ax1.set_xlabel('Strike Price', fontsize=12)
    ax1.set_ylabel('Signed GEX', fontsize=12)
    ax1.set_title('Signed Gamma Exposure (GEX) by Strike', fontsize=14, fontweight='bold')
    ax1.grid(True, alpha=0.3)
    ax1.legend()
    
    # Plot 2: Cumulative GEX
    cumulative_gex = gex_by_strike.cumsum()
    ax2 = axes[1]
    ax2.plot(cumulative_gex.index, cumulative_gex.values, marker='o', linewidth=2, markersize=6, color='darkblue')
    ax2.axhline(y=0, color='black', linestyle='-', linewidth=0.5)
    ax2.axvline(x=SPOT, color='blue', linestyle='--', linewidth=2, label=f'Spot: {SPOT}')
    ax2.fill_between(cumulative_gex.index, cumulative_gex.values, 0, alpha=0.3)
    ax2.set_xlabel('Strike Price', fontsize=12)
    ax2.set_ylabel('Cumulative GEX', fontsize=12)
    ax2.set_title('Cumulative Gamma Exposure by Strike', fontsize=14, fontweight='bold')
    ax2.grid(True, alpha=0.3)
    ax2.legend()
    
    plt.tight_layout()
    plt.savefig('gex_analysis.png', dpi=300, bbox_inches='tight')
    print(f"GEX chart saved to gex_analysis.png")

if __name__ == "__main__":
    main()
