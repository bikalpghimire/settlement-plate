import matplotlib.pyplot as plt
import pandas as pd
import math
import os
from matplotlib.backends.backend_pdf import PdfPages

# Step 1: Find all .xlsx files in the current directory
xlsx_files = [f for f in os.listdir() if f.endswith('.xlsx')]

# If no Excel files found, exit
if not xlsx_files:
    print("No .xlsx files found in the current directory.")
    input("Press Enter to exit...")
    exit()

# Step 2: Let user choose one
print("Available Excel files:\n")
for idx, file in enumerate(xlsx_files, 1):
    print(f"{idx}. {file}")

choice = input(f"\nEnter the number of the file to use (1-{len(xlsx_files)}): ").strip()

# Validate choice
if not choice.isdigit() or not (1 <= int(choice) <= len(xlsx_files)):
    print("Invalid choice. Exiting.")
    input("Press Enter to exit...")
    exit()

excel_file = xlsx_files[int(choice) - 1]
print(f"\n✅ Selected file: {excel_file}")

# 🆕 Create output PDF name from input filename
base_name = os.path.splitext(os.path.basename(excel_file))[0]
output_pdf = f'Plot_{base_name}.pdf'

# Step 3: Get all sheet names
sheet_names = pd.ExcelFile(excel_file).sheet_names

# Step 4: Generate plots
with PdfPages(output_pdf) as pdf:
    for tab in sheet_names:
        df = pd.read_excel(excel_file, sheet_name=tab)

        df['date'] = pd.to_datetime(df['date'], format='%d-%b-%y')
        df['days'] = (df['date'] - df['date'].iloc[0]).dt.days

        fig, ax = plt.subplots(figsize=(14, 7))
        ax.plot(df['days'], df['height_m'], marker='o', color='blue', label='Height (m)')
        ax.plot(df['days'], df['settlement_cm'], marker='o', color='red', label='Settlement (cm)')

        for i, (x, y) in enumerate(zip(df['days'], df['height_m'])):
            if i > 0:
                ax.text(x, y + 1, f'{y:.2f}', color='blue', fontsize=9, ha='center', va='bottom', rotation=90,
                        bbox=dict(facecolor='white', edgecolor='none', boxstyle='round,pad=0.2', alpha=0.7))
        for i, (x, y) in enumerate(zip(df['days'], df['settlement_cm'])):
            if i > 0:
                ax.text(x, y - 1, f'{y:.2f}', color='red', fontsize=9, ha='center', va='top', rotation=90,
                        bbox=dict(facecolor='white', edgecolor='none', boxstyle='round,pad=0.2', alpha=0.7))

        ax.axhline(0, color='black', linewidth=1)
        ymin = df['settlement_cm'].min()
        ymax = df['height_m'].max()
        ax.set_ylim(math.floor(ymin / 5) * 5 - 5, math.ceil(ymax / 5) * 5 + 5)
        ax.set_xlim(left=df['days'].iloc[0], right=df['days'].iloc[-1] + 1)

        ax.set_xlabel('Measurement Date')
        ax.set_xticks(df['days'])
        ax.set_xticklabels(df['date'].dt.strftime('%d-%b-%y'), rotation=90)

        ax2 = ax.twiny()
        ax2.set_xlim(ax.get_xlim())
        ax2.set_xticks(df['days'])
        ax2.set_xticklabels(df['days'], rotation=90)
        ax2.set_xlabel('Cumulative Days')

        ax.text(df['days'].iloc[0]+2, ymax * 0.5, 'Height (m)', color='blue', fontsize=12, rotation=90, va='center')
        ax.text(df['days'].iloc[0]+2, ymin * 0.5, 'Settlement (cm)', color='red', fontsize=12, rotation=90, va='center')

        ax.grid(True)
        ax.set_title(f'Height and Settlement vs. Time ({tab})')
        ax.legend()
        plt.tight_layout()
        pdf.savefig(fig)
        plt.close()

print(f"\n📄 PDF saved as '{output_pdf}'")
input("\nPress Enter to exit...")
