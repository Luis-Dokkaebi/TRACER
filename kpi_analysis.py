import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import sys

# --- CONFIGURATION ---
DATA_FILE = 'mock_data.csv'

# Define Teams and Members explicitly
MEMBERS_VENTAS = ['EDUARDO MANZANARES', 'SEBASTIAN PADILLA', 'RAMIRO RODRIGUEZ']
MEMBERS_TRACKER = ['JUDITH ECHAVARRIA', 'EDUARDO TERAN', 'ANGEL SALINAS']

# Colors
PALETTE_VENTAS = '#2c3e50'
PALETTE_TRACKER = '#e74c3c'

def load_and_clean_data(filepath):
    """Loads data from CSV and performs basic cleaning (dates, whitespace)."""
    try:
        df = pd.read_csv(filepath)
    except FileNotFoundError:
        print(f"Error: File {filepath} not found.")
        sys.exit(1)

    # Standardize string columns
    df['RESPONSABLE'] = df['RESPONSABLE'].str.upper().str.strip()
    df['ESTATUS'] = df['ESTATUS'].str.upper().str.strip()

    # Filter for Completed Tasks
    completed_statuses = ['DONE', 'COMPLETED', 'FINALIZADO', 'TERMINADO']
    df = df[df['ESTATUS'].isin(completed_statuses)].copy()

    # Date Handling
    for col in ['FECHA_INICIO', 'FECHA_FIN']:
        df[col] = pd.to_datetime(df[col], format='%d/%m/%y', errors='coerce')

    # Remove rows with invalid dates
    df = df.dropna(subset=['FECHA_INICIO', 'FECHA_FIN'])

    # Calculate Cycle Time
    df['CYCLE_TIME'] = (df['FECHA_FIN'] - df['FECHA_INICIO']).dt.days

    return df

def get_team_dataframe(full_df, members):
    """Filters the full dataframe for specific members."""
    return full_df[full_df['RESPONSABLE'].isin(members)].copy()

def calculate_kpis(df):
    """Calculates aggregated KPIs for the given dataframe."""
    if df.empty:
        return pd.DataFrame(columns=['RESPONSABLE', 'TOTAL_TASKS', 'AVG_CYCLE_TIME'])

    kpi_df = df.groupby('RESPONSABLE').agg(
        TOTAL_TASKS=('ESTATUS', 'count'),
        AVG_CYCLE_TIME=('CYCLE_TIME', 'mean')
    ).reset_index()

    kpi_df['AVG_CYCLE_TIME'] = kpi_df['AVG_CYCLE_TIME'].round(1)
    return kpi_df

def generate_team_dashboard(kpi_df, team_name, color, output_file):
    """Generates a 2-subplot dashboard for a specific team."""
    if kpi_df.empty:
        print(f"Warning: No data for {team_name}, skipping dashboard.")
        return

    # Setup Figure with 2 subplots side-by-side
    fig, axes = plt.subplots(1, 2, figsize=(14, 6))
    fig.suptitle(f'Dashboard Operativo: {team_name}', fontsize=20, fontweight='bold')
    sns.set_theme(style="whitegrid")

    # Plot 1: Efficiency (Cycle Time)
    sns.barplot(
        data=kpi_df,
        x='RESPONSABLE',
        y='AVG_CYCLE_TIME',
        color=color,
        ax=axes[0]
    )
    axes[0].set_title('Eficiencia (Promedio Días / Tarea)', fontsize=14)
    axes[0].set_xlabel('')
    axes[0].set_ylabel('Días')
    axes[0].tick_params(axis='x', rotation=15)

    # Add values to bars
    for p in axes[0].patches:
        height = p.get_height()
        if height > 0:
            axes[0].annotate(f'{height}',
                             (p.get_x() + p.get_width() / 2., height),
                             ha='center', va='bottom', fontsize=11, color='black', xytext=(0, 5),
                             textcoords='offset points')

    # Plot 2: Productivity (Volume)
    # Sort by Volume for better visual
    kpi_df_sorted = kpi_df.sort_values('TOTAL_TASKS', ascending=False)

    sns.barplot(
        data=kpi_df_sorted,
        x='RESPONSABLE',
        y='TOTAL_TASKS',
        color=color,
        ax=axes[1]
    )
    axes[1].set_title('Productividad (Volumen de Tareas)', fontsize=14)
    axes[1].set_xlabel('')
    axes[1].set_ylabel('Tareas Terminadas')
    axes[1].tick_params(axis='x', rotation=15)

    # Add values to bars
    for p in axes[1].patches:
        height = p.get_height()
        if height > 0:
            axes[1].annotate(f'{int(height)}',
                             (p.get_x() + p.get_width() / 2., height),
                             ha='center', va='bottom', fontsize=11, color='black', xytext=(0, 5),
                             textcoords='offset points')

    plt.tight_layout(rect=[0, 0.03, 1, 0.95]) # Adjust for suptitle
    plt.savefig(output_file)
    print(f"Generated {output_file}")
    plt.close()

def main():
    print("Starting Independent Team Analysis...")

    # 1. Load Data
    full_df = load_and_clean_data(DATA_FILE)

    # 2. Process Data Separately (Requirement 1)
    df_ventas = get_team_dataframe(full_df, MEMBERS_VENTAS)
    df_tracker = get_team_dataframe(full_df, MEMBERS_TRACKER)

    # 3. Calculate KPIs Separately
    kpi_ventas = calculate_kpis(df_ventas)
    kpi_tracker = calculate_kpis(df_tracker)

    # Print Tables (Requirement 3: First Ventas, then Tracker)
    print("\n--- TABLA KPI: VENTAS ---")
    print(kpi_ventas.to_string(index=False))

    print("\n--- TABLA KPI: TRACKER ---")
    print(kpi_tracker.to_string(index=False))

    # 4. Generate Visualization Separately (Requirement 2)
    # Figure A: Ventas
    generate_team_dashboard(kpi_ventas, "VENTAS", PALETTE_VENTAS, 'dashboard_VENTAS.png')

    # Figure B: Tracker
    generate_team_dashboard(kpi_tracker, "TRACKER", PALETTE_TRACKER, 'dashboard_TRACKER.png')

    print("\nAnalysis Complete.")

if __name__ == "__main__":
    main()
