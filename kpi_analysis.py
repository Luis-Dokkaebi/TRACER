import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import sys

# --- CONFIGURATION ---
DATA_FILE = 'mock_data.csv'

# Define Teams and Members
TEAMS = {
    'VENTAS': ['EDUARDO MANZANARES', 'SEBASTIAN PADILLA', 'RAMIRO RODRIGUEZ'],
    'TRACKER': ['JUDITH ECHAVARRIA', 'EDUARDO TERAN', 'ANGEL SALINAS']
}

# Flatten the list for filtering
ALL_MEMBERS = [member for team in TEAMS.values() for member in team]
MEMBER_TO_TEAM = {member: team for team, members in TEAMS.items() for member in members}

# Colors for visualization (Corporate/Sober)
PALETTE = {'VENTAS': '#2c3e50', 'TRACKER': '#e74c3c'}

def load_and_clean_data(filepath):
    """Loads data from CSV and performs cleaning and preprocessing."""
    try:
        df = pd.read_csv(filepath)
    except FileNotFoundError:
        print(f"Error: File {filepath} not found.")
        sys.exit(1)

    # Filter for specific users
    df['RESPONSABLE'] = df['RESPONSABLE'].str.upper().str.strip()
    df = df[df['RESPONSABLE'].isin(ALL_MEMBERS)].copy()

    # Map Users to Teams
    df['TEAM'] = df['RESPONSABLE'].map(MEMBER_TO_TEAM)

    # Filter for Completed Tasks
    completed_statuses = ['DONE', 'COMPLETED', 'FINALIZADO', 'TERMINADO']
    df['ESTATUS'] = df['ESTATUS'].str.upper().str.strip()
    df = df[df['ESTATUS'].isin(completed_statuses)].copy()

    # Date Handling
    # Explicitly specifying format to avoid warnings and ensure correctness
    for col in ['FECHA_INICIO', 'FECHA_FIN']:
        df[col] = pd.to_datetime(df[col], format='%d/%m/%y', errors='coerce')

    # Remove rows with invalid dates
    df = df.dropna(subset=['FECHA_INICIO', 'FECHA_FIN'])

    return df

def calculate_kpis(df):
    """Calculates Volume and Efficiency."""
    # Calculate Cycle Time (Days)
    df['CYCLE_TIME'] = (df['FECHA_FIN'] - df['FECHA_INICIO']).dt.days

    # Aggregation
    kpi_df = df.groupby(['RESPONSABLE', 'TEAM']).agg(
        TOTAL_TASKS=('ESTATUS', 'count'),
        AVG_CYCLE_TIME=('CYCLE_TIME', 'mean')
    ).reset_index()

    # Round Cycle Time for display
    kpi_df['AVG_CYCLE_TIME'] = kpi_df['AVG_CYCLE_TIME'].round(1)

    return kpi_df

def plot_efficiency(kpi_df):
    """Generates a bar chart for Efficiency (Avg Cycle Time)."""
    plt.figure(figsize=(10, 6))
    sns.set_theme(style="whitegrid")

    # Create Bar Plot
    ax = sns.barplot(
        data=kpi_df,
        x='RESPONSABLE',
        y='AVG_CYCLE_TIME',
        hue='TEAM',
        palette=PALETTE,
        dodge=False
    )

    plt.title('Eficiencia por Colaborador (Cycle Time Promedio)', fontsize=16, fontweight='bold', pad=20)
    plt.xlabel('Colaborador', fontsize=12)
    plt.ylabel('Días Promedio', fontsize=12)
    plt.xticks(rotation=45)
    plt.legend(title='Área')

    # Add labels on top of bars
    for p in ax.patches:
        height = p.get_height()
        if height > 0: # Avoid labeling empty bars
            ax.annotate(f'{height}',
                        (p.get_x() + p.get_width() / 2., height),
                        ha='center', va='bottom', fontsize=10, color='black', xytext=(0, 5),
                        textcoords='offset points')

    plt.tight_layout()
    plt.savefig('efficiency_chart.png')
    print("Graph 1 saved: efficiency_chart.png")

def plot_productivity(kpi_df):
    """Generates a horizontal bar chart for Total Productivity."""
    plt.figure(figsize=(10, 6))
    sns.set_theme(style="whitegrid")

    # Sort for better visualization
    kpi_df_sorted = kpi_df.sort_values('TOTAL_TASKS', ascending=False)

    ax = sns.barplot(
        data=kpi_df_sorted,
        y='RESPONSABLE',
        x='TOTAL_TASKS',
        hue='RESPONSABLE', # Fixed warning by assigning hue
        legend=False,
        palette=[PALETTE[t] for t in kpi_df_sorted['TEAM']] # Color by team
    )

    plt.title('Productividad Total (Tareas Terminadas)', fontsize=16, fontweight='bold', pad=20)
    plt.xlabel('Total Tareas', fontsize=12)
    plt.ylabel('Colaborador', fontsize=12)

    # Add labels
    for p in ax.patches:
        width = p.get_width()
        if width > 0:
            ax.annotate(f'{int(width)}',
                        (width, p.get_y() + p.get_height() / 2.),
                        ha='left', va='center', fontsize=10, color='black', xytext=(5, 0),
                        textcoords='offset points')

    plt.tight_layout()
    plt.savefig('productivity_chart.png')
    print("Graph 2 saved: productivity_chart.png")

def main():
    print("Starting KPI Analysis...")
    df = load_and_clean_data(DATA_FILE)

    if df.empty:
        print("No valid data found after filtering.")
        return

    kpi_df = calculate_kpis(df)

    print("\nCalculated KPIs:")
    print(kpi_df)

    print("\nGenerating Visualizations...")
    plot_efficiency(kpi_df)
    plot_productivity(kpi_df)
    print("Done.")

if __name__ == "__main__":
    main()
