import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import sys

# --- CONFIGURATION ---
DATA_FILE = 'project_data.xlsx'

# Define Teams and Members explicitly (Source Mapping)
SHEETS_VENTAS = ['Eduardo Manzanares', 'Sebastian Padilla', 'Ramiro Rodriguez']
SHEETS_TRACKER = ['Judith Echavarria', 'Eduardo Teran', 'Angel Salinas']

# Colors
PALETTE_VENTAS = '#2c3e50'
PALETTE_TRACKER = '#e74c3c'

def extract_and_load_data(filepath, sheet_names):
    """
    Protocolo de Extracción e Ingesta:
    Iterates over sheet_names, reads them, adds 'Usuario' column, and concatenates.
    """
    dfs = []
    try:
        # Read the Excel file to check available sheets first (optional but good for debugging)
        xl = pd.ExcelFile(filepath)
        available_sheets = xl.sheet_names
        print(f"Available sheets: {available_sheets}")
    except FileNotFoundError:
        print(f"Error: File {filepath} not found.")
        sys.exit(1)

    for name in sheet_names:
        if name in available_sheets:
            print(f"Reading sheet: {name}")
            try:
                # Paso B: Lectura
                df = pd.read_excel(filepath, sheet_name=name)

                # Paso C: Normalización (Agregar columna Usuario)
                df['Usuario'] = name.upper()

                dfs.append(df)
            except Exception as e:
                print(f"Error reading sheet {name}: {e}")
        else:
            print(f"Warning: Sheet '{name}' not found in {filepath}")

    if not dfs:
        return pd.DataFrame()

    # Paso D: Fusión
    return pd.concat(dfs, ignore_index=True)

def clean_data(df):
    """Performs standard cleaning on the consolidated DataFrame."""
    if df.empty:
        return df

    # Standardize string columns
    if 'ESTATUS' in df.columns:
        df['ESTATUS'] = df['ESTATUS'].astype(str).str.upper().str.strip()

    # Filter for Completed Tasks
    completed_statuses = ['DONE', 'COMPLETED', 'FINALIZADO', 'TERMINADO']
    df = df[df['ESTATUS'].isin(completed_statuses)].copy()

    # Date Handling
    # Assuming the Excel file might have actual datetime objects or strings
    for col in ['FECHA_INICIO', 'FECHA_FIN']:
        df[col] = pd.to_datetime(df[col], dayfirst=True, errors='coerce')

    # Remove rows with invalid dates
    df = df.dropna(subset=['FECHA_INICIO', 'FECHA_FIN'])

    # Calculate Cycle Time
    df['CYCLE_TIME'] = (df['FECHA_FIN'] - df['FECHA_INICIO']).dt.days

    return df

def calculate_kpis(df):
    """Calculates aggregated KPIs from the master dataframe."""
    if df.empty:
        return pd.DataFrame(columns=['Usuario', 'TOTAL_TASKS', 'AVG_CYCLE_TIME'])

    # Group by 'Usuario' (The normalized column we added)
    kpi_df = df.groupby('Usuario').agg(
        TOTAL_TASKS=('ESTATUS', 'count'),
        AVG_CYCLE_TIME=('CYCLE_TIME', 'mean')
    ).reset_index()

    kpi_df['AVG_CYCLE_TIME'] = kpi_df['AVG_CYCLE_TIME'].round(1)
    return kpi_df

def generate_team_dashboard(kpi_df, team_name, color, output_file):
    """Generates a 2-subplot dashboard for a specific team."""
    if kpi_df.empty:
        print(f"Warning: No data for {team_name}, skipping dashboard.")
        return None

    # Setup Figure with 2 subplots side-by-side
    fig, axes = plt.subplots(1, 2, figsize=(14, 6))
    fig.suptitle(f'Dashboard Operativo: {team_name}', fontsize=20, fontweight='bold')
    sns.set_theme(style="whitegrid")

    # Plot 1: Efficiency (Cycle Time)
    sns.barplot(
        data=kpi_df,
        x='Usuario',
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
        x='Usuario',
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
    if output_file:
        plt.savefig(output_file)
        print(f"Generated {output_file}")
    plt.close()

    return fig

def main():
    print("Starting ETL & KPI Analysis...")

    # 1. ETL: Extract & Load (Steps A, B, C, D)
    print("\n--- Processing Group: VENTAS ---")
    df_sales_raw = extract_and_load_data(DATA_FILE, SHEETS_VENTAS)
    df_sales_master = clean_data(df_sales_raw)

    print("\n--- Processing Group: TRACKER ---")
    df_tracker_raw = extract_and_load_data(DATA_FILE, SHEETS_TRACKER)
    df_tracker_master = clean_data(df_tracker_raw)

    # 3. Execution: Calculate KPIs & Graph
    kpi_ventas = calculate_kpis(df_sales_master)
    kpi_tracker = calculate_kpis(df_tracker_master)

    # Print Tables
    print("\n--- TABLA KPI: VENTAS ---")
    print(kpi_ventas.to_string(index=False))

    print("\n--- TABLA KPI: TRACKER ---")
    print(kpi_tracker.to_string(index=False))

    # Generate Visualization
    generate_team_dashboard(kpi_ventas, "VENTAS", PALETTE_VENTAS, 'dashboard_VENTAS.png')
    generate_team_dashboard(kpi_tracker, "TRACKER", PALETTE_TRACKER, 'dashboard_TRACKER.png')

    print("\nAnalysis Complete.")

if __name__ == "__main__":
    main()
