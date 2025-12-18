import unittest
import pandas as pd
import os
import matplotlib.pyplot as plt
from kpi_analysis import extract_and_load_data, clean_data, calculate_kpis, generate_team_dashboard, SHEETS_VENTAS, SHEETS_TRACKER, PALETTE_VENTAS, DATA_FILE

class TestKPIAnalysis(unittest.TestCase):

    def setUp(self):
        """Setup test environment. Ensure data file exists."""
        if not os.path.exists(DATA_FILE):
            self.fail(f"Critical: Data file {DATA_FILE} not found. Run generate_mock_excel.py first.")

    def test_01_ingesta_datos(self):
        """Test Case 1: Ingesta de Datos (Data Integrity)"""
        print("\nRunning Test 1: Data Ingestion...")

        # A. Verificación de Hojas & Volumen de Datos - VENTAS
        df_ventas = extract_and_load_data(DATA_FILE, SHEETS_VENTAS)
        self.assertFalse(df_ventas.empty, "FAIL: El DataFrame de VENTAS está vacío.")
        self.assertTrue('Usuario' in df_ventas.columns, "FAIL: Columna 'Usuario' no encontrada en Ventas.")

        users_found_ventas = df_ventas['Usuario'].unique()
        for user in SHEETS_VENTAS:
            self.assertIn(user.upper(), users_found_ventas, f"FAIL: No se encontraron datos para {user} en Ventas.")

        # B. Verificación de Hojas & Volumen de Datos - TRACKER
        df_tracker = extract_and_load_data(DATA_FILE, SHEETS_TRACKER)
        self.assertFalse(df_tracker.empty, "FAIL: El DataFrame de TRACKER está vacío.")
        self.assertTrue('Usuario' in df_tracker.columns, "FAIL: Columna 'Usuario' no encontrada en Tracker.")

        users_found_tracker = df_tracker['Usuario'].unique()
        for user in SHEETS_TRACKER:
            self.assertIn(user.upper(), users_found_tracker, f"FAIL: No se encontraron datos para {user} en Tracker.")

        print("OK: Data Ingestion passed.")

    def test_02_logica_negocio(self):
        """Test Case 2: Lógica de Negocio (Transformation Logic)"""
        print("\nRunning Test 2: Business Logic...")

        # Load raw data
        df_raw = extract_and_load_data(DATA_FILE, SHEETS_VENTAS) # Testing with Ventas group
        df_clean = clean_data(df_raw)

        # A. Validación de Fechas (End >= Start)
        invalid_dates = df_clean[df_clean['FECHA_FIN'] < df_clean['FECHA_INICIO']]
        self.assertTrue(invalid_dates.empty, f"FAIL: Se encontraron {len(invalid_dates)} filas con Fecha Fin anterior a Fecha Inicio.")

        # B. Cálculo de KPIs (Manual Verification)
        if not df_clean.empty:
            sample_row = df_clean.iloc[0]
            start = sample_row['FECHA_INICIO']
            end = sample_row['FECHA_FIN']
            calculated_cycle = sample_row['CYCLE_TIME']

            manual_cycle = (end - start).days
            self.assertEqual(calculated_cycle, manual_cycle, f"FAIL: Cálculo de Cycle Time incorrecto. Script: {calculated_cycle}, Manual: {manual_cycle}")
        else:
            self.fail("FAIL: No data available to test KPI calculation.")

        print("OK: Business Logic passed.")

    def test_03_renderizado(self):
        """Test Case 3: Renderizado (Smoke Test)"""
        print("\nRunning Test 3: Rendering Smoke Test...")

        df_raw = extract_and_load_data(DATA_FILE, SHEETS_VENTAS)
        df_clean = clean_data(df_raw)
        kpi_df = calculate_kpis(df_clean)

        # A. Validar Retorno de Objeto Figura
        fig = generate_team_dashboard(kpi_df, "VENTAS", PALETTE_VENTAS, None) # None output to skip saving
        self.assertIsInstance(fig, plt.Figure, "FAIL: La función de renderizado no retornó un objeto Figure válido.")

        # B. Validar Segregación de Datos (Solo Ventas)
        # Inspect axes to check labels
        axes = fig.axes
        if axes:
            # Check the first ax (Efficiency)
            # seaborn barplot x-tick labels
            xticklabels = [label.get_text() for label in axes[0].get_xticklabels()]

            # Verify only Ventas people are in the chart
            expected_users = [u.upper() for u in SHEETS_VENTAS]
            for label in xticklabels:
                if label: # Ignore empty if any
                    self.assertIn(label, expected_users, f"FAIL: Gráfica de Ventas contiene un usuario inesperado: {label}")

        print("OK: Rendering passed.")

if __name__ == '__main__':
    unittest.main()
