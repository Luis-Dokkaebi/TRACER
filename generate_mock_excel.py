import pandas as pd
import random
from datetime import datetime, timedelta

# Configuration
OUTPUT_FILE = 'project_data.xlsx'

GROUPS = {
    'VENTAS': ['Eduardo Manzanares', 'Sebastian Padilla', 'Ramiro Rodriguez'],
    'TRACKER': ['Judith Echavarria', 'Eduardo Teran', 'Angel Salinas']
}

STATUS_OPTIONS = ['Completed', 'Done', 'In Progress', 'Pending', 'Terminado']

def generate_random_dates(n=10):
    dates = []
    for _ in range(n):
        start = datetime(2023, 1, 1) + timedelta(days=random.randint(0, 180))
        duration = random.randint(1, 15)
        end = start + timedelta(days=duration)
        dates.append((start, end))
    return dates

def create_mock_excel():
    with pd.ExcelWriter(OUTPUT_FILE, engine='openpyxl') as writer:
        for group, members in GROUPS.items():
            for member in members:
                # Generate Data
                data = {
                    'ESTATUS': [random.choice(STATUS_OPTIONS) for _ in range(10)],
                }

                # Add dates
                date_pairs = generate_random_dates(10)
                data['FECHA_INICIO'] = [d[0].strftime('%d/%m/%y') for d in date_pairs]
                data['FECHA_FIN'] = [d[1].strftime('%d/%m/%y') for d in date_pairs]

                # Introduce some None/Empty values for robustness testing
                if random.random() > 0.8:
                    data['FECHA_FIN'][0] = None

                df = pd.DataFrame(data)

                # Write to sheet named after the user
                df.to_excel(writer, sheet_name=member, index=False)
                print(f"Created sheet: {member}")

    print(f"Successfully generated {OUTPUT_FILE}")

if __name__ == "__main__":
    create_mock_excel()
