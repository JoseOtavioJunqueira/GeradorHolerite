# Payslip Generator

A desktop tool (PySimpleGUI) that reads a pre-formatted spreadsheet of employee clock-in/clock-out times, computes pay — including overtime, night-shift, hazard pay (periculosidade), and unhealthy-work premiums (insalubridade) under Brazilian labor rules — and fills in a payslip template for each employee.

This was an early project built to automate a manual payroll process; it's kept here as-is rather than rewritten, as a snapshot of that stage of my learning.

## How It Works

1. **Login**: simple username/password gate, with a basic "register new user" flow guarded by an admin password.
2. **Hours table**: reads each employee's clock-in/clock-out hours from an Excel workbook (one sheet per day of the month).
3. **Calculations**: computes worked hours, overtime, night-shift hours, and applicable hazard/unhealthy-work premiums.
4. **Payslip generation**: fills in an Excel payslip template and saves one file per employee.

## Tech Stack

- Python
- pandas, openpyxl (spreadsheet I/O)
- PySimpleGUI (desktop UI)

## Getting Started

1. Clone the repository:
   ```bash
   git clone https://github.com/JoseOtavioJunqueira/GeradorHolerite.git
   cd GeradorHolerite
   ```

2. Create and activate a virtual environment (recommended):
   ```bash
   python -m venv venv
   source venv/bin/activate  # On Windows: venv\Scripts\activate
   ```

3. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
   > **Note:** PySimpleGUI moved to a paid license model after this project was built. If `pip install PySimpleGUI` fails for you, see their [current distribution instructions](https://www.pysimplegui.com/).

4. Provide `funcionarios.xlsx` (the hours table) and `Holerite.xlsx` (the payslip template) in the project folder, or point to your own via the `FUNCIONARIOS_XLSX` / `HOLERITE_TEMPLATE_XLSX` environment variables. Set `ADMIN_PASSWORD` to override the default admin password used for registering new users.

5. Run the app:
   ```bash
   python main.py
   ```

## Known Limitations

This was built as a learning project, not production payroll software:

- User credentials are stored in plain text (`usuarios.txt` / `senhas.txt`), not hashed.
- It assumes a specific spreadsheet layout for both the hours table and the payslip template.
- No automated tests.

## License

MIT — see [LICENSE](LICENSE).

## Contact

José Otávio — joseotavio.jr1104@gmail.com
