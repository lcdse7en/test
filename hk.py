from typing import Dict, List

import matplotlib.pyplot as plt
import pandas as pd
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

class Loan:
    def __init__(
        self,
        principal: float,
        annual_rate: float,
        years: int,
        payments_per_year: int = 12,
        loan_type: str = "equal_payment",  # "equal_payment" 或 "equal_principal"
    ):
        self.principal = principal
        self.annual_rate = annual_rate
        self.years = years
        self.payments_per_year = payments_per_year
        self.loan_type = loan_type

        self.period_rate = self.annual_rate / self.payments_per_year
        self.total_periods = self.years * self.payments_per_year

        if self.loan_type == "equal_payment":
            self.payment = self._calculate_equal_payment()
        else:
            self.payment = None

    def _calculate_equal_payment(self) -> float:
        i = self.period_rate
        N = self.total_periods
        A = (self.principal * i * (1 + i) ** N) / ((1 + i) ** N - 1)
        return A

    def generate_schedule(self) -> List[Dict]:
        balance = self.principal
        schedule = []

        N = self.total_periods
        i = self.period_rate
        equal_principal = balance / N if self.loan_type == "equal_principal" else None

        for period in range(1, int(N) + 1):
            interest = balance * i

            if self.loan_type == "equal_payment":
                principal_payment = self.payment - interest
                total_payment = self.payment
            else:
                principal_payment = equal_principal
                total_payment = principal_payment + interest

            balance -= principal_payment
            balance = max(balance, 0)

            schedule.append(
                {
                    "Period": period,
                    "Principal Payment": round(principal_payment, 2),
                    "Interest Payment": round(interest, 2),
                    "Total Payment": round(total_payment, 2),
                    "Remaining Balance": round(balance, 2),
                }
            )

        return schedule

    def summary(self):
        schedule = self.generate_schedule()
        total_payment = sum(item["Total Payment"] for item in schedule)
        total_interest = sum(item["Interest Payment"] for item in schedule)

        print("\n========== Loan Summary ==========")
        print(f"贷款本金：{self.principal} 元")
        print(f"年利率：{self.annual_rate * 100}%")
        print(f"贷款年限：{self.years} 年（共 {int(self.total_periods)} 期）")
        print(
            f"还款方式：{'等额本息' if self.loan_type == 'equal_payment' else '等额本金'}"
        )
        if self.loan_type == "equal_payment":
            print(f"每期还款：{self.payment:.2f} 元")
        print(f"总还款：{total_payment:.2f} 元")
        print(f"总利息：{total_interest:.2f} 元")
        print("===================================")

    def print_schedule(self):
        schedule = self.generate_schedule()

        print(
            f"\n{'期数':<5} {'本金':>12} {'利息':>12} {'总还款':>12} {'剩余贷款':>15}"
        )
        for item in schedule:
            print(
                f"{item['Period']:<5} {item['Principal Payment']:>12.2f} {item['Interest Payment']:>12.2f} "
                f"{item['Total Payment']:>12.2f} {item['Remaining Balance']:>15.2f}"
            )

    def export_to_excel(self, filename: str = "loan_schedule.xlsx"):
        schedule = self.generate_schedule()
        df = pd.DataFrame(schedule)

        # 先四舍五入金额列两位小数，避免浮点显示问题
        money_cols = [
            "Principal Payment",
            "Interest Payment",
            "Total Payment",
            "Remaining Balance",
        ]
        for col in money_cols:
            df[col] = df[col].round(2)

        with pd.ExcelWriter(filename, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="Loan Schedule", index=False)
            workbook = writer.book
            worksheet = writer.sheets["Loan Schedule"]

            summary_info = [
                ["贷款本金", self.principal],
                ["年利率", f"{self.annual_rate * 100}%"],
                ["贷款年限", f"{self.years} 年"],
                [
                    "还款方式",
                    "等额本息" if self.loan_type == "equal_payment" else "等额本金",
                ],
                ["每期还款", self.payment if self.payment else "不同"],
                ["总还款", df["Total Payment"].sum().round(2)],
                ["总利息", df["Interest Payment"].sum().round(2)],
            ]

            summary_start = len(df) + 4
            for idx, (label, value) in enumerate(summary_info):
                worksheet.cell(row=summary_start + idx, column=1, value=label)
                worksheet.cell(row=summary_start + idx, column=2, value=value)

            header_font = Font(bold=True, color="FFFFFF")
            header_fill = PatternFill("solid", fgColor="4F81BD")
            thin_border = Border(
                left=Side(style="thin"),
                right=Side(style="thin"),
                top=Side(style="thin"),
                bottom=Side(style="thin"),
            )

            for cell in worksheet[1]:
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = Alignment(horizontal="center", vertical="center")

            for row in worksheet.iter_rows(
                min_row=1,
                max_row=worksheet.max_row,
                min_col=1,
                max_col=worksheet.max_column,
            ):
                for cell in row:
                    cell.border = thin_border

            col_index_map = {
                col: idx + 1 for idx, col in enumerate(df.columns)
            }  # 列名映射列号

            # “Period”列居中
            period_col = "Period"
            period_col_letter = get_column_letter(col_index_map[period_col])
            for row in range(2, len(df) + 2):
                cell = worksheet[f"{period_col_letter}{row}"]
                cell.alignment = Alignment(horizontal="center", vertical="center")

            # 金额列千分位+两位小数，右对齐
            for col in money_cols:
                col_letter = get_column_letter(col_index_map[col])
                for row in range(2, len(df) + 2):
                    cell = worksheet[f"{col_letter}{row}"]
                    cell.number_format = "#,##0.00"
                    cell.alignment = Alignment(horizontal="right", vertical="center")

            # 自动列宽
            for column_cells in worksheet.columns:
                length = max(
                    len(str(cell.value)) if cell.value is not None else 0
                    for cell in column_cells
                )
                worksheet.column_dimensions[
                    get_column_letter(column_cells[0].column)
                ].width = length + 2

            # 冻结表头
            worksheet.freeze_panes = "A2"

            print(
                f"\n✅ 已导出美化 Excel（期数列居中，金额列千分位+两位小数）：{filename}"
            )

    def plot(self, save_path: str = "loan_plot.png"):
        schedule = self.generate_schedule()
        df = pd.DataFrame(schedule)

        plt.figure(figsize=(10, 6))

        plt.plot(
            df["Period"],
            df["Remaining Balance"],
            label="Remaining Balance",
            color="blue",
            linewidth=2,
        )
        plt.plot(
            df["Period"],
            df["Principal Payment"],
            label="Principal Payment",
            color="green",
            linestyle="--",
        )
        plt.plot(
            df["Period"],
            df["Interest Payment"],
            label="Interest Payment",
            color="red",
            linestyle=":",
        )

        plt.title("Loan Payment Breakdown Over Time")
        plt.xlabel("Period")
        plt.ylabel("Amount (元)")
        plt.grid(True, linestyle="--", alpha=0.7)
        plt.legend()
        plt.tight_layout()

        plt.savefig(save_path, dpi=300)
        print(f"✅ 图表已保存为：{save_path}")

        plt.show()

def main():
    loan = Loan(
        principal=1000000, annual_rate=0.048, years=30, loan_type="equal_payment"
    )

    loan.summary()
    loan.print_schedule()
    loan.export_to_excel("loan_schedule.xlsx")
    loan.plot("loan_schedule_plot.png")

if __name__ == "__main__":
    main()
