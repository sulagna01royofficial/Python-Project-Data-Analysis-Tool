print("PYTHON COURSE DAY30-PYTHON DATA ANALYSIS PROJECT USING AI")

import os
import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.pyplot as plt

# ==========================================================
# DATA ANALYSIS PROJECT USING AI
# Corporate GUI-Based Data Analysis Tool
# Author: Generated from Specification
# ==========================================================


class DataAnalysisApp:

    def __init__(self, root):

        self.root = root
        self.root.title("Data Analysis Tool")
        self.root.geometry("1400x850")

        self.df = None
        self.report_df = None
        self.file_path = ""
        self.chart_figure = None
        self.text_columns = []
        self.numeric_columns = []

        self.create_widgets()

    # ======================================================
    # GUI DESIGN
    # ======================================================

    def create_widgets(self):

        top_frame = ttk.LabelFrame(
            self.root,
            text="File Selection"
        )
        top_frame.pack(fill="x", padx=10, pady=5)

        self.file_var = tk.StringVar()

        ttk.Entry(
            top_frame,
            textvariable=self.file_var,
            width=100
        ).pack(side="left", padx=5, pady=5)

        ttk.Button(
            top_frame,
            text="Browse",
            command=self.browse_file
        ).pack(side="left", padx=5)

        ttk.Button(
            top_frame,
            text="Read",
            command=self.read_file
        ).pack(side="left", padx=5)

        # --------------------------------------------------

        info_frame = ttk.LabelFrame(
            self.root,
            text="Dataset Information"
        )
        info_frame.pack(fill="x", padx=10, pady=5)

        self.info_text = tk.Text(
            info_frame,
            height=1
        )
        self.info_text.pack(fill="x", padx=5, pady=5)

        # --------------------------------------------------

        options_frame = ttk.LabelFrame(
            self.root,
            text="Analysis Options"
        )
        options_frame.pack(fill="x", padx=10, pady=5)

        self.clean_var = tk.BooleanVar()

        tk.Checkbutton(
            options_frame,
            text="Clean Data",
            variable=self.clean_var,
            font=("Arial", 10)
        ).grid(row=0, column=0, padx=10, pady=5)

        ttk.Label(
            options_frame,
            text="Group By"
        ).grid(row=0, column=1, padx=5)

        self.group_combo = ttk.Combobox(
            options_frame,
            state="readonly",
            width=25
        )
        self.group_combo.grid(row=0, column=2, padx=5)

        ttk.Label(
            options_frame,
            text="Aggregation"
        ).grid(row=0, column=3, padx=5)

        self.agg_combo = ttk.Combobox(
            options_frame,
            state="readonly",
            width=15,
            values=[
                "sum",
                "mean",
                "average",
                "max",
                "min",
                "count",
                "median"
            ]
        )
        self.agg_combo.grid(row=0, column=4, padx=5)

        ttk.Label(
            options_frame,
            text="Value Column"
        ).grid(row=0, column=5, padx=5)

        self.value_combo = ttk.Combobox(
            options_frame,
            state="readonly",
            width=20
        )
        self.value_combo.grid(row=0, column=6, padx=5)

        ttk.Button(
            options_frame,
            text="Preview Report",
            command=self.preview_report
        ).grid(row=0, column=7, padx=10)

        # --------------------------------------------------

        export_frame = ttk.LabelFrame(
            self.root,
            text="Export Report"
        )
        export_frame.pack(fill="x", padx=10, pady=5)

        self.export_type = tk.StringVar(value="Excel")

        ttk.Radiobutton(
            export_frame,
            text="Excel (.xlsx)",
            variable=self.export_type,
            value="Excel"
        ).pack(side="left", padx=10)

        ttk.Radiobutton(
            export_frame,
            text="CSV (.csv)",
            variable=self.export_type,
            value="CSV"
        ).pack(side="left", padx=10)

        ttk.Button(
            export_frame,
            text="Export Report",
            command=self.export_report
        ).pack(side="left", padx=20)

        # --------------------------------------------------

        chart_frame = ttk.LabelFrame(
            self.root,
            text="Chart Builder"
        )
        chart_frame.pack(fill="x", padx=10, pady=5)

        ttk.Label(
            chart_frame,
            text="Chart Type"
        ).pack(side="left", padx=10)

        self.chart_combo = ttk.Combobox(
            chart_frame,
            state="readonly",
            values=[
                "Bar Chart",
                "Column Chart",
                "Line Chart",
                "Pie Chart"
            ],
            width=15
        )
        self.chart_combo.pack(side="left", padx=10)

        ttk.Button(
            chart_frame,
            text="Preview Chart",
            command=self.preview_chart
        ).pack(side="left", padx=10)

        ttk.Button(
            chart_frame,
            text="Export Chart",
            command=self.export_chart
        ).pack(side="left", padx=10)

        # --------------------------------------------------

        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill="both", expand=True, padx=10, pady=5)

        report_frame = ttk.LabelFrame(
            main_frame,
            text="Report Preview"
        )
        report_frame.pack(
            side="left",
            fill="both",
            expand=True)

        self.tree = ttk.Treeview(report_frame)

        yscroll = ttk.Scrollbar(
            report_frame,
            orient="vertical",
            command=self.tree.yview
        )

        xscroll = ttk.Scrollbar(
            report_frame,
            orient="horizontal",
            command=self.tree.xview
        )

        self.tree.configure(
            yscrollcommand=yscroll.set,
            xscrollcommand=xscroll.set
        )

        self.tree.pack(
            fill="both",
            expand=True,
            side="left"
        )

        yscroll.pack(
            side="right",
            fill="y"
        )

        xscroll.pack(
            side="bottom",
            fill="x"
        )

        # --------------------------------------------------

        self.chart_container = ttk.LabelFrame(
            main_frame,
            text="Chart Preview"
        )

        self.chart_container.pack(
            side="right",
            fill="both",
            expand=True
        )

    # ======================================================
    # FILE OPERATIONS
    # ======================================================

    def browse_file(self):

        file_path = filedialog.askopenfilename(
            filetypes=[
                ("CSV Files", "*.csv"),
                ("Excel Files", "*.xlsx *.xls")
            ]
        )

        if file_path:
            self.file_path = file_path
            self.file_var.set(file_path)

    def read_file(self):

        try:

            if not self.file_path:
                raise Exception(
                    "Please select a file."
                )

            if self.file_path.lower().endswith(".csv"):
                self.df = pd.read_csv(
                    self.file_path
                )
            else:
                self.df = pd.read_excel(
                    self.file_path
                )

            self.detect_columns()

            self.show_dataset_info()

            messagebox.showinfo(
                "Success",
                "Dataset loaded successfully."
            )

        except Exception as e:
            messagebox.showerror(
                "Error",
                str(e)
            )

    # ======================================================
    # COLUMN DETECTION
    # ======================================================

    def detect_columns(self):

        self.text_columns = list(
            self.df.select_dtypes(
                include=["object"]
            ).columns
        )

        self.numeric_columns = []

        for col in self.df.columns:

            try:

                pd.to_numeric(
                    self.df[col]
                    .dropna()
                    .astype(str)
                    .str.replace(",", ""),
                    errors="raise"
                )

                self.numeric_columns.append(col)

            except:
                if pd.api.types.is_numeric_dtype(
                    self.df[col]
                ):
                    self.numeric_columns.append(
                        col
                    )

        self.group_combo["values"] = (
            self.text_columns
        )

        self.value_combo["values"] = (
            self.numeric_columns
        )

    # ======================================================
    # DATA CLEANING
    # ======================================================

    def clean_data(self):

        df = self.df.copy()

        df = df.drop_duplicates()

        df = df.dropna(how="all")

        text_cols = df.select_dtypes(
            include=["object"]
        ).columns

        for col in text_cols:

            df[col] = (
                df[col]
                .astype(str)
                .str.strip()
                .str.title()
            )

        return df

    # ======================================================
    # SHOW DATASET INFO
    # ======================================================

    def show_dataset_info(self):

        self.info_text.delete(
            "1.0",
            tk.END
        )

        rows, cols = self.df.shape

        info = (
            f"Rows : {rows}\n"
            f"Columns : {cols}\n\n"
            f"Column Names:\n"
        )

        for col in self.df.columns:
            info += f"• {col}\n"

        self.info_text.insert(
            tk.END,
            info
        )

    # ======================================================
    # REPORT BUILDER
    # ======================================================

    def preview_report(self):

        try:

            if self.df is None:
                raise Exception(
                    "Please click Read first."
                )

            group_col = self.group_combo.get()
            agg_method = self.agg_combo.get()
            value_col = self.value_combo.get()

            if not group_col:
                raise Exception(
                    "Please select Group By column."
                )

            if not agg_method:
                raise Exception(
                    "Please select Aggregation."
                )

            if not value_col:
                raise Exception(
                    "Please select Value column."
                )

            working_df = (
                self.clean_data()
                if self.clean_var.get()
                else self.df.copy()
            )

            agg_map = {
                "sum": "sum",
                "mean": "mean",
                "average": "mean",
                "max": "max",
                "min": "min",
                "count": "count",
                "median": "median"
            }

            self.report_df = (
                working_df
                .groupby(group_col)[value_col]
                .agg(agg_map[agg_method])
                .reset_index()
            )

            self.report_df = (
                self.report_df
                .sort_values(
                    by=value_col,
                    ascending=False
                )
            )

            self.display_report()

        except Exception as e:

            messagebox.showerror(
                "Error",
                str(e)
            )

    # ======================================================
    # REPORT DISPLAY
    # ======================================================

    def display_report(self):

        self.tree.delete(
            *self.tree.get_children()
        )

        self.tree["columns"] = list(
            self.report_df.columns
        )

        self.tree["show"] = "headings"

        for col in self.report_df.columns:

            self.tree.heading(
                col,
                text=col
            )

            self.tree.column(
                col,
                width=150,
                anchor="center"
            )

        for row in (
            self.report_df
            .values
            .tolist()
        ):

            self.tree.insert(
                "",
                tk.END,
                values=row
            )

    # ======================================================
    # EXPORT REPORT
    # ======================================================

    def export_report(self):

        try:

            if self.report_df is None:
                raise Exception(
                    "Generate report first."
                )

            folder = os.path.dirname(
                self.file_path
            )

            if (
                self.export_type.get()
                == "Excel"
            ):

                save_path = os.path.join(
                    folder,
                    "Analysis_Report.xlsx"
                )

                self.report_df.to_excel(
                    save_path,
                    index=False
                )

            else:

                save_path = os.path.join(
                    folder,
                    "Analysis_Report.csv"
                )

                self.report_df.to_csv(
                    save_path,
                    index=False
                )

            messagebox.showinfo(
                "Success",
                f"Report exported:\n{save_path}"
            )

        except Exception as e:

            messagebox.showerror(
                "Error",
                str(e)
            )

    # ======================================================
    # CHART BUILDER
    # ======================================================

    def preview_chart(self):

        try:

            if self.report_df is None:
                raise Exception(
                    "Generate report first."
                )

            chart_type = (
                self.chart_combo.get()
            )

            if not chart_type:
                raise Exception(
                    "Select chart type."
                )

            for widget in (
                self.chart_container
                .winfo_children()
            ):
                widget.destroy()

            x = self.report_df.iloc[:, 0]
            y = self.report_df.iloc[:, 1]

            fig, ax = plt.subplots(
                figsize=(12,8)
            )

            if chart_type == "Bar Chart":

                ax.barh(x, y)

            elif chart_type == "Column Chart":

                ax.bar(x, y)
                
                ax.tick_params(axis='x',labelrotation=90)

            elif chart_type == "Line Chart":

                ax.plot(
                    x,
                    y,
                    marker="o"
                )

            elif chart_type == "Pie Chart":

                ax.pie(
                    y,
                    labels=x,
                    autopct="%1.1f%%"
                )

            ax.set_title(
                chart_type
            )

            plt.xticks(rotation=0,
                       ha="right")

            self.chart_figure = fig

            canvas = FigureCanvasTkAgg(
                fig,
                master=self.chart_container
            )

            canvas.draw()

            canvas.get_tk_widget().pack(
                fill="both",
                expand=True
            )

        except Exception as e:

            messagebox.showerror(
                "Error",
                str(e)
            )

    # ======================================================
    # EXPORT CHART
    # ======================================================

    def export_chart(self):

        try:

            if self.chart_figure is None:
                raise Exception(
                    "Generate chart first."
                )

            folder = os.path.dirname(
                self.file_path
            )

            save_path = os.path.join(
                folder,
                "Analysis_Chart.png"
            )

            self.chart_figure.savefig(
                save_path,
                dpi=300,
                bbox_inches="tight"
            )

            messagebox.showinfo(
                "Success",
                f"Chart exported:\n{save_path}"
            )

        except Exception as e:

            messagebox.showerror(
                "Error",
                str(e)
            )


# ==========================================================
# APPLICATION START
# ==========================================================

if __name__ == "__main__":

    root = tk.Tk()

    style = ttk.Style()

    try:
        style.theme_use("clam")
    except:
        pass

    app = DataAnalysisApp(root)

    root.mainloop()




