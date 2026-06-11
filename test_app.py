"""
Automated smoke-test for GraphIt using DataVisualizer_TestData.xlsx.
Runs the app headlessly (no mainloop), exercises every core method,
and prints PASS/FAIL for each check.
"""

import os, sys, io
import tkinter as tk

# Fix Windows console encoding for emoji / unicode output
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

# Redirect stderr temporarily to suppress Tk warnings during headless test
os.environ.setdefault("MPLBACKEND", "TkAgg")

from data_visualizer import DataVisualizerApp
import pandas as pd

DATA_PATH = os.path.join(os.path.dirname(__file__), "DataVisualizer_TestData.xlsx")

results = []

def check(name, condition, detail=""):
    status = "PASS" if condition else "FAIL"
    results.append((name, status, detail))
    icon = "✅" if condition else "❌"
    print(f"  {icon} {status}: {name}" + (f"  ({detail})" if detail else ""))


def run_tests():
    print("=" * 65)
    print("  GraphIt — Automated Test Suite")
    print("  Dataset: DataVisualizer_TestData.xlsx")
    print("=" * 65)

    # --- 1. Create app ---
    print("\n[1] App Initialization")
    root = tk.Tk()
    root.withdraw()  # hide window
    app = DataVisualizerApp(root)
    check("App created", app is not None)
    check("Initial state: no data", app.df is None and app.filtered_df is None)

    # --- 2. Load Excel ---
    print("\n[2] File Loading")
    df = pd.read_excel(DATA_PATH, engine="openpyxl")
    app.df = df
    app.filtered_df = df.copy()
    app.current_file = DATA_PATH
    app._populate_table()
    app._update_info()
    app._populate_column_combos()
    root.update_idletasks()

    check("Data loaded", app.df is not None, f"{len(app.df)} rows × {len(app.df.columns)} cols")
    check("Row count correct", len(app.df) == 12, f"got {len(app.df)}")
    check("Column count", len(app.df.columns) == 8, f"{len(app.df.columns)} columns")

    # --- 3. Data Preview ---
    print("\n[3] Data Preview / Table")
    children = app.tree.get_children()
    check("Table populated (10 rows)", len(children) == 10, f"got {len(children)} rows")
    cols_in_tree = app.tree["columns"]
    check("Table columns match df", len(cols_in_tree) == len(app.df.columns))

    info_text = app.info_label.cget("text")
    check("Info label has row count", "12" in info_text, info_text[:80])
    check("Info label has col count", str(len(app.df.columns)) in info_text)

    # --- 4. Column combos populated ---
    print("\n[4] Dynamic Column Selection")
    x_vals = app.x_combo["values"]
    y_vals = app.y_combo["values"]
    check("X combo populated", len(x_vals) > 0, f"{len(x_vals)} entries")
    check("Y combo populated", len(y_vals) > 0)
    check("X default set", app.x_col_var.get() != "")
    check("Y default set", app.y_col_var.get() != "")

    # --- 5. Column search ---
    print("\n[5] Column Search")
    app.col_search_var.set("Pro")
    root.update_idletasks()
    search_result = app.col_search_result.cget("text")
    check("Search 'Pro' finds matches", "Product" in search_result, search_result[:80])

    app.col_search_var.set("xyznotacolumn")
    root.update_idletasks()
    search_result2 = app.col_search_result.cget("text")
    check("Search garbage → no match", "No matching" in search_result2)

    # --- 6. Filtering ---
    print("\n[6] Data Filtering")
    app.filter_col_var.set("Category")
    app.filter_val_var.set("Electronics")
    app._apply_filter()
    root.update_idletasks()
    n_filtered = len(app.filtered_df)
    check("Filter Category=Electronics", n_filtered > 0 and n_filtered < len(app.df),
          f"{n_filtered} rows")

    filter_info = app.filter_info_label.cget("text")
    check("Filter info label updated", str(n_filtered) in filter_info)

    # Clear filter
    app._clear_filter()
    root.update_idletasks()
    check("Clear filter restores all rows", len(app.filtered_df) == len(app.df))

    # Numeric filter
    app.filter_col_var.set("Sales")
    app.filter_val_var.set("50000")
    app._apply_filter()
    root.update_idletasks()
    n_sales = len(app.filtered_df)
    check("Numeric filter Sales=50000", n_sales > 0, f"{n_sales} rows")
    app._clear_filter()

    # --- 7. Plotting — all 6 chart types ---
    print("\n[7] Chart Plotting (6 types)")
    # Set numeric columns for plotting
    app.x_col_var.set("Product")
    app.y_col_var.set("Sales")

    for chart in DataVisualizerApp.CHART_TYPES:
        app.chart_type_var.set(chart)

        # Pie chart needs a categorical column
        if chart == "Pie Chart":
            app.x_col_var.set("Category")

        # Histogram and Box Plot need a numeric column on the X-axis
        if chart in ("Histogram", "Box Plot"):
            app.x_col_var.set("Sales")

        app._plot_chart()
        root.update_idletasks()

        axes = app.figure.axes
        has_plot = len(axes) > 0
        check(f"Plot: {chart}", has_plot)

        # Restore cols
        app.x_col_var.set("Product")
        app.y_col_var.set("Sales")

    # --- 8. Save / Export (dry-run) ---
    print("\n[8] Save & Export (method existence)")
    check("save_chart method exists", callable(getattr(app, "_save_chart", None)))
    check("export_csv method exists", callable(getattr(app, "_export_csv", None)))

    # Test actual CSV export to a temp file
    import tempfile
    tmp = os.path.join(os.path.dirname(__file__), "_test_export.csv")
    try:
        app.filtered_df.to_csv(tmp, index=False)
        reloaded = pd.read_csv(tmp)
        check("CSV export roundtrip", len(reloaded) == len(app.filtered_df),
              f"{len(reloaded)} rows written")
    finally:
        if os.path.exists(tmp):
            os.remove(tmp)

    # Test chart save
    tmp_img = os.path.join(os.path.dirname(__file__), "_test_chart.png")
    try:
        app.figure.savefig(tmp_img, dpi=72)
        check("Chart PNG save", os.path.exists(tmp_img),
              f"{os.path.getsize(tmp_img)} bytes")
    finally:
        if os.path.exists(tmp_img):
            os.remove(tmp_img)

    # --- 9. Error handling ---
    print("\n[9] Error Handling")
    # Non-numeric column for scatter
    app.chart_type_var.set("Scatter Plot")
    app.x_col_var.set("Product")  # string column
    app.y_col_var.set("Sales")
    # We test the validator directly
    from data_visualizer import DataVisualizerApp as DVA
    try:
        DVA._assert_numeric(app.df, "Product")
        check("Non-numeric assertion", False, "Should have raised ValueError")
    except ValueError:
        check("Non-numeric assertion raises ValueError", True)

    try:
        DVA._assert_numeric(app.df, "Sales")
        check("Numeric assertion passes for Sales", True)
    except ValueError:
        check("Numeric assertion passes for Sales", False)

    # --- 10. Dark mode toggle ---
    print("\n[10] Dark Mode Toggle")
    check("Starts in light mode", app.dark_mode is False)
    app._toggle_dark_mode()
    root.update_idletasks()
    check("Toggle → dark mode", app.dark_mode is True)
    check("Theme accent changed", app.theme["accent"] == "#7B68EE")
    app._toggle_dark_mode()
    root.update_idletasks()
    check("Toggle back → light mode", app.dark_mode is False)

    # --- 11. Reset ---
    print("\n[11] Reset")
    app._reset()
    root.update_idletasks()
    check("Reset clears df", app.df is None)
    check("Reset clears filtered_df", app.filtered_df is None)
    check("Reset clears table", len(app.tree.get_children()) == 0)
    check("Info label reset", "No dataset" in app.info_label.cget("text"))

    # --- Summary ---
    print("\n" + "=" * 65)
    passed = sum(1 for _, s, _ in results if s == "PASS")
    failed = sum(1 for _, s, _ in results if s == "FAIL")
    total = len(results)
    print(f"  Results: {passed}/{total} passed, {failed} failed")
    if failed == 0:
        print("  🎉 ALL TESTS PASSED!")
    else:
        print("  ⚠️  Failing tests:")
        for name, status, detail in results:
            if status == "FAIL":
                print(f"     ❌ {name}: {detail}")
    print("=" * 65)

    root.destroy()
    sys.exit(0 if failed == 0 else 1)


if __name__ == "__main__":
    run_tests()
