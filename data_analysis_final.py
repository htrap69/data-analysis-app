from tkinter import *
from tkcalendar import Calendar
import pandas as pd
import numpy as np
from mlxtend.frequent_patterns import apriori
from mlxtend.frequent_patterns import association_rules
from tkinter import ttk
from threading import Thread
import os
import subprocess
from tkinter import messagebox


def threading(func):
    t = Thread(target=func)
    t.start()


def show_error(message):
    messagebox.showerror("Error", message)


def progress_bar_update(msg, progress=0):
    progress_bar["value"] = progress
    message.set(msg)
    window.update_idletasks()


def get_data_from_db(start_date, end_date):
    try:
        conn = pyodbc.connect(
            DRIVER="{ODBC Driver}",
            SERVER="SERVER_ADDRESS",
            DATABASE="DATABASE_NAME",
            UID="USERNAME",
            PWD="PASSWORD",
        )
        progress_bar_update('Connected to Database', 10)
    except Exception as e:
        show_error(f"Connection Error: {str(e)}")

    query = f"SELECT * FROM TABLE_NAME WHERE [Invoice Date] >= '{start_date}' AND [Invoice Date] <= '{end_date}'"
    progress_bar_update('Collecting Data', 15)
    data = pd.read_sql(query, conn)
    conn.close()
    progress_bar_update('Data Collected', 20)
    return data


def encode_units(x):
    if x <= 0:
        return 0
    if x >= 1:
        return 1


def run_apriori_analysis():
    try:
        progress_bar_update('Starting Apriori Analysis', 0)
        level = level_var.get()
        start_date = start_date_entry.get()
        end_date = end_date_entry.get()
        min_support = min_support_entry.get()
        min_threshold = min_threshold_entry.get()
        df = get_data_from_db(start_date, end_date)
        # df['Invoice Date'] = pd.to_datetime(df['Invoice Date'])
        if level == "SKU":
            pivot_to = "Product Code"
        elif level == "Category":
            pivot_to = "CATEGORY"
        elif level == "SubCategory":
            pivot_to = "Product-CategoryUDF"
        df = df.pivot_table(
            index="Invoice#",
            columns=pivot_to,
            values="OrderedQty",
            aggfunc=np.sum,
            fill_value=0,
        )
        progress_bar_update('Organizing Data', 50)
        basket_sets = df.applymap(encode_units)
        progress_bar_update('Running Apriori Analysis', 60)
        frequent_itemsets = apriori(
            basket_sets, min_support=float(min_support), use_colnames=True
        )
        rules = association_rules(
            frequent_itemsets, metric="lift", min_threshold=float(min_threshold)
        )
        progress_bar_update('Saving Output', 80)
        rules.to_csv("Rules.csv")
        df.to_csv("Data.csv")
        progress_bar_update('Analysis Completed', 100)
        print(df)
        print(frequent_itemsets)
        print(rules)
        print("Done")
    except Exception as e:
        show_error(f"An error occured: {str(e)}")


def run_basket_pen_analysis():
    try:
        progress_bar_update('Starting Basket Analysis', 0)
        level = level_var.get()
        start_date = pd.to_datetime(start_date_entry.get(), dayfirst=True)
        end_date = pd.to_datetime(end_date_entry.get(), dayfirst=True)
        df = get_data_from_db(start_date, end_date)
        progress_bar_update('Organizing Data', 50)
        df["Invoice Date"] = pd.to_datetime(df["Invoice Date"], dayfirst=True)
        if level == "SKU":
            pivot_to = "Product Code"
        elif level == "Category":
            pivot_to = "CATEGORY"
        elif level == "SubCategory":
            pivot_to = "Product-CategoryUDF"
        df = df.pivot_table(
            index="Invoice#",
            columns=pivot_to,
            values="OrderedQty",
            aggfunc=np.sum,
            fill_value=0,
        )
        progress_bar_update('Running Basket Pen Analysis', 60)
        basket_sets = df.applymap(encode_units)
        basket_pen = basket_sets
        basket_pen.loc["Total"] = basket_pen.sum()
        basket_pen = basket_pen.loc["Total"]
        progress_bar_update('Saving Output', 80)
        basket_pen.to_csv(f"{level} - Basket_Penetration.csv")
        item_qty = basket_sets
        item_qty["Total"] = item_qty.sum(axis=1)
        item_qty = item_qty.groupby(["Total"]).size()
        item_qty.to_csv(f"{level} - cart_size.csv")
        item_1qty = item_qty[item_qty["Total"] == 1]
        item_1qty["Total"] = item_1qty.sum()
        item_1qty.to_csv(f"{level} - single_purchases.csv")
        progress_bar_update('Analysis Completed', 100)
        print("Analysis Complete")
    except Exception as e:
        show_error(f"An error occured: {str(e)}")


def open_output_file():
    file_path = "Output.xlsx"
    if os.path.exists(file_path):
        if os.name == 'nt':  # For Windows
            os.startfile(file_path)
        else:  # For other OS (Mac, Linux)
            subprocess.run(['open', file_path])
    else:
        print("File not found.")


if __name__ == '__main__':
    window = Tk()
    frame = Frame(window)
    window.geometry('800x600')

    heading_label = Label(frame, text="Analysis Tool",
                          font=("SegoeUI", 24, "bold"))
    level_label = Label(frame, text="Select the level of analysis:")
    level_var = StringVar(frame)
    level_var.set("Select")
    level_dropdown = OptionMenu(
        frame, level_var, "Select", "SKU", "Category", "SubCategory")

    start_date_entry = Entry(frame)
    end_date_entry = Entry(frame)
    date_label1 = Label(frame, text="Enter the Start date:")
    start_date_button = Button(
        frame, text="Select", command=lambda: choose_date(start_date_entry))
    date_label2 = Label(frame, text="Enter the End date:")
    end_date_button = Button(frame, text="Select",
                             command=lambda: choose_date(end_date_entry))

    min_support_label = Label(frame, text="Enter the minimum support:")
    min_support_entry = Entry(frame)
    min_threshold_label = Label(frame, text="Enter the minimum threshold:")
    min_threshold_entry = Entry(frame)

    apriori_button = Button(frame, text="Run Apriori Analysis",
                            command=lambda: threading(run_apriori_analysis))
    basket_pen_button = Button(frame, text="Run Basket Analysis",
                               command=lambda: threading(run_basket_pen_analysis))

    message = StringVar()
    progress_label = Label(frame, textvariable=message)
    progress_bar = ttk.Progressbar(frame, length=200, mode="determinate")

    heading_label.grid(column=0, row=0, columnspan=4, sticky='ew', pady=10)
    level_label.grid(column=0, row=1, sticky='w')
    level_dropdown.grid(column
