from flask import Flask, render_template, send_file
import pandas as pd
import numpy as np
# import matplotlib.pyplot as plt
import os

import matplotlib
matplotlib.use("Agg")   # Fix tkinter error
import matplotlib.pyplot as plt

app = Flask(__name__)

# Load data
df = pd.read_excel("sales.xlsx")

# covert date
df["Order_Date"] = pd.to_datetime(df["Order_Date"])

# create cart folder
if not os.path.exists("static/charts"):
    os.makedirs("static/charts")


@app.route("/")
def dashboard():
    total_revenue = np.sum(df["Total_Sales"])
    total_order = df["Order_ID"].count()

    # ===== Monthly Sales =====
    df["Month"] = df["Order_Date"].dt.month

    monthly_sales = df.groupby("Month")["Total_Sales"].sum()

    monthly_sales = monthly_sales.sort_index()   # important!

    plt.figure(figsize=(8,5))
    plt.plot(monthly_sales.index, monthly_sales.values, marker='o')
    plt.xticks(range(1,13))
    plt.title("Monthly Sales")
    plt.xlabel("Month")
    plt.ylabel("Revenue")
    plt.grid(True)

    plt.tight_layout()
    plt.savefig("static/charts/monthly_sales.png")
    plt.close()


    # product sales
    product_sales = df.groupby("Product")["Total_Sales"].sum()
    plt.figure()
    product_sales.plot(kind="bar")
    plt.title("Sales by product")
    plt.savefig("static/charts/product_sales.png")
    plt.close()

    # region sales
    region_sales = df.groupby("Region")["Total_Sales"].sum()
    plt.figure()
    region_sales.plot(kind="pie", autopct='%1.1f%%')
    plt.title("Sales by region")
    plt.ylabel("")
    plt.savefig("static/charts/region_sales.png")
    plt.close()

    return render_template("index.html", total_revenue=total_revenue, total_order=total_order)


@app.route("/download")
def download():
    report = df.groupby("Product")["Total_Sales"].sum().reset_index()
    report.to_csv("report.csv", index=False)
    return send_file("report.csv", as_attachment=True)


if __name__ == "__main__":
    app.run(debug=True)