from flask import Flask, jsonify, request
from flask_cors import CORS
import pandas as pd
import re
import pdb
import openpyxl
import numpy as np
from flask import send_from_directory
import os
import math

app = Flask(__name__, static_folder='build/static', static_url_path='/static')
CORS(app)
# Serve static files from React build
@app.route('/static/<path:path>')
def serve_static(path):
    return send_from_directory(os.path.join('build', 'static'), path)
# Serve React App
@app.route('/', defaults={'path': ''})
@app.route('/<path:path>')
def serve_react(path):
    full_path = os.path.join('build', path)
    if path != "" and os.path.exists(full_path):
        return send_from_directory('build', path)
    return send_from_directory('build', 'index.html')

@app.route('/api/add', methods=['POST'])
def add_numbers():
    data = request.get_json()
    a = data.get('a')
    b = data.get('b')
    if a is None or b is None:
        return jsonify({"error": "Missing parameters 'a' and 'b'"}), 400
    try:
        result = float(a) + float(b)
    except ValueError:
        return jsonify({"error": "Invalid input"}), 400
    return jsonify({"result": result})

def load_assets_excel(file_path='assets.xlsx'):
    df = pd.read_excel(file_path)
    # Clean column names: remove spaces and symbols, keep letters, numbers, and underscores
    df.columns = [
        re.sub(r'\W+', '', col.replace(' ', '_')) for col in df.columns
    ]
    return df


def load_fund_excel(file_path='fund_terms.xlsx'):
    df = pd.read_excel(file_path)
    # Clean column names: remove spaces and symbols, keep letters, numbers, and underscores
    df.columns = [
        re.sub(r'\W+', '', col.replace(' ', '_')) for col in df.columns
    ]
    return df

@app.route('/api/read_assets', methods=['GET'])
def read_assets_excel(file_path='assets.xlsx'):
    try:
        df = load_assets_excel()
        data = df.to_dict(orient='records')
        columns = list(df.columns)
        return jsonify({"data": data, "columns": columns})
    except Exception as e:
        print(f"Error reading {file_path}: {e}")
        return jsonify({"error": str(e)}), 500

def generate_monthly_cashflows(df):
    """
    Monthly loan-level cashflows:
    - Cash interest
    - PIK interest (capitalised monthly)
    - Unfunded fees
    - Updated balances
    - Guaranteed principal repayment at exit
    """

    flows = []

    for _, row in df.iterrows():

        # Extract core loan fields
        model_ref = row['Model_Ref']
        funded = float(row['Funded_EUR'])
        committed = float(row['Committed_EUR'])
        unfunded = committed - funded

        base_rate = float(row['Base_Rate'])
        cash_margin = float(row['Margin'])
        pik_margin = float(row.get('PIK', 0))

        # Rates
        cash_rate = base_rate + cash_margin
        pik_rate = pik_margin

        # Unfunded fee (corrected double-underscore version)
        unfunded_fee_rate = float(row['Unfunded_Fee__of_margin']) * cash_margin

        day_basis = float(row['Day_Basis'])

        company = row['Company__Issuer_Name']
        instrument = row['Instrument_Name']
        ccy = row['Ccy']

        # Dates
        start_date = pd.to_datetime(row['Initial_Funding_Date'])
        exit_date = pd.to_datetime(row['Exit_Date'])    # correct

        # Generate month ends from funding month → exit month
        month_starts = pd.date_range(
            start=start_date.to_period("M").to_timestamp(),
            end=exit_date.to_period("M").to_timestamp(),
            freq="MS"
        )

        principal = funded  # starting balance including PIK

        for dt in month_starts:

            period_start = dt
            period_end = dt + pd.offsets.MonthEnd(1)
            period_end = min(period_end, exit_date)  # clamp to exit

            days = (period_end - period_start).days

            balance_start = principal

            # Cash interest
            interest_cash = balance_start * cash_rate * days / day_basis

            # PIK interest capitalised monthly
            interest_pik = balance_start * pik_rate * days / day_basis
            principal = balance_start + interest_pik

            # Unfunded fee
            unf_fee = unfunded * unfunded_fee_rate * days / day_basis

            # Principal repayment (forced on exit period)
            is_exit_period = (dt == month_starts[-1])
            principal_repayment = principal if is_exit_period else 0
            balance_end = 0 if is_exit_period else principal

            flows.append({
                'Model_Ref': model_ref,
                'Date': period_end,
                'Company': company,
                'Instrument': instrument,
                'Currency': ccy,

                'Balance_Start_EUR': balance_start,
                'Interest_Cash_EUR': interest_cash,
                'Interest_PIK_EUR': interest_pik,
                'Unfunded_Fee_EUR': unf_fee,

                'Principal_EUR': principal_repayment,
                'Balance_End_EUR': balance_end
            })

    return pd.DataFrame(flows)

def aggregate_cashflows(cashflow_df):
    """
    Aggregates all cashflows per month.
    """
    grouped = (
        cashflow_df
        .groupby('Date')[['Balance_Start_EUR','Interest_Cash_EUR', 'Interest_PIK_EUR', 'Unfunded_Fee_EUR', 'Principal_EUR','Balance_End_EUR']]
        .sum()
        .reset_index()
    )
    return grouped

def aggregate_cashflows_for_react(cashflow_df):
    """
    Returns a pivoted table with:
    - Dates as columns
    - Metrics as rows
    For easy consumption by React frontend.
    """
    # First aggregate normal cashflows per month
    grouped = (
        cashflow_df
        .groupby('Date')[[
            'Balance_Start_EUR',
            'Interest_Cash_EUR',
            'Interest_PIK_EUR',
            'Unfunded_Fee_EUR',
            'Principal_EUR',
            'Balance_End_EUR'
        ]]
        .sum()
        .reset_index()
    )
    # Pivot into React-friendly shape
    pivot_df = grouped.set_index('Date').T
    # Optional: reset index to get "metric" column for React
    pivot_df = pivot_df.reset_index().rename(columns={'index': 'Metric'})
    return pivot_df

@app.route('/api/cfs', methods=['GET'])
def get_initial_cashflows():
    df = load_assets_excel()
    flows = generate_monthly_cashflows(df)
    pivot = aggregate_cashflows_for_react(flows)
    return jsonify(to_highcharts_series(pivot))

def to_highcharts_series(pivot_df):
    """
    Converts the pivoted table (Metric × Date columns) into
    Highcharts-ready JSON structure:

    {
        "categories": [...dates...],
        "series": [
            { "name": "Balance_Start_EUR", "data": [...] },
            { "name": "Interest_Cash_EUR", "data": [...] },
            ...
        ]
    }
    """
    # categories = all date columns (skip the Metric column)
    #categories = list(pivot_df.columns[1:])
    categories = [
        pd.to_datetime(col).strftime("%Y-%m-%d") 
        for col in pivot_df.columns[1:]
    ]
    # metrics = every row
    series = []
    for _, row in pivot_df.iterrows():
        metric = row['Metric']

        # Extract numeric values only for date columns
        values = [row[col] for col in pivot_df.columns[1:]]

        series.append({
            "name": metric,
            "data": values
        })

    return {
        "categories": categories,
        "series": series
    }

@app.route("/api/recompute-cashflows", methods=["POST"])
def recompute_cashflows():
    data = request.get_json()
    margin_shock_bps = float(data.get("margin_shock_bps", 0))
    exit_shock_months = int(data.get("exit_shock_months", 0))
    # Load base data
    df = load_assets_excel("assets.xlsx").copy()
    # Margin shock (bps → decimal)
    margin_shock = margin_shock_bps / 10000.0
    df["Margin"] = df["Margin"] + margin_shock
    df["Total_Allin_Margin_Margin"] = (
        df["Margin"] + df["Base_Rate"] + df["PIK"]
    )
    # Exit shock
    df["Exit_Date"] = pd.to_datetime(df["Exit_Date"]) + pd.DateOffset(
        months=exit_shock_months
    )
    # Recompute flows
    flows = generate_monthly_cashflows(df)
    # Pivot cashflows
    pivot = aggregate_cashflows_for_react(flows)
    # Format for React table
    hc_data = to_highcharts_series(pivot)
    return jsonify({
        "cashflows": hc_data,
        "assets": df.to_dict(orient="records")
    })



def compute_irr(agg_df, total_funded):
    """
    Computes Monthly IRR and Annualised IRR safely:
    - Guarantees sign change
    - Avoids NaN or None outputs
    """
    # Monthly net cashflows (interest + fees + principal)
    flows = agg_df.sort_values("Date")[["Net_Cashflow"]].copy()
    # Initial outflow
    irr_series = [-total_funded] + list(flows["Net_Cashflow"].values)
    # If no positive inflows exist, IRR is undefined
    if max(irr_series) <= 0:
        return None, None
    # Compute IRR
    try:
        irr_monthly = float(np.irr(irr_series))
    except Exception:
        return None, None
    if irr_monthly is None or math.isnan(irr_monthly):
        return None, None
    irr_annual = (1 + irr_monthly)**12 - 1
    return irr_monthly, irr_annual

def portfolio_summary(agg_df, loan_df):
    summary = {}
    total_funded = loan_df["Funded_EUR"].sum()
    summary["Total_Capital_Funded"] = total_funded

    summary["Start_Balance"] = agg_df["Balance_Start_EUR"].sum()
    summary["Total_Interest_Cash"] = agg_df["Interest_Cash_EUR"].sum()
    summary["Total_Interest_PIK"] = agg_df["Interest_PIK_EUR"].sum()
    summary["Total_Unfunded_Fees"] = agg_df["Unfunded_Fee_EUR"].sum()
    summary["Total_Principal_Returned"] = agg_df["Principal_EUR"].sum()
    summary["End_Balance"] = agg_df["Balance_End_EUR"].sum()

    agg_df = agg_df.copy()
    agg_df["Net_Cashflow"] = (
        agg_df["Interest_Cash_EUR"]
        + agg_df["Interest_PIK_EUR"]
        + agg_df["Unfunded_Fee_EUR"]
        + agg_df["Principal_EUR"]
    )

    summary["Total_Net_Cashflow"] = agg_df["Net_Cashflow"].sum()

    # Compute IRR safely
    irr_monthly, irr_annual = compute_irr(agg_df, total_funded)
    summary["IRR_Monthly"] = irr_monthly
    summary["IRR_Annualised"] = irr_annual

    # WAL
    cashflows = agg_df.sort_values("Date")
    total_principal = summary["Total_Principal_Returned"]

    if total_principal > 0:
        first_year = cashflows["Date"].dt.year.min()
        months_offset = (
            (cashflows["Date"].dt.year - first_year) * 12 +
            cashflows["Date"].dt.month
        )
        WAL_months = (cashflows["Principal_EUR"] * months_offset).sum() / total_principal
        summary["WAL_Years"] = WAL_months / 12
    else:
        summary["WAL_Years"] = None

    # Clean NaN for JSON
    cleaned = {
        k: (None if isinstance(v, float) and math.isnan(v) else v)
        for k, v in summary.items()
    }

    return cleaned


@app.route('/api/portfolio_summary', methods=['GET'])
def get_portfolio_summary():
    df = load_assets_excel()
    flows = generate_monthly_cashflows(df)
    agg = aggregate_cashflows(flows)
    summary = portfolio_summary(agg, df)
    return jsonify(summary)

@app.route("/debug")
def debug():
    import os
    return {
        "cwd": os.getcwd(),
        "files": os.listdir(os.getcwd()),
        "parent_files": os.listdir(os.path.dirname(os.getcwd())),
        "build_exists_here": os.path.exists("build"),
        "build_exists_parent": os.path.exists("../build"),
        "build_static_exists": os.path.exists("build/static"),
    }

if __name__ == '__main__':
    
    app.run(debug=True)