from flask import Flask, jsonify, request
from flask_cors import CORS
import pandas as pd
import re
import pdb
import openpyxl
import numpy as np
from flask import send_from_directory
import os

app = Flask(__name__)
CORS(app)


@app.route('/static/<path:filename>')
def serve_static(filename):
    return send_from_directory(os.path.join('build', 'static'), filename)

@app.route('/', defaults={'path': ''})
@app.route('/<path:path>')
def serve_react(path):
    if path != "" and os.path.exists(os.path.join("build", path)):
        return send_from_directory("build", path)
    else:
        return send_from_directory("build", "index.html")

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

def generate_monthly_cashflows(df, start_date='2023-01-01', end_date='2031-01-01'):
    """
    Generates monthly loan-level cashflows including:
    - Cash interest
    - PIK interest (capitalised)
    - Unfunded fees
    - Updated principal balance each month
    """

    start_date = pd.to_datetime(start_date)
    end_date = pd.to_datetime(end_date)
    flows = []

    for _, row in df.iterrows():

        model_ref = row['Model_Ref']
        company = row['Company__Issuer_Name']
        instrument = row['Instrument_Name']
        ccy = row['Ccy']

        funded = float(row['Funded_EUR'])
        committed = float(row['Committed_EUR'])
        unfunded = committed - funded

        base_rate = float(row['Base_Rate'])
        cash_margin = float(row['Margin'])
        pik_margin = float(row.get('PIK', 0))

        cash_rate = base_rate + cash_margin
        pik_rate = pik_margin

        unfunded_fee_rate = float(row['Unfunded_Fee__of_margin']) * cash_margin

        day_basis = float(row['Day_Basis'])

        funding_date = pd.to_datetime(row['Initial_Funding_Date'])
        maturity_date = pd.to_datetime(row['Exit_Date'])

        monthly_dates = pd.date_range(
            start=funding_date.to_period("M").to_timestamp(),
            end=maturity_date.to_period("M").to_timestamp(),
            freq="MS"
        )

        principal = funded  # starting balance including future PIK

        for dt in monthly_dates:

            period_start = dt
            period_end = dt + pd.offsets.MonthEnd(1)
            days = (period_end - period_start).days

            # Beginning balance for the period
            balance_start = principal

            # Cash interest
            interest_cash = balance_start * cash_rate * days / day_basis

            # PIK interest (capitalised)
            interest_pik = balance_start * pik_rate * days / day_basis

            # Capitalise PIK at end of period
            principal = balance_start + interest_pik

            # Unfunded fees
            unfunded_fee = unfunded * unfunded_fee_rate * days / day_basis

            # Principal repayment only at maturity
            principal_repayment = principal if period_end >= maturity_date else 0

            # Ending balance after capitalisation (before repayment)
            balance_end = principal if period_end < maturity_date else 0

            flows.append({
                'Model_Ref': model_ref,
                'Date': period_end,
                'Company': company,
                'Instrument': instrument,
                'Currency': ccy,

                'Balance_Start_EUR': balance_start,
                'Interest_Cash_EUR': interest_cash,
                'Interest_PIK_EUR': interest_pik,
                'Unfunded_Fee_EUR': unfunded_fee,

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

import math

def portfolio_summary(agg_df, loan_df):
    summary = {}

    # --- CAPITAL CONTRIBUTED ---
    total_funded = loan_df['Funded_EUR'].sum()
    summary['Total_Capital_Funded'] = total_funded

    # --- TOTAL FLOWS ---
    summary['Start_Balance'] = agg_df['Balance_Start_EUR'].sum()
    summary['Total_Interest_Cash'] = agg_df['Interest_Cash_EUR'].sum()
    summary['Total_Interest_PIK'] = agg_df['Interest_PIK_EUR'].sum()
    summary['Total_Unfunded_Fees'] = agg_df['Unfunded_Fee_EUR'].sum()
    summary['Total_Principal_Returned'] = agg_df['Principal_EUR'].sum()
    summary['End_Balance'] = agg_df['Balance_End_EUR'].sum()

    # Net cashflow
    agg_df = agg_df.copy()
    agg_df['Net_Cashflow'] = (
        agg_df['Interest_Cash_EUR'] +
        agg_df['Interest_PIK_EUR'] +
        agg_df['Unfunded_Fee_EUR'] +
        agg_df['Principal_EUR']
    )

    summary['Total_Net_Cashflow'] = agg_df['Net_Cashflow'].sum()

    # --- IRR ---
    cashflows = agg_df.sort_values('Date')[['Date', 'Net_Cashflow', 'Principal_EUR']].copy()
    irr_series = [-total_funded] + list(cashflows['Net_Cashflow'])

    try:
        irr_monthly = float(np.irr(irr_series))
    except Exception:
        irr_monthly = None

    summary['IRR_Monthly'] = irr_monthly
    summary['IRR_Annualised'] = (
        (1 + irr_monthly)**12 - 1
        if irr_monthly is not None and not math.isnan(irr_monthly)
        else None
    )

    # --- MOIC ---
    summary['MOIC'] = (
        summary['Total_Net_Cashflow'] / total_funded
        if total_funded > 0 else None
    )

    # --- WAL ---
    cashflows['Months'] = (
        (cashflows['Date'].dt.year - cashflows['Date'].dt.year.min()) * 12 +
        cashflows['Date'].dt.month
    )

    total_principal = summary['Total_Principal_Returned']

    if total_principal > 0:
        WAL_months = (cashflows['Principal_EUR'] * cashflows['Months']).sum() / total_principal
        summary['WAL_Years'] = WAL_months / 12
    else:
        summary['WAL_Years'] = None

    # --- CLEAN NaN VALUES ---
    cleaned = {
    k: (None if (isinstance(v, float) and math.isnan(v)) else v)
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