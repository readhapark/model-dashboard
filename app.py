import math
from flask import Flask, jsonify, request, send_from_directory
from flask_cors import CORS
import pandas as pd
import numpy as np
import re
import os
from pyxirr import xirr

app = Flask(__name__, static_folder="build/static", static_url_path="/static")
CORS(app)

###############################################################################
# STATIC ROUTES
###############################################################################

@app.route('/static/<path:path>')
def serve_static(path):
    return send_from_directory(os.path.join("build", "static"), path)

@app.route('/', defaults={'path': ''})
@app.route('/<path:path>')
def serve_react(path):
    fp = os.path.join("build", path)
    if path != "" and os.path.exists(fp):
        return send_from_directory("build", path)
    return send_from_directory("build", "index.html")

###############################################################################
# LOAD ASSETS
###############################################################################

def load_assets_excel(file_path="assets.xlsx"):
    df = pd.read_excel(file_path)

    df.columns = (
        df.columns
        .str.strip()
        .str.replace(r"\s+", "_", regex=True)
        .str.replace(r"[^\w_]", "", regex=True)
    )

    df["Initial_Funding_Date"] = pd.to_datetime(df["Initial_Funding_Date"], dayfirst=True, errors="coerce")
    df["Exit_Date"] = pd.to_datetime(df["Exit_Date"], dayfirst=True, errors="coerce")

    numeric_cols = [
        "Funded_EUR","Committed_EUR","Margin","Base_Rate","PIK",
        "Unfunded_Fee__of_margin","Day_Basis"
    ]
    for c in numeric_cols:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)

    return df

###############################################################################
# /api/read_assets
###############################################################################

@app.route('/api/read_assets', methods=['GET'])
def api_read_assets():
    df = load_assets_excel()
    return jsonify({
        "data": df.to_dict("records"),
        "columns": list(df.columns)
    })

###############################################################################
# MONTHLY CASHFLOWS â€” bullet repayment once at exit
###############################################################################

def generate_monthly_cashflows(df, exit_multiple=1.0):
    flows = []

    for row in df.itertuples(index=False):

        funded    = row.Funded_EUR
        committed = row.Committed_EUR
        unfunded  = committed - funded

        cash_rate = row.Base_Rate + row.Margin
        pik_rate  = row.PIK
        uf_rate   = row.Unfunded_Fee__of_margin * row.Margin
        day_basis = row.Day_Basis

        start_date = row.Initial_Funding_Date
        exit_date  = row.Exit_Date

        start_month = start_date.replace(day=1)
        ms_end      = exit_date.replace(day=1)

        month_starts = pd.date_range(start=start_month, end=ms_end, freq="MS")

        # If exit date is midâ€‘month, include it
        if exit_date != ms_end:
            month_starts = month_starts.append(pd.DatetimeIndex([exit_date]))

        principal = funded

        for dt in month_starts:

            # Exit row â€” repay once and stop
            if dt == exit_date:
                flows.append({
                    "Date": exit_date,
                    "Balance_Start_EUR": principal,
                    "Interest_Cash_EUR": 0,
                    "Interest_PIK_EUR": 0,
                    "Unfunded_Fee_EUR": 0,
                    "Principal_EUR": principal * exit_multiple,
                    "Balance_End_EUR": 0
                })
                principal = 0
                break

            period_start = dt
            period_end = dt + pd.offsets.MonthEnd(1)
            if period_end > exit_date:
                period_end = exit_date

            days = (period_end - period_start).days
            bal_start = principal

            int_cash = bal_start * cash_rate * days / day_basis
            int_pik  = bal_start * pik_rate  * days / day_basis
            bal_after_pik = bal_start + int_pik

            unf_fee = unfunded * uf_rate * days / day_basis

            if period_end == exit_date:
                prin_repay = bal_after_pik * exit_multiple
                bal_end = 0
                principal = 0
            else:
                prin_repay = 0
                bal_end = bal_after_pik
                principal = bal_end

            flows.append({
                "Date": period_end,
                "Balance_Start_EUR": bal_start,
                "Interest_Cash_EUR": int_cash,
                "Interest_PIK_EUR": int_pik,
                "Unfunded_Fee_EUR": unf_fee,
                "Principal_EUR": prin_repay,
                "Balance_End_EUR": bal_end
            })

    return pd.DataFrame(flows)

###############################################################################
# CFS SUPPORT
###############################################################################

def aggregate_for_cfs(flows):
    return (
        flows.groupby("Date")[[
            "Balance_Start_EUR","Interest_Cash_EUR",
            "Interest_PIK_EUR","Unfunded_Fee_EUR",
            "Principal_EUR","Balance_End_EUR"
        ]].sum().reset_index()
    )

def pivot_for_highcharts(df):
    dfp = df.set_index("Date").T
    return dfp.reset_index().rename(columns={"index":"Metric"})

def to_highcharts(df):
    categories = [pd.to_datetime(c).strftime("%Y-%m-%d") for c in df.columns[1:]]
    series = [{"name": r["Metric"], "data":[r[c] for c in df.columns[1:]]}
              for _, r in df.iterrows()]
    return {"categories":categories,"series":series}

###############################################################################
# /api/cfs
###############################################################################

@app.route('/api/cfs', methods=['GET'])
def api_cfs():
    df = load_assets_excel()
    flows = generate_monthly_cashflows(df)
    grouped = aggregate_for_cfs(flows)
    pivot = pivot_for_highcharts(grouped)
    return jsonify(to_highcharts(pivot))

###############################################################################
# PORTFOLIO IRR (pyxirr)
###############################################################################

def build_portfolio_cashflows(flows, loans_df):
    """
    Combines:
      - monthly interest, PIK, fees, principal
      - initial capital calls (negative)
      - produces:
          Gross_Income  = inflows including principal
          Net_Cashflow  = Gross_Income (same definition)
    """

    f = flows.copy()

    # Gross Income now includes principal repayments
    f["Gross_Income"] = (
        f["Interest_Cash_EUR"] +
        f["Interest_PIK_EUR"] +
        f["Unfunded_Fee_EUR"] +
        f["Principal_EUR"]
    )

    # Net Cashflow = Gross Income (same meaning for portfolio waterfall)
    f["Net_Cashflow"] = f["Gross_Income"]

    # Capital calls are still negative
    calls = loans_df[["Initial_Funding_Date", "Funded_EUR"]].copy()
    calls = calls.rename(columns={"Initial_Funding_Date": "Date"})
    calls["Gross_Income"] = 0
    calls["Net_Cashflow"] = -calls["Funded_EUR"]

    combined = pd.concat([
        f[["Date", "Gross_Income", "Net_Cashflow"]],
        calls[["Date", "Gross_Income", "Net_Cashflow"]]
    ])
    combined = to_month_end(combined, "Date")
    return (
        combined.groupby("Date")
        [["Gross_Income", "Net_Cashflow"]]
        .sum()
        .reset_index()
        .sort_values("Date")
    )

def compute_xirr(port_cf, total_funded):
    print("REEADHA:total_funded:", total_funded)
    print("REEADHA:port_cf head:\n", port_cf.head())
    dates = port_cf["Date"].tolist()
    amounts = [-total_funded] + port_cf["Gross_Income"].tolist()

    cf_map = dict(zip(dates, amounts))

    try:
        irr_annual = xirr(cf_map)
    except:
        return None, None

    irr_monthly = (1 + irr_annual)**(1/12) - 1
    return irr_monthly, irr_annual

###############################################################################
# PORTFOLIO SUMMARY (includes WAL)
###############################################################################

def portfolio_summary(flows, loans, warehouse_advance):
    total_funded = loans["Funded_EUR"].sum()
    port_cf = build_portfolio_cashflows(flows, loans)
    wf = run_waterfall(port_cf, warehouse_advance)
    lp_performance=compute_lp_performance(wf)
    irr_m, irr_y = compute_xirr(port_cf, total_funded)

    total_cash = flows["Interest_Cash_EUR"].sum()
    total_pik  = flows["Interest_PIK_EUR"].sum()
    total_fee  = flows["Unfunded_Fee_EUR"].sum()
    total_prin = flows["Principal_EUR"].sum()

    moic = (total_cash + total_pik + total_fee + total_prin) / total_funded if total_funded > 0 else None

    # WAL
    total_principal = total_prin
    flows_sorted = flows.sort_values("Date")

    if total_principal > 0:
        first_year = flows_sorted["Date"].dt.year.min()
        months_offset = (
            (flows_sorted["Date"].dt.year - first_year) * 12 +
            flows_sorted["Date"].dt.month
        )
        WAL_months = (flows_sorted["Principal_EUR"] * months_offset).sum() / total_principal
        WAL_years = WAL_months / 12
    else:
        WAL_years = None

    return {
        "Total_Capital_Funded": total_funded,
        "Total_Interest_Cash": total_cash,
        "Total_Interest_PIK": total_pik,
        "Total_Unfunded_Fees": total_fee,
        "Total_Principal_Returned": total_prin,
        "MOIC": moic,
        "IRR_Monthly": irr_m,
        "IRR_Annualised": irr_y,
        "WAL_Years": WAL_years,
        "LP_Net_IRR": lp_performance.get("LP_Net_IRR", 0.0),
        "LP_Net_MOIC": lp_performance.get("LP_Net_MOIC", 0.0)
    }

###############################################################################
# /api/portfolio_summary
###############################################################################

@app.route('/api/portfolio_summary', methods=['GET'])
def api_summary():
    loans = load_assets_excel()
    flows = generate_monthly_cashflows(loans)
    return jsonify(portfolio_summary(flows, loans, warehouse_advance=0.0))

###############################################################################
# /api/recompute-cashflows
###############################################################################

@app.route("/api/recompute-cashflows", methods=["POST"])
def api_recompute():
    data = request.get_json()

    margin_shock  = float(data.get("margin_shock_bps", 0))
    exit_shock    = int(data.get("exit_shock_months", 0))
    exit_multiple = float(data.get("exit_multiple", 1.0))
    warehouse_advance = float(data.get("warehouse_advance", 0.0))

    df = load_assets_excel().copy()

    df["Margin"] += margin_shock / 10000.0
    df["Exit_Date"] += pd.DateOffset(months=exit_shock)

    flows = generate_monthly_cashflows(df, exit_multiple)
    summary = portfolio_summary(flows, df, warehouse_advance)

    grouped = aggregate_for_cfs(flows)
    pivot = pivot_for_highcharts(grouped)
    # ----------------------------------------------------------------------
    # NEW: Build portfolio cashflow (combined inflows/outflows)
    # ----------------------------------------------------------------------
    port_cf = build_portfolio_cashflows(flows, df)
    port_cf = to_month_end(port_cf)
    port_cf = port_cf.rename(columns={"Date": "Date", "Net_Cashflow": "Net_Cashflow"})
    # Waterfall
    wf = run_waterfall(port_cf, warehouse_advance)

    lp_perf = compute_lp_performance(wf)
    wf_chart = waterfall_to_highcharts(wf)
    # Merge into summary object
    summary.update(lp_perf)
    return jsonify({
        "cashflows": to_highcharts(pivot),
        "assets": df.to_dict("records"),
        "summary": summary,
        "waterfall": wf_chart
    })

###############################################################################
# PRIVATE CREDIT WATERFALL (GP/LP)
###############################################################################
def to_month_end(df, date_col="Date"):
    df = df.copy()
    df[date_col] = pd.to_datetime(df[date_col]) + pd.offsets.MonthEnd(0)
    return df

def run_waterfall(
    cashflows_df,
    warehouse_advance_rate,
    mgmt_fee_rate=0.01,
    hurdle_rate=0.05,
    carry_rate_lp=0.85,
    carry_rate_gp=0.15,
    sub_line_rate=0.05
):
    df = cashflows_df.copy().sort_values("Date")

    # State
    lp_contrib_total = 0.0
    lp_unreturned_capital = 0.0
    lp_accrued_pref = 0.0
    warehouse_balance = 0.0
    nav = 0.0

    results = []

    for _, row in df.iterrows():
        dt = row["Date"]
        gross_income = float(row.get("Gross_Income", 0.0))
        inflow = float(row.get("Net_Cashflow", gross_income))

        # 1) CAPITAL CALLS (negative inflow)
        if inflow < 0:
            call = -inflow
            warehouse_draw = warehouse_advance_rate * call
            lp_draw = (1 - warehouse_advance_rate) * call

            warehouse_balance += warehouse_draw
            lp_contrib_total += lp_draw
            lp_unreturned_capital += lp_draw
            nav += call  # NAV increases by capital deployed

            results.append({
                "Date": dt,
                "NAV": nav,
                "Gross_Income": 0.0,
                "Warehouse_Paid": 0.0,
                "LP_Dist": 0.0,
                "GP_Dist": 0.0,
                "Warehouse_Balance": warehouse_balance,
                "LP_Balance": lp_unreturned_capital,
                "LP_Contrib": lp_draw
            })
            continue

        # 2) DISTRIBUTION MONTH
        cash = gross_income

        # A. Warehouse interest (expense)
        wh_interest = warehouse_balance * sub_line_rate / 12.0
        wh_int_paid = min(cash, wh_interest)
        cash -= wh_int_paid

        # B. Management fee (to GP)
        mgmt_fee = nav * mgmt_fee_rate / 12.0
        mgmt_paid = min(cash, mgmt_fee)
        cash -= mgmt_paid

        # C. Return LP capital
        lp_cap_return = min(cash, lp_unreturned_capital)
        lp_unreturned_capital -= lp_cap_return
        lp_contrib_total = max(0.0, lp_contrib_total - lp_cap_return)
        cash -= lp_cap_return

        # D. Accrue LP pref if capital outstanding
        if lp_unreturned_capital > 0:
            lp_accrued_pref += lp_unreturned_capital * hurdle_rate / 12.0

        # E. Pay LP pref once capital is fully returned
        lp_pref_paid = 0.0
        if lp_unreturned_capital == 0.0 and lp_accrued_pref > 0.0:
            lp_pref_paid = min(cash, lp_accrued_pref)
            lp_accrued_pref -= lp_pref_paid
            cash -= lp_pref_paid

        # F. Pay warehouse principal after LP capital + pref are cleared
        wh_prin_paid = 0.0
        if lp_unreturned_capital == 0.0 and lp_accrued_pref == 0.0:
            wh_prin_paid = min(cash, warehouse_balance)
            warehouse_balance -= wh_prin_paid
            cash -= wh_prin_paid

        # G. Residual split 85/15 (no GP 100% catch-up here)
        gp_share = 0.0
        lp_share = 0.0
        if lp_unreturned_capital == 0.0 and lp_accrued_pref == 0.0 and warehouse_balance == 0.0 and cash > 0.0:
            gp_share = cash * carry_rate_gp
            lp_share = cash * carry_rate_lp
            cash = 0.0

        # Assemble reported buckets
        warehouse_paid = wh_int_paid + wh_prin_paid
        lp_dist = lp_cap_return + lp_pref_paid + lp_share
        gp_dist = mgmt_paid + gp_share
        # Note: warehouse_paid is excluded from gp_dist to avoid overstating distributions.

        # Optional sanity check
        # total_out = warehouse_paid + lp_dist + gp_dist
        # if abs(total_out - gross_income) > 1e-6:
        #     print("WARN: payouts != income", dt, total_out, gross_income)

        results.append({
            "Date": dt,
            "NAV": nav,
            "Gross_Income": gross_income,
            "Warehouse_Paid": warehouse_paid,
            "LP_Dist": lp_dist,
            "GP_Dist": gp_dist,
            "Warehouse_Balance": warehouse_balance,
            "LP_Balance": lp_unreturned_capital,
            "LP_Contrib": 0.0
        })

    # 3) FINAL CLEANUP
    if results:
        last = results[-1]
        if last["LP_Balance"] > 0:
            last["LP_Dist"] += last["LP_Balance"]
            last["LP_Balance"] = 0.0
        results[-1] = last

    return pd.DataFrame(results)


def compute_lp_performance(wf_df):
    dates = wf_df["Date"].tolist()

    # True LP contributions ONLY come from LP_Contrib
    c = wf_df["LP_Contrib"].fillna(0).astype(float).tolist()
    contribs = [-x for x in c]  # convert to negative cashflows

    # LP distributions
    dists = wf_df["LP_Dist"].fillna(0).astype(float).tolist()

    # Net LP cashflow each month
    lp_flows = [a + b for a, b in zip(contribs, dists)]
    lp_gross_flows = [a + b for a, b in zip(c, dists)]  # includes contributions as negative

    print("LP_Contrib:", contribs)
    print("LP_Dist:", dists)
    print("LP Cashflows:", lp_flows)

    # Must have at least one negative and one positive
    if not any(x < 0 for x in lp_flows) or not any(x > 0 for x in lp_flows):
        return {"LP_Net_IRR": 0.0, "LP_Net_MOIC": 0.0}
    
    amounts = [sum(contribs)] + dists
    # IRR
    try:
        #cf_dict = dict(zip(dates, lp_flows))
        cf_dict = dict(zip(dates, amounts))
        irr_annual = xirr(cf_dict)
    except:
        irr_annual = None

    # MOIC
    total_contrib = abs(sum(x for x in lp_flows if x < 0))
    total_dist = sum(x for x in lp_flows if x > 0)
    moic = total_dist / total_contrib if total_contrib else None

    return {
        "LP_Net_IRR": 0.0 if not irr_annual == irr_annual else irr_annual,
        "LP_Net_MOIC": 0.0 if not moic == moic else moic
    }


def waterfall_to_highcharts(wf):
    """
    Converts waterfall dataframe into Highcharts format
    Row order:
      NAV
      Gross_Income
      Warehouse_Paid
      LP_Dist
      GP_Dist
      Warehouse_Balance
      LP_Balance
    """

    wf = wf.copy()

    # Ensure all required columns exist
    required_cols = [
        "Date",
        "NAV",
        "Gross_Income",
        "Warehouse_Paid",
        "LP_Dist",
        "GP_Dist",
        "Warehouse_Balance",
        "LP_Balance"
    ]

    for col in required_cols:
        if col not in wf.columns:
            wf[col] = 0

    # Reorder explicitly
    wf = wf[required_cols]

    # Convert to wide format
    wide = wf.set_index("Date").T.reset_index()
    wide = wide.rename(columns={"index": "Metric"})

    categories = [d.strftime("%Y-%m-%d") for d in wide.columns[1:]]

    series = [
        {
            "name": row["Metric"],
            "data": [row[c] for c in wide.columns[1:]]
        }
        for _, row in wide.iterrows()
    ]

    return {
        "categories": categories,
        "series": series
    }

###############################################################################
if __name__ == "__main__":
    app.run(debug=True)