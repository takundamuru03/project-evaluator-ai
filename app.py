import streamlit as st
import pandas as pd
import numpy as np
import numpy_financial as npf

st.title("📊 Finance AI Agent - Project Evaluation")

# Upload Excel
file = st.file_uploader("Upload Excel File", type=["xlsx"])

if file:
    projects = pd.read_excel(file, sheet_name="Projects")
    params = pd.read_excel(file, sheet_name="Parameters")

    discount_rate = params.loc[params['Parameter']=="Discount Rate","Value"].values[0]

    results = []

    for i in range(len(projects)):
        row = projects.iloc[i]
        name = row[0]
        cashflows = row[1:].values.astype(float)

        # Payback Period
        cumulative = np.cumsum(cashflows)
        payback = np.argmax(cumulative >= 0) if any(cumulative >= 0) else None

        # NPV
        npv = npf.npv(discount_rate, cashflows)

        # Profitability Index
        pi = (npv / abs(cashflows[0])) + 1

        # IRR
        try:
            irr = npf.irr(cashflows)
        except:
            irr = None

        # MIRR
        try:
            finance_rate = params.loc[params['Parameter']=="Finance Rate","Value"].values[0]
            reinvest_rate = params.loc[params['Parameter']=="Reinvestment Rate","Value"].values[0]
            mirr = npf.mirr(cashflows, finance_rate, reinvest_rate)
        except:
            mirr = None

        results.append([name, payback, npv, pi, irr, mirr])

    results_df = pd.DataFrame(results, columns=[
        "Project", "Payback", "NPV", "PI", "IRR", "MIRR"
    ])

    st.subheader("📈 Evaluation Results")
    st.dataframe(results_df)

    # Recommendation
    best_project = results_df.loc[results_df["NPV"].idxmax()]

    st.success(f"✅ Recommended Project: {best_project['Project']}")