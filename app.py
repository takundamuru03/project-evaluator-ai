import streamlit as st
import pandas as pd
import numpy as np
import numpy_financial as npf

st.title("📊 Project Evaluator AI")

# Upload Excel
file = st.file_uploader("Upload Excel File", type=["xlsx"])

if file is not None:
    # Read Excel file
    projects = pd.read_excel(file, sheet_name='Projects')
    params = pd.read_excel(file, sheet_name='Parameters')
    
    # Get discount rate
    discount_rate = params.loc[params['Parameter'] == "Discount Rate", "Value"].values[0]
    
    results = []
    
    for i in range(len(projects)):
        row = projects.iloc[i]
        
        name = row["Project"]
        cashflows = row.drop("Project").values.astype(float)

        # Payback Period
        cumulative = np.cumsum(cashflows)
        payback = np.argmax(cumulative >= 0) + 1 if any(cumulative >= 0) else None

        # NPV
        npv = npf.npv(discount_rate, cashflows)

        # Profitability Index
        pi = (npv / abs(cashflows[0])) + 1 if cashflows[0] != 0 else None

        # IRR
        try:
            irr = npf.irr(cashflows)
        except:
            irr = None

        # MIRR
        try:
            finance_rate = params.loc[params['Parameter'] == "Finance Rate", "Value"].values[0]
            reinvest_rate = params.loc[params['Parameter'] == "Reinvestment Rate", "Value"].values[0]
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
    if not results_df.empty:
        best_project = results_df.loc[results_df["NPV"].idxmax()]
        st.success(f"✅ Recommended Project: {best_project['Project']}")
else:
    st.info("Please upload an Excel file to proceed.")
