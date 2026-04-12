import streamlit as st
import pandas as pd
import numpy as np
import numpy_financial as npf

st.title("📊 Finance AI Agent - Project Evaluation")

# Upload Excel
file = st.file_uploader("Upload Excel File", type=["xlsx"])

if file is not None:
    try:
        projects = pd.read_excel(file, sheet_name='Projects')
        params = pd.read_excel(file, sheet_name='Parameters')
    except Exception as e:
        st.error(f"Error reading Excel file: {e}")
        st.stop()
    
    # Extract parameters
    discount_rate = params.loc[params['Parameter'] == "Discount Rate", "Value"].values[0]
    finance_rate = params.loc[params['Parameter'] == "Finance Rate", "Value"].values[0]
    reinvest_rate = params.loc[params['Parameter'] == "Reinvestment Rate", "Value"].values[0]
    capital_budget = params.loc[params['Parameter'] == "Capital Budget", "Value"].values[0]
    divisible = params.loc[params['Parameter'] == "Divisible", "Value"].values[0]
    
    results = []
    
    for i in range(len(projects)):
        row = projects.iloc[i]
        
        name = row["Project"]
        cashflows = row.drop("Project").values.astype(float)
        
        if cashflows[0] >= 0:
            st.error(f"Project {name}: First cash flow must be negative (initial investment). Skipping this project.")
            continue
        
        initial_investment = abs(cashflows[0])
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
            mirr = npf.mirr(cashflows, finance_rate, reinvest_rate)
        except:
            mirr = None

        initial_investment = abs(cashflows[0])
        results.append([name, initial_investment, payback, npv, pi, irr, mirr])
    results_df = pd.DataFrame(results, columns=[
    "Project", "Initial Investment", "Payback", "NPV", "PI", "IRR", "MIRR"
])
    st.subheader("📈 Evaluation Results")
    st.dataframe(results_df)

    # Recommendation
    if not results_df.empty:
        best_project = results_df.loc[results_df["NPV"].idxmax()]
        st.success(f"✅ Recommended Project: {best_project['Project']}")

        if divisible == "Yes":
            st.subheader("💰 Capital Rationing (Divisible Projects)")

            df = results_df.sort_values(by="PI", ascending=False)

            budget = capital_budget
            selected = []
            total_npv = 0

            for _, row in df.iterrows():
                if budget <= 0:
                    break
                
                invest = row["Initial Investment"]

                if invest <= budget:
                    selected.append((row["Project"], 1))
                    budget -= invest
                    total_npv += row["NPV"]
                else:
                    fraction = budget / invest
                    selected.append((row["Project"], round(fraction, 2)))
                    total_npv += row["NPV"] * fraction
                    budget = 0

            st.write("### Selected Projects (with fractions):")
            selected_df = pd.DataFrame(selected, columns=["Project", "Fraction"])
            st.dataframe(selected_df)
            st.success(f"Total NPV: {round(total_npv, 2)} | Remaining Budget: {round(budget, 2)}")

        else:
            st.subheader("💰 Capital Rationing (Indivisible Projects)")

            from itertools import combinations

            projects_list = results_df.to_dict("records")

            if len(projects_list) > 15:
                st.warning("Many projects detected. Indivisible rationing may be slow due to brute-force calculation.")

            best_npv = -np.inf
            best_combo = []

            for r in range(1, len(projects_list) + 1):
                for combo in combinations(projects_list, r):
                    total_cost = sum(p["Initial Investment"] for p in combo)
                    total_npv = sum(p["NPV"] for p in combo)

                    if total_cost <= capital_budget and total_npv > best_npv:
                        best_npv = total_npv
                        best_combo = combo

            selected_projects = [p["Project"] for p in best_combo]
            total_cost = sum(p["Initial Investment"] for p in best_combo)

            st.write("### Optimal Project Combination:")
            st.write(", ".join(selected_projects))
            st.success(f"Total NPV: {round(best_npv, 2)} | Total Cost: {round(total_cost, 2)} | Remaining Budget: {round(capital_budget - total_cost, 2)}")
else:
    st.info("Please upload an Excel file to proceed.")