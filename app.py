import streamlit as st
import pandas as pd
import numpy as np
import math
import io  # <-- We’ll use BytesIO from the io module

def erlang_c(traffic_intensity, agents, target_answer_time, aht):
    if agents <= 0 or traffic_intensity >= agents:
        return 1.0

    rho = traffic_intensity / agents
    fact = math.factorial

    sum_terms = sum((traffic_intensity ** n) / fact(n) for n in range(agents))
    erlang_c_prob = ((traffic_intensity ** agents) / (fact(agents) * (1 - rho)))
    erlang_c_prob /= (sum_terms + erlang_c_prob)

    pwait = erlang_c_prob * np.exp(-(agents - traffic_intensity) * target_answer_time / aht)
    return pwait

st.title("📞 Erlang C Forecasting Dashboard")
uploaded_file = st.file_uploader("Upload your average call volume per interval (.xlsx)", type="xlsx")

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    
    # Set defaults
    aht = st.number_input("Average Handle Time (seconds)", value=360)
    interval_length = st.number_input("Interval Length (minutes)", value=30)
    target_service_level = st.slider("Service Level Target (%)", 50, 100, 80)
    target_answer_time = st.number_input("Target Answer Time (seconds)", value=30)
    shrinkage = st.slider("Shrinkage (%)", 0, 50, 17)
    
    # Calculate Erlangs and required agents
    df['Workload (Erlangs)'] = (df['Inbound Calls'] * aht) / (interval_length * 60)
    
    def calc_agents_needed(erlangs):
        for agents in range(1, 100):
            if erlang_c(erlangs, agents, target_answer_time, aht) >= (target_service_level / 100):
                return agents
        return 100
    
    df['Agents Needed (No Shrinkage)'] = df['Workload (Erlangs)'].apply(calc_agents_needed)
    df['Agents Needed (With Shrinkage)'] = (df['Agents Needed (No Shrinkage)'] / (1 - shrinkage / 100)).apply(np.ceil)
    
    st.subheader("📊 Forecast Results")
    st.dataframe(df)
    
    # Write the DataFrame to an in-memory buffer
    buffer = io.BytesIO()
    df.to_excel(buffer, index=False)
    
    # Use the buffer's contents in the download button
    st.download_button(
        label="Download Forecast as Excel",
        data=buffer.getvalue(),
        file_name="Erlang_Forecast_Updated.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
