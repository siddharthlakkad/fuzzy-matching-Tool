import streamlit as st
import pandas as pd
from rapidfuzz import process, fuzz
import io

# Function to perform fuzzy matching with a threshold
def fuzzy_match(val, choices, df2, threshold=0):
    if pd.isna(val):  # Check if the value is NaN (empty cell)
        return None, 0, pd.Series([None]*len(df2.columns), index=df2.columns)  # Return no match and score of 0
    
    result = process.extractOne(val, choices, scorer=fuzz.ratio)
    if result:  # If result is not None
        best_match, score, idx = result  # Unpack best match, score, and index
        if score >= threshold:  # Check if score meets the threshold
            matched_row = df2.iloc[idx]  # Get the row from df2 that matches the best result
            return best_match, score, matched_row
    return None, 0, pd.Series([None]*len(df2.columns), index=df2.columns)  # Return no match and score of 0 if no match or below threshold

# Streamlit App
st.title('Fuzzy Matching System for Excel Columns')

# File upload
st.header('Upload two Excel files for comparison')
uploaded_file1 = st.file_uploader("Choose the first Excel file", type="xlsx")
uploaded_file2 = st.file_uploader("Choose the second Excel file", type="xlsx")

if uploaded_file1 and uploaded_file2:
    # Read Excel files into pandas DataFrames
    df1 = pd.read_excel(uploaded_file1)
    df2 = pd.read_excel(uploaded_file2)
    
    # Display column selection options
    st.write("Select columns from each file to compare:")
    
    col1 = st.selectbox('Select Column from First File', df1.columns)
    col2 = st.selectbox('Select Column from Second File', df2.columns)

    # Add a slider to set the threshold
    threshold = st.slider('Set Matching Threshold (0-100)', min_value=0, max_value=100, value=80)
    
    if st.button('Start Matching'):
        # Perform fuzzy matching between selected columns
        results = []
        for _, row in df1.iterrows():
            value = row[col1]
            best_match, score, matched_row = fuzzy_match(value, df2[col2], df2, threshold=threshold)
            if best_match is not None and score > 0:  # Only include rows with valid matches
                combined_row = {**row.to_dict(), 'Score': score, **matched_row.to_dict()}
                results.append(combined_row)
        
        # Convert results to DataFrame
        result_df = pd.DataFrame(results)
        
        # Display results in Streamlit (only rows with best matches)
        st.write(f"Fuzzy Matching Results (Threshold: {threshold})")
        if not result_df.empty:
            st.dataframe(result_df)
        else:
            st.write("No matches found above the given threshold.")
        
        # Provide download option
        if not result_df.empty:
            towrite = io.BytesIO()
            downloaded_file = result_df.to_excel(towrite, index=False, engine='openpyxl')
            towrite.seek(0)
            st.download_button(label="Download results as Excel", data=towrite, file_name="fuzzy_match_results.xlsx", mime="application/vnd.ms-excel")
