import streamlit as st
import json

st.set_page_config(page_title="TPP Election Toolkit", layout="wide")

st.title("🗳️ TPP Election Toolkit")

uploaded_file = st.file_uploader("Upload your savefile", type=["json"])

if uploaded_file:
    try:
        raw_data = uploaded_file.read()
        data = json.loads(raw_data)

        # Top-level keys we care about
        wanted_keys = [
            "electNightSB", "electNightCC", "electNightM",
            "electNightStH", "electNightStS", "electNightG",
            "electNightUSH", "electNightUSS", "electNightP"
        ]

        # Grab only what's present
        extracted_data = {k: data[k] for k in wanted_keys if k in data}

        st.success("Election data extracted from top-level.")
        st.write("Included election types:", list(extracted_data.keys()))
        st.session_state["election_data"] = extracted_data

    except Exception as e:
        st.error(f"Failed to load file: {e}")
else:
    st.info("Please upload a JSON savefile.")

import streamlit as st

# Define mapping from state codes to full names
state_code_to_name = {
    "AL": "Alabama", "AK": "Alaska", "AZ": "Arizona", "AR": "Arkansas",
    "CA": "California", "CO": "Colorado", "CT": "Connecticut", "DE": "Delaware",
    "FL": "Florida", "GA": "Georgia", "HI": "Hawaii", "ID": "Idaho",
    "IL": "Illinois", "IN": "Indiana", "IA": "Iowa", "KS": "Kansas",
    "KY": "Kentucky", "LA": "Louisiana", "ME": "Maine", "MD": "Maryland",
    "MA": "Massachusetts", "MI": "Michigan", "MN": "Minnesota", "MS": "Mississippi",
    "MO": "Missouri", "MT": "Montana", "NE": "Nebraska", "NV": "Nevada",
    "NH": "New Hampshire", "NJ": "New Jersey", "NM": "New Mexico", "NY": "New York",
    "NC": "North Carolina", "ND": "North Dakota", "OH": "Ohio", "OK": "Oklahoma",
    "OR": "Oregon", "PA": "Pennsylvania", "RI": "Rhode Island", "SC": "South Carolina",
    "SD": "South Dakota", "TN": "Tennessee", "TX": "Texas", "UT": "Utah",
    "VT": "Vermont", "VA": "Virginia", "WA": "Washington", "WV": "West Virginia",
    "WI": "Wisconsin", "WY": "Wyoming", "DC": "District of Columbia"
}

# Define election types and their corresponding keys in the JSON
election_types = {
    "President": "electNightP",
    "Senate": "electNightUSS",
    "Governor": "electNightG",
    "U.S. House": "electNightUSH",
    "State House (Player State)": "electNightStH",
    "State Senate (Player State)": "electNightStS"
}

# Primary dropdown for election type
selected_election_type = st.selectbox("Select Election Type", list(election_types.keys()))

# Retrieve the corresponding key from the JSON
election_key = election_types[selected_election_type]

if "election_data" not in st.session_state:
st.session_state["election_data"] = {}

# Check if the selected election type exists in the data
if election_key in st.session_state["election_data"]:
    election_data = st.session_state["election_data"][election_key]

    # For President, Senate, and Governor, provide a secondary dropdown for state selection
    if selected_election_type in ["President", "Senate", "Governor"]:
        # Extract available states from the election data
        available_states = [entry.get("state") for entry in election_data.get("elections", [])]
        # Remove duplicates and sort
        available_states = sorted(set(available_states))
        # Map state codes to full names
        state_options = ["National View"] + [state_code_to_name.get(code, code) for code in available_states]
        # Secondary dropdown for state selection
        selected_state = st.selectbox("Select State", state_options)

        # Filter data based on state selection
        if selected_state != "National View":
            # Find the state code from the full name
            state_code = next((code for code, name in state_code_to_name.items() if name == selected_state), None)
            if state_code:
                # Filter the election data for the selected state
                state_data = [entry for entry in election_data.get("elections", []) if entry.get("state") == state_code]
                st.write(f"Results for {selected_state}:")
                st.json(state_data)
            else:
                st.warning("Selected state not found in the data.")
        else:
            st.write("National Results:")
            st.json(election_data.get("elections", []))
    else:
        # For other election types, display the data directly
        st.write(f"{selected_election_type} Results:")
        st.json(election_data.get("elections", []))
else:
    st.warning(f"No data available for {selected_election_type}.")


