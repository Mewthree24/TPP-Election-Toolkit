import streamlit as st
import json
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from io import BytesIO

# Initialize session
if "election_data" not in st.session_state:
    st.session_state["election_data"] = {}

st.set_page_config(page_title="TPP Election Toolkit", layout="wide")
st.title("🗳️ TPP Election Toolkit")

# Upload file
uploaded_file = st.file_uploader("Upload your savefile", type=["json"])

if uploaded_file:
    try:
        raw_data = uploaded_file.read()
        data = json.loads(raw_data)

        wanted_keys = [
            "electNightSB", "electNightCC", "electNightM",
            "electNightStH", "electNightStS", "electNightG",
            "electNightUSH", "electNightUSS", "electNightP"
        ]

        extracted_data = {k: data[k] for k in wanted_keys if k in data}
        st.session_state["election_data"] = extracted_data

        st.success("Election data extracted successfully.")

    except Exception as e:
        st.error(f"Failed to load file: {e}")

if st.session_state["election_data"]:

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

    election_types = {
        "President": "electNightP",
        "Senate": "electNightUSS",
        "Governor": "electNightG",
        "U.S. House": "electNightUSH",
        "State House (Player State)": "electNightStH",
        "State Senate (Player State)": "electNightStS"
    }

    available_election_types = [etype for etype, ekey in election_types.items() if ekey in st.session_state["election_data"]]

    if available_election_types:
        selected_election_type = st.selectbox("Select Election Type", available_election_types)
        election_key = election_types[selected_election_type]
        election_data = st.session_state["election_data"][election_key]

        entries_to_convert = []

        if selected_election_type in ["President", "Senate", "Governor"]:
            available_states = [entry.get("state") for entry in election_data.get("elections", [])]
            available_states = sorted(set(available_states))
            state_options = ["National View"] + [state_code_to_name.get(code, code) for code in available_states]
            selected_state = st.selectbox("Select State", state_options)

            if selected_state != "National View":
                state_code = next((code for code, name in state_code_to_name.items() if name == selected_state), None)
                if state_code:
                    state_entries = [entry for entry in election_data.get("elections", []) if entry.get("state") == state_code]
                    if state_entries:
                        state_entry = state_entries[0]
                        counties = state_entry.get("counties", [])
                        spreadsheet_rows = []

                        for county in counties:
                            county_name = county.get("name", "Unknown County")
                            for cand in county.get("cands", []):
                                spreadsheet_rows.append({
                                    "State": state_code,
                                    "County": county_name,
                                    "Candidate": cand.get("name", ""),
                                    "Party": cand.get("party", ""),
                                    "Votes": cand.get("votes", 0),
                                    "Incumbent": cand.get("incumbent", False),
                                    "Caucus": cand.get("caucus", ""),
                                })

                        if spreadsheet_rows:
                            df = pd.DataFrame(spreadsheet_rows)
                            st.subheader(f"🧾 {selected_state} County-Level Results")
                            st.dataframe(df)

                          # Generate XLSX download
                            wb = Workbook()
                            ws = wb.active
                            ws.title = f"{state_code} County Results"

                            # Determine which parties exist
                            parties_present = sorted({cand.get("party") for cand in state_entry.get("cands", [])})
                            party_order = ["D", "R", "I"]
                            parties = [p for p in party_order if p in parties_present]

                            # Header row 1: party names (merged)
                            col = 2
                            for party in parties:
                                ws.cell(row=1, column=col, value={"D": "Democratic", "R": "Republican", "I": "Independent"}.get(party, party))
                                ws.merge_cells(start_row=1, start_column=col, end_row=1, end_column=col + 1)
                                col += 2

                            ws.cell(row=1, column=col, value="Margins & Rating")
                            ws.merge_cells(start_row=1, start_column=col, end_row=1, end_column=col + 3)

                            # Header row 2
                            col = 1
                            ws.cell(row=2, column=col, value="County")
                            col += 1
                            for party in parties:
                                ws.cell(row=2, column=col, value="Candidate")
                                ws.cell(row=2, column=col + 1, value="%")
                                col += 2
                            ws.cell(row=2, column=col, value="Margin #")
                            ws.cell(row=2, column=col + 1, value="Margin %")
                            ws.cell(row=2, column=col + 2, value="Total Vote")
                            ws.cell(row=2, column=col + 3, value="Rating")

                            # Format headers
                            for r in range(1, 3):
                                for c in range(1, col + 4):
                                    cell = ws.cell(row=r, column=c)
                                    cell.font = Font(bold=True)
                                    cell.alignment = Alignment(horizontal="center", vertical="center")

                            # Body rows (just filling in counties for now)
                            for i, county in enumerate(counties, start=3):
                                ws.cell(row=i, column=1, value=county.get("name", "Unknown County"))

                            # Save file to memory
                            file_stream = BytesIO()
                            wb.save(file_stream)
                            file_stream.seek(0)

                            st.download_button(
                                label="📥 Download County-Level Spreadsheet",
                                data=file_stream,
                                file_name=f"{state_code}_county_results.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                        else:
                            st.warning("No county-level data found.")
                else:
                    st.warning("Selected state not found.")
            else:
                # National View logic (flatten by state)
                entries_to_convert = election_data.get("elections", [])
                spreadsheet_rows = []
                for entry in entries_to_convert:
                    state = entry.get("state", "Unknown")
                    for cand in entry.get("cands", []):
                        spreadsheet_rows.append({
                            "State": state,
                            "Candidate": cand.get("name", ""),
                            "Party": cand.get("party", ""),
                            "Votes": cand.get("votes", 0),
                            "Incumbent": cand.get("incumbent", False),
                            "Caucus": cand.get("caucus", ""),
                        })

                if spreadsheet_rows:
                    df = pd.DataFrame(spreadsheet_rows)
                    st.subheader("🧾 National Election Results")
                    st.dataframe(df)
                else:
                    st.warning("No state-level data found.")
        else:
            # For U.S. House, State House, State Senate — national-level only
            entries_to_convert = election_data.get("elections", [])
            spreadsheet_rows = []
            for entry in entries_to_convert:
                state = entry.get("state", "Unknown")
                for cand in entry.get("cands", []):
                    spreadsheet_rows.append({
                        "State": state,
                        "Candidate": cand.get("name", ""),
                        "Party": cand.get("party", ""),
                        "Votes": cand.get("votes", 0),
                        "Incumbent": cand.get("incumbent", False),
                        "Caucus": cand.get("caucus", ""),
                    })

            if spreadsheet_rows:
                df = pd.DataFrame(spreadsheet_rows)
                st.subheader("🧾 Election Results Spreadsheet")
                st.dataframe(df)
            else:
                st.warning("No data available to display.")
    else:
        st.warning("No recognized election data found in this file.")
else:
    st.info("Please upload a JSON savefile.")
