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

                            wb = Workbook()
                            ws = wb.active
                            ws.title = f"{state_code} County Results"

                            # === Create party blocks dynamically from candidates ===
                            party_labels = {"D": "Democratic", "R": "Republican", "I": "Independent"}

                            candidates = state_entry.get("cands", [])
                            party_to_candidates = {}

                            for cand in candidates:
                                party = cand["party"]
                                if party not in party_to_candidates:
                                    party_to_candidates[party] = []
                                party_to_candidates[party].append(cand["name"])

                            party_order = list(party_to_candidates.keys())

                            # === Header rows ===
                            ws.cell(row=2, column=1, value="County")
                            col = 2

                            for party in party_order:
                                full_party_name = party_labels.get(party, party)
                                ws.cell(row=1, column=col, value=full_party_name)
                                ws.merge_cells(start_row=1, start_column=col, end_row=1, end_column=col + 1)

                                candidate_names = party_to_candidates[party]
                                if candidate_names:
                                    ws.cell(row=2, column=col, value=candidate_names[0])
                                else:
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
                            
                            # === Data rows ===
                            totals = {party: 0 for party in parties}
                            row_idx = 3
                            
                            for county in counties:
                                clean_name = county.get("name", "Unknown County").replace(" County", "").title()
                                ws.cell(row=row_idx, column=1, value=clean_name)
                                county_votes = {c["party"]: round(c["votes"], 2) for c in county.get("cands", [])}
                                total_vote = sum(county_votes.values())
                                vote_values = []
                            
                                col = 2
                                for party in parties:
                                    v = int(round(county_votes.get(party, 0)))
                                    pct = round(v / total_vote * 100, 2) if total_vote else 0
                                    ws.cell(row=row_idx, column=col, value="{:,}".format(v))
                                    ws.cell(row=row_idx, column=col + 1, value="{:.2f}%".format(pct))
                                    col += 2
                                    totals[party] += v
                                    vote_values.append((party, v))
                            
                                vote_values.sort(key=lambda x: x[1], reverse=True)
                                margin = vote_values[0][1] - vote_values[1][1] if len(vote_values) > 1 else vote_values[0][1]
                                margin_pct = round(margin / total_vote * 100, 2) if total_vote else 0
                                rating = "Tilt" if margin_pct < 1 else "Lean" if margin_pct < 5 else "Likely" if margin_pct < 10 else "Safe"
                                winner_party = vote_values[0][0] if vote_values else "?"
                                rating_label = f"{rating} {party_codes.get(winner_party, winner_party)}"
                            
                                ws.cell(row=row_idx, column=col, value="{:,}".format(margin))
                                ws.cell(row=row_idx, column=col + 1, value="{:.2f}%".format(margin_pct))
                                ws.cell(row=row_idx, column=col + 2, value="{:,}".format(int(round(total_vote))))
                                ws.cell(row=row_idx, column=col + 3, value=rating_label)
                                row_idx += 1
                            
                            # === Totals row ===
                            ws.cell(row=row_idx, column=1, value="TOTALS")
                            grand_total = sum(totals.values())
                            col = 2
                            sorted_totals = [(p, totals[p]) for p in parties]
                            for party in parties:
                                v = int(round(totals[party]))
                                pct = round(v / grand_total * 100, 2) if grand_total else 0
                                ws.cell(row=row_idx, column=col, value="{:,}".format(v))
                                ws.cell(row=row_idx, column=col + 1, value="{:.2f}%".format(pct))
                                col += 2
                            
                            sorted_totals.sort(key=lambda x: x[1], reverse=True)
                            margin = sorted_totals[0][1] - sorted_totals[1][1] if len(sorted_totals) > 1 else sorted_totals[0][1]
                            margin_pct = round(margin / grand_total * 100, 2) if grand_total else 0
                            rating = "Tilt" if margin_pct < 1 else "Lean" if margin_pct < 5 else "Likely" if margin_pct < 10 else "Safe"
                            winner_party = sorted_totals[0][0] if sorted_totals else "?"
                            rating_label = f"{rating} {party_codes.get(winner_party, winner_party)}"
                            
                            ws.cell(row=row_idx, column=col, value="{:,}".format(margin))
                            ws.cell(row=row_idx, column=col + 1, value="{:.2f}%".format(margin_pct))
                            ws.cell(row=row_idx, column=col + 2, value="{:,}".format(int(round(grand_total))))
                            ws.cell(row=row_idx, column=col + 3, value=rating_label)
                        
                        from openpyxl.styles import Font

                        # Bold Totals row
                        for col_idx in range(1, col + 4):
                            ws.cell(row=row_idx, column=col_idx).font = Font(bold=True)
                        
                        # Save file to memory
                        file_stream = BytesIO()
                        wb.save(file_stream)
                        file_stream.seek(0)
                        
                        # Create download button (one time only)
                        st.download_button(
                            label="📥 Download County-Level Spreadsheet",
                            data=file_stream,
                            file_name=f"{state_code}_{selected_election_type.replace(' ', '')}_County_Results.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key=f"county_download_{state_code}"
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