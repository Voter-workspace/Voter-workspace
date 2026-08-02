from datetime import date, datetime
import streamlit as st

from storage.database import add_work_entry


def show():
    # CSS - Force standard legible font size across all widgets and hide default header
    st.markdown(
        """
        <style>
        header[data-testid="stHeader"] {
            visibility: hidden !important;
            display: none !important;
        }
        .block-container {
            padding: 1rem 1.5rem !important;
            max-width: 100% !important;
        }
        /* Target all Streamlit widget labels directly for standard clean sizing */
        label div p, .stTextInput label, .stSelectbox label, .stDateInput label, .stNumberInput label {
            font-size: 1rem !important;
            font-weight: 600 !important;
            color: #1e293b !important;
        }
        .stButton > button {
            font-size: 1rem !important;
            font-weight: bold !important;
            padding: 8px 12px !important;
            min-height: 42px !important;
        }
        div[data-baseweb="input"] input, div[data-baseweb="select"] div {
            min-height: 40px !important;
            font-size: 1rem !important;
        }
        </style>
    """,
        unsafe_allow_html=True,
    )

    st.markdown("### ➕ New Work Entry Form")

    if "form_data" not in st.session_state:
        st.session_state.form_data = {
            "ref": "", "state": "", "ac": "", "applicant": "",
            "form": "FORM6", "date": date.today(), "status": "",
        }

    paste_data = st.text_area("Paste Excel Data (Reference First)", height=45, placeholder="Paste Excel row here...", label_visibility="visible")

    col_b1, col_b2 = st.columns(2)
    if col_b1.button("⚡ Auto Fill Form", use_container_width=True):
        if paste_data.strip():
            try:
                first_line = paste_data.strip().split("\n")[0]
                if not ("Reference" in first_line or "Client" in first_line):
                    cols = first_line.split("\t")
                    while len(cols) < 8:
                        cols.append("")
                    st.session_state.form_data = {
                        "ref": cols[0].strip(), "state": cols[1].strip(), "ac": cols[2].strip(),
                        "applicant": f"{cols[3].strip()} {cols[4].strip()}".strip(),
                        "form": cols[5].strip() or "FORM6",
                        "date": datetime.strptime(cols[6].strip(), "%d/%m/%Y").date() if cols[6].strip() else date.today(),
                        "status": cols[7].strip(),
                    }
                    st.rerun()
            except:
                pass

    if col_b2.button("🗑 Clear Form", use_container_width=True):
        st.session_state.form_data = {
            "ref": "", "state": "", "ac": "", "applicant": "",
            "form": "FORM6", "date": date.today(), "status": "",
        }
        st.rerun()

    st.divider()

    left_col, right_col = st.columns([1.1, 0.9])

    with left_col:
        st.markdown("#### 📝 Application Details")
        reference_no = st.text_input("Reference No *", value=st.session_state.form_data["ref"])
        client_name = st.text_input("Client Name * (Manual)", value="")
        
        lc1, lc2 = st.columns(2)
        with lc1:
            state = st.text_input("State", value=st.session_state.form_data["state"])
            applicant_name = st.text_input("Applicant Name", value=st.session_state.form_data["applicant"])
        with lc2:
            ac_name = st.text_input("AC Name", value=st.session_state.form_data["ac"])
            
            form_options = ["FORM6", "FORM7", "FORM8", "Correction", "Other"]
            current_form = st.session_state.form_data["form"]
            form_type = st.selectbox("Form Type", form_options, index=form_options.index(current_form) if current_form in form_options else 0)

    with right_col:
        st.markdown("#### 💰 Payment & Work Status")
        rc1, rc2 = st.columns(2)
        with rc1:
            amount = st.number_input("Total Amount", min_value=0.0, step=100.0, value=0.0)
            submission_date = st.date_input("Submission Date", value=st.session_state.form_data["date"])
            work_status = st.selectbox("Work Status", ["Pending", "In Progress", "Completed", "Rejected"])
        with rc2:
            payment_received = st.number_input("Payment Received", min_value=0.0, step=100.0, value=0.0)
            balance = amount - payment_received
            st.markdown(f"<p style='margin-top:22px; font-size:1.1rem; font-weight:bold; color:#0f172a;'>Balance: ₹{balance:,.0f}</p>", unsafe_allow_html=True)
            remarks = st.text_input("Remarks", value="")

    if balance <= 0 and amount > 0:
        payment_status = "Paid"
    elif payment_received > 0:
        payment_status = "Partial"
    else:
        payment_status = "Pending"

    st.divider()

    if st.button("💾 Save Entry to Cloud", use_container_width=True, type="primary"):
        if not reference_no.strip() or not client_name.strip():
            st.error("⚠️ Reference Number और Client Name दोनों भरना अनिवार्य है!")
        else:
            day_count = max(0, (date.today() - submission_date).days)
            data = {
                "reference_no": reference_no.strip(),
                "client_name": client_name.strip(),
                "applicant_name": applicant_name.strip(),
                "mobile": "",
                "ac_name": f"{state} - {ac_name}" if state else ac_name,
                "part_no": "",
                "form_type": form_type,
                "submission_date": str(submission_date),
                "work_status": work_status,
                "work_done": True if work_status == "Completed" else False,
                "day_count": day_count,
                "amount": float(amount),
                "payment_received": float(payment_received),
                "balance_amount": float(balance),
                "payment_status": payment_status,
                "remarks": remarks.strip(),
            }
            try:
                add_work_entry(data)
                st.success("🎉 रिकॉर्ड सफलतापूर्वक Supabase में सेव हो गया है!")
                st.balloons()
            except Exception as e:
                st.error(f"❌ Error: {e}")