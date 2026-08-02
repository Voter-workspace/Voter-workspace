from datetime import datetime, timedelta
import pandas as pd
from storage.database import get_all_entries
import streamlit as st
from supabase import create_client

# डायरेक्ट Supabase कनेक्शन
SUPABASE_URL = "https://nahwsdwzlocezbcukrmv.supabase.co"
SUPABASE_KEY = "sb_publishable_esoz3SCBUza9ufdF5Ia5eg_WD1QJ_2-"
supabase = create_client(SUPABASE_URL, SUPABASE_KEY)


def check_supabase_connection():
    try:
        res = supabase.table("work_entries").select("id", count="exact").limit(1).execute()
        return True if res else False
    except Exception:
        return False


def show():
    # CSS - Match exactly with New Entry size and fonts to prevent jumping size on switch
    st.markdown(
        """
        <style>
        header[data-testid="stHeader"] {
            visibility: hidden !important;
            display: none !important;
        }

        .block-container {
            padding: 0.8rem 1rem !important;
            max-width: 100% !important;
        }
        
        div.row-widget.stVerticalBlock {
            gap: 0.05rem !important;
        }
        
        h1, h2, h3, h4 {
            padding-top: 0px !important;
            margin-top: 0px !important;
            margin-bottom: 0.1rem !important;
            font-size: 1.15rem !important;
        }
        
        p, span, label {
            margin-bottom: 0rem !important;
            margin-top: 0px !important;
            font-size: 1rem !important;
        }

        hr {
            margin-top: 0.2rem !important;
            margin-bottom: 0.2rem !important;
        }

        .stButton > button {
            border-radius: 4px;
            font-weight: bold;
            font-size: 1rem !important;
            padding: 6px 12px !important;
            min-height: 40px !important;
        }
        
        /* 1. Approve Button -> Blue */
        div.stButton > button[kind="primary"] {
            background-color: #2563eb !important;
            color: white !important;
            border: none !important;
        }
        
        /* 2. Reject Button -> Red */
        div.stButton > button:has-text("Reject") {
            background-color: #dc2626 !important;
            color: white !important;
            border: none !important;
        }
        
        /* 3. Move Button -> Yellow/Amber */
        div.stButton > button:has-text("Move") {
            background-color: #eab308 !important;
            color: black !important;
            border: none !important;
        }

        div[data-testid="stMetric"] {
            background-color: #f1f5f9;
            padding: 6px 10px;
            border-radius: 4px;
        }
        div[data-testid="stMetric"] label {
            font-size: 0.9rem !important;
            font-weight: 600 !important;
        }
        div[data-testid="stMetric"] div[data-testid="stMetricValue"] {
            font-size: 1.2rem !important;
        }
        </style>
    """,
        unsafe_allow_html=True,
    )

    # ---------------- LOAD DATA ----------------
    try:
        rows = get_all_entries()
    except:
        rows = []

    if not rows:
        st.info("No Records Found in Database.")
        return

    df = pd.DataFrame(rows)

    if "ac_name" in df.columns:
        df["ac_name"] = (
            df["ac_name"]
            .astype(str)
            .str.replace("NCT of Delhi - ", "", case=False, regex=False)
            .str.replace("NCT of Delhi-", "", case=False, regex=False)
            .str.strip()
        )

    if "status_filter" not in st.session_state:
        st.session_state.status_filter = "All"
    if "day_filter" not in st.session_state:
        st.session_state.day_filter = "All"

    # ---------------- HEADER, DATE-TIME & CONNECTION LIGHT ----------------
    is_connected = check_supabase_connection()
    if is_connected:
        status_light_html = """
        <div style="display: flex; align-items: center; justify-content: flex-end; gap: 8px; font-size: 0.95rem; font-weight: 600; color: #166534; padding-top: 4px;">
            <span style="height: 12px; width: 12px; background-color: #22c55e; border-radius: 50%; display: inline-block; box-shadow: 0 0 8px #22c55e;"></span>
            Connected
        </div>
        """
    else:
        status_light_html = """
        <div style="display: flex; align-items: center; justify-content: flex-end; gap: 8px; font-size: 0.95rem; font-weight: 600; color: #991b1b; padding-top: 4px;">
            <span style="height: 12px; width: 12px; background-color: #ef4444; border-radius: 50%; display: inline-block; box-shadow: 0 0 8px #ef4444;"></span>
            Disconnected
        </div>
        """

    col_h1, col_h2, col_h3, col_h4 = st.columns([1.2, 1.2, 2.0, 1.0])
    with col_h1:
        current_dt = datetime.now().strftime("%d %b %Y | %I:%M %p")
        st.markdown(f"**📊 Reports 2 Workspace**<br><span style='font-size:0.85rem; color:#475569;'>🕒 {current_dt}</span>", unsafe_allow_html=True)
    
    with col_h2:
        search_by = st.selectbox(
            "Search By",
            ["All Fields", "Client Name", "Reference / ID", "Applicant Name"],
            label_visibility="collapsed",
            key="report2_search_by"
        )

    with col_h3:
        search = st.text_input(
            "Search",
            placeholder="🔍 Type to search instantly...",
            label_visibility="collapsed",
            key="report2_instant_search_box"
        )

    with col_h4:
        st.markdown(status_light_html, unsafe_allow_html=True)

    # ---------------- APPLY SEARCH FILTER ----------------
    if search:
        search_lower = search.lower()
        if search_by == "Client Name" and "client_name" in df.columns:
            df = df[df["client_name"].astype(str).str.lower().str.contains(search_lower)]
        elif search_by == "Reference / ID" and "id" in df.columns:
            df = df[df["id"].astype(str).str.lower().str.contains(search_lower)]
        elif search_by == "Applicant Name" and "applicant_name" in df.columns:
            df = df[df["applicant_name"].astype(str).str.lower().str.contains(search_lower)]
        else:
            df = df[
                df.astype(str)
                .apply(lambda x: x.str.lower())
                .apply(lambda x: x.str.contains(search_lower))
                .any(axis=1)
            ]

    # ---------------- FINANCIAL SUMMARY & FILTERS ----------------
    total_amount = sum(float(x.get("amount") or 0) for x in rows)
    received = sum(float(x.get("payment_received") or 0) for x in rows)
    balance = total_amount - received

    total_count = len(rows)
    completed_count = len([x for x in rows if str(x.get("work_status")).strip().lower() in ["completed", "complete"]])
    pending_count = len([x for x in rows if str(x.get("work_status")).strip().lower() == "pending"])

    c_fin1, c_fin2, c_fin3, c_b1, c_b2, c_b3 = st.columns([1, 1, 1, 1.2, 1.2, 1.2])
    
    with c_fin1:
        st.metric("Total", f"₹{total_amount:,.0f}")
    with c_fin2:
        st.metric("Rec", f"₹{received:,.0f}")
    with c_fin3:
        st.metric("Bal", f"₹{balance:,.0f}")

    with c_b1:
        if st.button(f"📄 Total ({total_count})", use_container_width=True, key="rep2_btn_total"):
            st.session_state.status_filter = "All"
    with c_b2:
        if st.button(f"✅ Comp ({completed_count})", use_container_width=True, key="rep2_btn_completed"):
            st.session_state.status_filter = "Complete"
    with c_b3:
        if st.button(f"⏳ Pend ({pending_count})", use_container_width=True, key="rep2_btn_pending"):
            st.session_state.status_filter = "Pending"

    d1, d2, d3, col_btn1, col_btn2, col_btn3 = st.columns([1, 1.2, 1, 1.2, 1.2, 1.2])
    with d1:
        if st.button("🔄 All Time", use_container_width=True, key="rep2_btn_all_time"):
            st.session_state.day_filter = "All"
    with d2:
        if st.button("⏳ 7+ Days", use_container_width=True, key="rep2_btn_7_plus"):
            st.session_state.day_filter = "7_plus"
    with d3:
        if st.button("⚡ < 7 Days", use_container_width=True, key="rep2_btn_less_7"):
            st.session_state.day_filter = "less_7"

    # ---------------- APPLY STATUS & DAY FILTERS ----------------
    if st.session_state.status_filter != "All":
        df = df[
            df["work_status"]
            .astype(str)
            .str.strip()
            .str.lower()
            .isin([st.session_state.status_filter.lower(), st.session_state.status_filter.lower() + "d"])
        ]

    if "day_count" in df.columns:
        df["day_count"] = pd.to_numeric(df["day_count"], errors="coerce").fillna(0)
        if st.session_state.day_filter == "7_plus":
            df = df[df["day_count"] >= 7]
        elif st.session_state.day_filter == "less_7":
            df = df[df["day_count"] < 7]

    # ---------------- TABLE & ACTION BUTTONS ----------------
    if not df.empty:
        desired_columns = ["id", "client_name", "ac_name", "applicant_name", "form_type", "day_count", "amount", "payment_received", "work_status", "remarks"]
        valid_columns = [col for col in desired_columns if col in df.columns]

        df_table = df[valid_columns].copy()
        rename_mapping = {
            "id": "Serial No", "client_name": "Client", "ac_name": "AC Name",
            "applicant_name": "Applicant", "form_type": "Form", "day_count": "Days",
            "amount": "Amt", "payment_received": "Recv", "work_status": "Status", "remarks": "Remarks"
        }
        df_table = df_table.rename(columns={k: v for k, v in rename_mapping.items() if k in df_table.columns})
        df_table.insert(0, "Select", False)

        with col_btn1:
            btn_app = st.button("✅ Approve", use_container_width=True, type="primary", key="rep2_approve_btn")
        with col_btn2:
            btn_rej = st.button("❌ Reject", use_container_width=True, key="rep2_reject_btn")
        with col_btn3:
            btn_mov = st.button("📦 Move", use_container_width=True, key="rep2_move_btn")

        # Delete Button row
        col_del_btn, _ = st.columns([1, 2])
        with col_del_btn:
            btn_del = st.button("🗑️ Delete Selected", use_container_width=True, key="rep2_delete_records_btn")

        edited_df = st.data_editor(
            df_table,
            column_config={"Select": st.column_config.CheckboxColumn("Select", default=False)},
            disabled=[col for col in df_table.columns if col not in ["Select", "Status"]],
            use_container_width=True,
            hide_index=True,
            key="report2_interactive_table",
        )

        if btn_app:
            selected_rows = edited_df[edited_df["Select"] == True]
            if not selected_rows.empty:
                for s_id in selected_rows["Serial No"].tolist():
                    supabase.table("work_entries").update({"work_status": "Complete"}).eq("id", s_id).execute()
                st.rerun()

        if btn_rej:
            selected_rows = edited_df[edited_df["Select"] == True]
            if not selected_rows.empty:
                for s_id in selected_rows["Serial No"].tolist():
                    supabase.table("work_entries").update({"work_status": "Reject"}).eq("id", s_id).execute()
                st.rerun()

        if btn_mov:
            selected_rows = edited_df[edited_df["Select"] == True]
            if not selected_rows.empty:
                for s_id in selected_rows["Serial No"].tolist():
                    supabase.table("work_entries").update({"work_status": "Moved"}).eq("id", s_id).execute()
                st.rerun()

        # Delete record from Supabase Database and refresh
        if btn_del:
            selected_rows = edited_df[edited_df["Select"] == True]
            if not selected_rows.empty:
                for s_id in selected_rows["Serial No"].tolist():
                    supabase.table("work_entries").delete().eq("id", s_id).execute()
                st.success("Record deleted successfully from database!")
                st.rerun()
            else:
                st.warning("Please select a record using the checkbox to delete.")

    else:
        st.info("No records match the selected filter or search.")