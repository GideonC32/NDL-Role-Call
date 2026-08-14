import streamlit as st
import sqlite3
import datetime
import os
import pandas as pd
from excel_generator import generate_excel_report
from pdf_generator import generate_pdf_report

DB_PATH = '/workspace/ndl_monitoring.db'

# Page config
st.set_page_config(
    page_title="NDL Learner Monitoring & Roll-Call",
    page_icon="👥",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for Mobile-First styling
st.markdown("""
<style>
    /* Styling headers */
    .main-title {
        color: #1F4E78;
        font-weight: bold;
        font-size: 2.2rem;
        margin-bottom: 0px;
    }
    .sub-title {
        color: #7F7F7F;
        font-size: 1rem;
        margin-bottom: 20px;
    }
    .section-header {
        color: #1F4E78;
        font-size: 1.4rem;
        font-weight: bold;
        border-bottom: 2px solid #0070C0;
        padding-bottom: 5px;
        margin-top: 15px;
        margin-bottom: 15px;
    }
    /* Mobile-friendly large status badge styles */
    .badge {
        display: inline-block;
        padding: 6px 12px;
        font-weight: bold;
        border-radius: 4px;
        text-align: center;
        font-size: 0.9rem;
    }
    .badge-present {
        background-color: #E2EFDA;
        color: #375623;
        border: 1px solid #375623;
    }
    .badge-absent {
        background-color: #FCE4D6;
        color: #C65911;
        border: 1px solid #C65911;
    }
    .badge-none {
        background-color: #F2F2F2;
        color: #595959;
        border: 1px solid #595959;
    }
    /* Quick Roll Call styling */
    .roll-call-card {
        background-color: #F9FAFB;
        padding: 15px;
        border-radius: 8px;
        border-left: 5px solid #0070C0;
        margin-bottom: 12px;
    }
    /* Hide top margin */
    .stApp {
        margin-top: -30px;
    }
</style>
""", unsafe_allow_html=True)

# Database Helper Functions
def get_db_connection():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def load_learners():
    conn = get_db_connection()
    learners = conn.execute("SELECT * FROM learners ORDER BY name ASC").fetchall()
    conn.close()
    return [dict(l) for l in learners]

def get_learner_by_id(learner_id):
    conn = get_db_connection()
    learner = conn.execute("SELECT * FROM learners WHERE id = ?", (learner_id,)).fetchone()
    conn.close()
    return dict(learner) if learner else None

def load_updates(filters=None):
    conn = get_db_connection()
    query = "SELECT * FROM updates"
    params = []
    
    if filters:
        conditions = []
        for col, val in filters.items():
            if val:
                if col == 'search':
                    conditions.append("(learner_name LIKE ? OR assigned_teacher LIKE ?)")
                    params.extend([f"%{val}%", f"%{val}%"])
                else:
                    conditions.append(f"{col} = ?")
                    params.append(val)
        if conditions:
            query += " WHERE " + " AND ".join(conditions)
            
    query += " ORDER BY timestamp DESC"
    updates = conn.execute(query, params).fetchall()
    conn.close()
    return [dict(u) for u in updates]

def get_latest_status_map():
    conn = get_db_connection()
    query = '''
    SELECT 
        l.id as learner_id,
        u.status,
        u.event,
        u.recorded_by,
        u.date,
        u.time,
        u.note
    FROM learners l
    LEFT JOIN (
        SELECT u1.* FROM updates u1
        JOIN (
            SELECT learner_id, MAX(timestamp) as max_ts 
            FROM updates 
            GROUP BY learner_id
        ) u2 ON u1.learner_id = u2.learner_id AND u1.timestamp = u2.max_ts
    ) u ON l.id = u.learner_id
    '''
    rows = conn.execute(query).fetchall()
    conn.close()
    return {r['learner_id']: dict(r) for r in rows}

def save_roll_call(teacher_name, event_name, attendance_dict, notes_dict=None):
    conn = get_db_connection()
    now = datetime.datetime.now()
    date_str = now.strftime('%Y-%m-%d')
    time_str = now.strftime('%H:%M:%S')
    
    # Check if there are duplicate saves within the last 3 seconds for the same event and teacher to prevent duplicates
    recent_check = conn.execute('''
        SELECT COUNT(*) FROM updates 
        WHERE recorded_by = ? AND event = ? AND date = ? 
        AND timestamp >= datetime('now', '-3 seconds')
    ''', (teacher_name, event_name, date_str)).fetchone()[0]
    
    if recent_check > 0:
        conn.close()
        return False, "Duplicate roll call submission blocked. Your changes are already saved!"
        
    for learner_id, status in attendance_dict.items():
        learner = conn.execute("SELECT name, assigned_teacher FROM learners WHERE id = ?", (learner_id,)).fetchone()
        if learner:
            note = notes_dict.get(learner_id, "") if notes_dict else ""
            conn.execute('''
                INSERT INTO updates (learner_id, learner_name, assigned_teacher, recorded_by, event, status, note, date, time)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (learner_id, learner['name'], learner['assigned_teacher'], teacher_name, event_name, status, note, date_str, time_str))
            
    conn.commit()
    conn.close()
    return True, "Roll Call saved successfully!"

def update_learner_details(learner_id, field_values, editor_name):
    conn = get_db_connection()
    old_learner = conn.execute("SELECT * FROM learners WHERE id = ?", (learner_id,)).fetchone()
    if not old_learner:
        conn.close()
        return False, "Learner not found."
    
    now = datetime.datetime.now()
    date_str = now.strftime('%Y-%m-%d')
    time_str = now.strftime('%H:%M:%S')
    
    # Update learner
    set_clauses = []
    params = []
    for field, new_val in field_values.items():
        set_clauses.append(f"{field} = ?")
        params.append(new_val)
        
        # Log to audit trail if value changed
        old_val = str(old_learner[field]) if old_learner[field] is not None else "Not Provided"
        if str(new_val) != old_val:
            conn.execute('''
                INSERT INTO audit_trail (learner_id, learner_name, field_changed, old_value, new_value, changed_by, date, time)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            ''', (learner_id, old_learner['name'], field, old_val, str(new_val), editor_name, date_str, time_str))
            
    params.append(learner_id)
    conn.execute(f"UPDATE learners SET {', '.join(set_clauses)} WHERE id = ?", params)
    
    conn.commit()
    conn.close()
    return True, "Learner updated successfully."

def get_audit_trail():
    conn = get_db_connection()
    audit = conn.execute("SELECT * FROM audit_trail ORDER BY timestamp DESC").fetchall()
    conn.close()
    return [dict(a) for a in audit]

# Initialize Session State
if 'current_teacher' not in st.session_state:
    st.session_state.current_teacher = None
if 'another_teacher_name' not in st.session_state:
    st.session_state.another_teacher_name = ""

# Sidebar - Teacher Selection and Navigation
with st.sidebar:
    st.markdown("<h2 style='color:#1F4E78; margin-bottom:0;'>⚙️ NDL CONTROL PANEL</h2>", unsafe_allow_html=True)
    st.markdown("<hr style='margin:10px 0;'>", unsafe_allow_html=True)
    
    # Teacher Selector
    teachers = [
        "Select a Teacher",
        "Mr. Gideon",
        "Ms. Morgan",
        "Ms. Natasha",
        "Ms. Likando",
        "Mr. Saka (ISAZ Chairperson)",
        "Another Teacher"
    ]
    
    selected_teacher_raw = st.selectbox(
        "Who is using the app?",
        teachers,
        index=teachers.index(st.session_state.current_teacher) if st.session_state.current_teacher in teachers else 0
    )
    
    # Handle another teacher input
    if selected_teacher_raw == "Another Teacher":
        another_name = st.text_input("Enter your name:", value=st.session_state.another_teacher_name)
        if another_name:
            st.session_state.current_teacher = f"Another Teacher: {another_name}"
            st.session_state.another_teacher_name = another_name
        else:
            st.session_state.current_teacher = "Another Teacher"
    elif selected_teacher_raw != "Select a Teacher":
        st.session_state.current_teacher = selected_teacher_raw
    else:
        st.session_state.current_teacher = None
        
    if st.session_state.current_teacher:
        st.info(f"Active User: **{st.session_state.current_teacher}**")
    else:
        st.warning("Please select your name to access personalized features.")
        
    st.markdown("<hr style='margin:15px 0;'>", unsafe_allow_html=True)
    
    # Navigation Menu
    st.markdown("### 🗺️ NAVIGATION")
    nav_options = [
        "🏠 Dashboard",
        "👥 My Assigned Learners",
        "⚡ Quick Roll Call",
        "👥 All Learners — Status",
        "📋 Learner Profiles & Info",
        "🔴 View All Updates Feed",
        "🕘 Historical Event Log",
        "⚙️ Settings & Administration"
    ]
    
    # Let's restrict Roll Call and My Learners if no teacher selected
    nav_selection = st.radio(
        "Go to:",
        nav_options,
        index=0
    )
    
    st.markdown("<hr style='margin:15px 0;'>", unsafe_allow_html=True)
    
    # Export Section
    st.markdown("### 📤 SHARE REPORT")
    st.write("Generate and download the live data report.")
    
    # Generate files on-the-fly for download buttons
    if st.button("🔄 Refresh Export Files"):
        with st.spinner("Generating files..."):
            generate_excel_report('/workspace/scratch/NDL_Status_Report.xlsx')
            generate_pdf_report('/workspace/scratch/NDL_Status_Report.pdf')
        st.success("Files ready for download!")
        
    # Provide download links if they exist, or generate them
    if not os.path.exists('/workspace/scratch/NDL_Status_Report.xlsx'):
        generate_excel_report('/workspace/scratch/NDL_Status_Report.xlsx')
    if not os.path.exists('/workspace/scratch/NDL_Status_Report.pdf'):
        generate_pdf_report('/workspace/scratch/NDL_Status_Report.pdf')
        
    with open('/workspace/scratch/NDL_Status_Report.xlsx', 'rb') as f_excel:
        st.download_button(
            label="📊 Download Excel Worksheet",
            data=f_excel,
            file_name="NDL_Learner_Monitoring_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        
    with open('/workspace/scratch/NDL_Status_Report.pdf', 'rb') as f_pdf:
        st.download_button(
            label="📄 Download PDF Status Report",
            data=f_pdf,
            file_name="NDL_Learner_Status_Report.pdf",
            mime="application/pdf",
            use_container_width=True
        )

# Main Application Logic
st.markdown("<div class='main-title'>NDL Real-Time Learner Monitoring</div>", unsafe_allow_html=True)
st.markdown("<div class='sub-title'>National Debate League Delegation Attendance and Safety App</div>", unsafe_allow_html=True)

# Navigation dispatcher
if nav_selection == "🏠 Dashboard":
    # -------------------------------------------------------------
    # DASHBOARD
    # -------------------------------------------------------------
    if st.session_state.current_teacher:
        st.subheader(f"Welcome, {st.session_state.current_teacher}!")
    else:
        st.info("Welcome! Please select your name in the sidebar to access personalized features.")
        
    # Summary Cards
    learners = load_learners()
    status_map = get_latest_status_map()
    
    total_l = len(learners)
    present_l = sum(1 for status in status_map.values() if status['status'] == 'Present')
    absent_l = sum(1 for status in status_map.values() if status['status'] == 'Absent')
    no_update_l = total_l - present_l - absent_l
    
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("Total Learners", total_l)
    with col2:
        st.metric("Present Now", present_l)
    with col3:
        st.metric("Absent Now", absent_l, delta_color="inverse")
    with col4:
        st.metric("No Update Yet", no_update_l)
        
    st.markdown("<div class='section-header'>⚡ QUICK ACTIONS</div>", unsafe_allow_html=True)
    
    col_act1, col_act2, col_act3 = st.columns(3)
    with col_act1:
        if st.session_state.current_teacher:
            st.markdown("#### **Roll Call**")
            st.write("Start a quick event-based roll call for your assigned learners.")
            st.info("💡 Pro-Tip: Use 'Mark All Present' to record in 3 seconds!")
        else:
            st.markdown("#### **Roll Call**")
            st.write("Please select your name first on the left sidebar to record updates.")
    with col_act2:
        st.markdown("#### **All Statuses**")
        st.write("See real-time attendance and events for all 16 delegation learners.")
        st.write(f"🟢 Present: **{present_l}** | 🔴 Absent: **{absent_l}**")
    with col_act3:
        st.markdown("#### **Print & Export**")
        st.write("Generate clean Excel sheets or formatted PDF safety summaries for other chaperones.")
        
    # Recent Updates Feed (Dashboard)
    st.markdown("<div class='section-header'>🔴 RECENT LIVE UPDATES FROM TEAM</div>", unsafe_allow_html=True)
    recent_updates = load_updates()[:5]
    if not recent_updates:
        st.write("No updates have been recorded yet. Select 'Quick Roll Call' to log the first event!")
    else:
        for u in recent_updates:
            color = "🟢" if u['status'] == "Present" else "🔴"
            note_str = f" (*Note: {u['note']}*)" if u['note'] else ""
            st.markdown(f"**{u['date']} {u['time']}** — {color} **{u['learner_name']}** marked **{u['status']}** for **{u['event']}** by **{u['recorded_by']}**{note_str}")

elif nav_selection == "👥 My Assigned Learners":
    # -------------------------------------------------------------
    # MY ASSIGNED LEARNERS
    # -------------------------------------------------------------
    if not st.session_state.current_teacher:
        st.warning("⚠️ Please select your name in the sidebar to view your assigned learners.")
    else:
        st.markdown(f"<div class='section-header'>👥 LEARNERS ASSIGNED TO {st.session_state.current_teacher.upper()}</div>", unsafe_allow_html=True)
        
        # Filter assigned learners
        learners = load_learners()
        assigned = [l for l in learners if l['assigned_teacher'] == st.session_state.current_teacher]
        
        if not assigned:
            if "Mr. Saka" in st.session_state.current_teacher:
                st.info("Chairperson Mr. Saka has no learners directly assigned. You can monitor and view all updates on the other tabs.")
            else:
                st.info("You currently have no assigned learners. Go to Settings to configure your learner roster.")
        else:
            status_map = get_latest_status_map()
            
            for l in assigned:
                st_data = status_map.get(l['id'], {})
                status_text = st_data.get('status')
                event_text = st_data.get('event')
                up_time = f"{st_data.get('date')} {st_data.get('time')}" if st_data.get('date') else "Never"
                
                if not status_text:
                    badge_html = "<span class='badge badge-none'>⚪ No Update Yet</span>"
                elif status_text == "Present":
                    badge_html = f"<span class='badge badge-present'>🟢 Present ({event_text})</span>"
                else:
                    badge_html = f"<span class='badge badge-absent'>🔴 Absent ({event_text})</span>"
                    
                with st.container():
                    col_l1, col_l2 = st.columns([3, 1])
                    with col_l1:
                        st.markdown(f"### **{l['name']}**")
                        st.markdown(f"**Room Number:** {l['room']} | **Class/Grade:** {l['class_grade']}")
                        st.markdown(f"**Latest Status:** {badge_html} *(Updated: {up_time})*", unsafe_allow_html=True)
                        if st_data.get('note'):
                            st.caption(f"✍️ **Note:** {st_data.get('note')}")
                    with col_l2:
                        st.markdown("<br>", unsafe_allow_html=True)
                        # Quick status update button
                        if st.button(f"✏️ Quick Update Status", key=f"quick_up_{l['id']}"):
                            st.session_state[f"show_quick_panel_{l['id']}"] = True
                            
                    # Toggleable quick update form
                    if st.session_state.get(f"show_quick_panel_{l['id']}", False):
                        with st.form(key=f"quick_form_{l['id']}"):
                            st.write(f"**Update Status for {l['name']}**")
                            q_event = st.selectbox("Select Event:", [
                                "🍳 Breakfast", "🏨 At the Hotel", "🚌 Boarding Bus", "📍 Arrived at Venue",
                                "🗣️ Debate", "🍽️ Lunch", "🍽️ Dinner", "🛏️ Bedtime", "✈️ Airport / Flight",
                                "🏠 Picked by Parents", "Other"
                            ])
                            if q_event == "Other":
                                q_event = st.text_input("Type Event Name:")
                            q_status = st.radio("Status:", ["Present", "Absent"], horizontal=True)
                            q_note = st.text_input("Optional Note:")
                            
                            col_f1, col_f2 = st.columns(2)
                            with col_f1:
                                if st.form_submit_button("Save Status Update"):
                                    res, msg = save_roll_call(
                                        st.session_state.current_teacher,
                                        q_event,
                                        {l['id']: q_status},
                                        {l['id']: q_note}
                                    )
                                    if res:
                                        st.success(msg)
                                        st.session_state[f"show_quick_panel_{l['id']}"] = False
                                        st.rerun()
                                    else:
                                        st.error(msg)
                            with col_f2:
                                if st.form_submit_button("Cancel"):
                                    st.session_state[f"show_quick_panel_{l['id']}"] = False
                                    st.rerun()
                    st.markdown("---")

elif nav_selection == "⚡ Quick Roll Call":
    # -------------------------------------------------------------
    # QUICK ROLL CALL
    # -------------------------------------------------------------
    if not st.session_state.current_teacher:
        st.warning("⚠️ Please select your name in the sidebar to perform a Quick Roll Call.")
    else:
        st.markdown("<div class='section-header'>⚡ QUICK ROLL CALL</div>", unsafe_allow_html=True)
        st.write("Record attendance for all your assigned learners in seconds.")
        
        # Step 1: Select Event
        st.markdown("### **Step 1: Select Event**")
        events_list = [
            "🍳 Breakfast",
            "🏨 At the Hotel",
            "🚌 Boarding Bus",
            "📍 Arrived at Venue",
            "🗣️ Debate",
            "🍽️ Lunch",
            "🍽️ Dinner",
            "🛏️ Bedtime",
            "✈️ Airport / Flight",
            "🏠 Picked by Parents",
            "✏️ Other Event"
        ]
        
        selected_event = st.radio("What event are you tracking?", events_list, horizontal=True)
        
        event_name = selected_event
        if selected_event == "✏️ Other Event":
            event_name = st.text_input("Please specify the custom event name:")
            
        # Get assigned learners
        learners = load_learners()
        assigned = [l for l in learners if l['assigned_teacher'] == st.session_state.current_teacher]
        
        if not assigned:
            if "Mr. Saka" in st.session_state.current_teacher:
                st.info("Chairperson Mr. Saka has no learners directly assigned. You can view all roll calls or log updates via Settings/profiles.")
            else:
                st.info("You have no assigned learners. Go to Settings/Administration to assign learners to yourself first!")
        elif not event_name:
            st.warning("Please specify the custom event name.")
        else:
            # Step 2: Mark Present/Absent
            st.markdown("<hr style='margin:15px 0;'>", unsafe_allow_html=True)
            st.markdown("### **Step 2: Mark Attendance**")
            
            # Helper: Mark All Present
            col_hp1, col_hp2 = st.columns([1, 4])
            with col_hp1:
                if st.button("✓ Mark All Present", type="primary", use_container_width=True):
                    for l in assigned:
                        st.session_state[f"roll_{l['id']}"] = "Present"
                    st.toast("Marked all as Present! Adjust absentees below as needed.")
                    
            st.write("") # Spacer
            
            attendance_dict = {}
            notes_dict = {}
            
            for l in assigned:
                # Key initialization if not in session state
                state_key = f"roll_{l['id']}"
                if state_key not in st.session_state:
                    st.session_state[state_key] = "Present"
                    
                st.markdown(f"<div class='roll-call-card'>", unsafe_allow_html=True)
                col_rc1, col_rc2, col_rc3 = st.columns([3, 2, 3])
                with col_rc1:
                    st.markdown(f"👤 **{l['name']}**")
                    st.caption(f"Room: {l['room']}")
                with col_rc2:
                    st.session_state[state_key] = st.radio(
                        "Status", 
                        ["Present", "Absent"], 
                        key=f"radio_{l['id']}", 
                        index=0 if st.session_state[state_key] == "Present" else 1,
                        horizontal=True,
                        label_visibility="collapsed"
                    )
                    attendance_dict[l['id']] = st.session_state[state_key]
                with col_rc3:
                    notes_dict[l['id']] = st.text_input(
                        "Optional note (e.g. 'unwell')", 
                        key=f"note_{l['id']}", 
                        placeholder="Optional Note", 
                        label_visibility="collapsed"
                    )
                st.markdown("</div>", unsafe_allow_html=True)
                
            # Step 3: Save Roll Call
            st.markdown("<hr style='margin:15px 0;'>", unsafe_allow_html=True)
            st.markdown("### **Step 3: Save Roll Call**")
            
            if st.button("💾 SAVE ROLL CALL", type="primary", use_container_width=True):
                res, msg = save_roll_call(
                    st.session_state.current_teacher,
                    event_name,
                    attendance_dict,
                    notes_dict
                )
                if res:
                    st.success(f"🎉 {msg}")
                    # Clear roll call session states after successful save
                    for l in assigned:
                        if f"roll_{l['id']}" in st.session_state:
                            del st.session_state[f"roll_{l['id']}"]
                    st.rerun()
                else:
                    st.error(f"❌ {msg}")

elif nav_selection == "👥 All Learners — Status":
    # -------------------------------------------------------------
    # ALL LEARNERS STATUS
    # -------------------------------------------------------------
    st.markdown("<div class='section-header'>👥 ALL LEARNERS CURRENT STATUS</div>", unsafe_allow_html=True)
    st.write("Live status overview of all 16 delegation learners.")
    
    status_map = get_latest_status_map()
    learners = load_learners()
    
    # Simple search and filter
    col_sf1, col_sf2 = st.columns(2)
    with col_sf1:
        search_query = st.text_input("🔎 Search Learner by Name:", "")
    with col_sf2:
        filter_teacher = st.selectbox("Filter by Assigned Chaperone:", ["All"] + sorted(list(set(l['assigned_teacher'] for l in learners))))
        
    filtered_learners = learners
    if search_query:
        filtered_learners = [l for l in filtered_learners if search_query.lower() in l['name'].lower()]
    if filter_teacher != "All":
        filtered_learners = [l for l in filtered_learners if l['assigned_teacher'] == filter_teacher]
        
    # Render Status Cards
    st.markdown("<br>", unsafe_allow_html=True)
    
    # Table headers
    col_th1, col_th2, col_th3, col_th4, col_th5, col_th6 = st.columns([3, 2, 1, 2, 2, 2])
    with col_th1: st.markdown("**Learner Name**")
    with col_th2: st.markdown("**Assigned Teacher**")
    with col_th3: st.markdown("**Room**")
    with col_th4: st.markdown("**Latest Status**")
    with col_th5: st.markdown("**Event**")
    with col_th6: st.markdown("**Updated At**")
    st.markdown("<hr style='margin:5px 0;'>", unsafe_allow_html=True)
    
    for l in filtered_learners:
        st_data = status_map.get(l['id'], {})
        status_text = st_data.get('status')
        event_text = st_data.get('event', 'N/A')
        up_time = f"{st_data.get('date')} {st_data.get('time')}" if st_data.get('date') else "No update yet"
        recorded_by = st_data.get('recorded_by', 'N/A')
        
        if not status_text:
            badge_html = "<span class='badge badge-none'>⚪ No Update Yet</span>"
        elif status_text == "Present":
            badge_html = f"<span class='badge badge-present'>🟢 Present</span>"
        else:
            badge_html = f"<span class='badge badge-absent'>🔴 Absent</span>"
            
        col_td1, col_td2, col_td3, col_td4, col_td5, col_td6 = st.columns([3, 2, 1, 2, 2, 2])
        with col_td1: 
            st.markdown(f"**{l['name']}**")
            if st_data.get('note'):
                st.caption(f"✍️ Note: {st_data.get('note')}")
        with col_td2: st.write(l['assigned_teacher'])
        with col_td3: st.write(l['room'])
        with col_td4: st.markdown(badge_html, unsafe_allow_html=True)
        with col_td5: st.write(event_text)
        with col_td6: st.write(up_time)
        st.markdown("<hr style='margin:3px 0; border:0; border-top:1px solid #ECECEC;'>", unsafe_allow_html=True)

elif nav_selection == "📋 Learner Profiles & Info":
    # -------------------------------------------------------------
    # LEARNER PROFILES & INFO
    # -------------------------------------------------------------
    st.markdown("<div class='section-header'>📋 DELEGATION LEARNER DIRECTORY & PROFILES</div>", unsafe_allow_html=True)
    st.write("Access full safety files, medical alerts, emergency contacts, and history logs.")
    
    # Access controls warning
    st.info("🔒 SENSITIVE INFORMATION DETECTED: Parent phone numbers and medical info are protected. Toggle authorization below to view.")
    authorized_view = st.checkbox("🔑 I am an Authorized Chaperone / Chairperson (Show sensitive details)")
    
    learners = load_learners()
    search_prof = st.text_input("🔎 Search Profile Directory (Name or Teacher):", "")
    
    filtered_p = learners
    if search_prof:
        filtered_p = [l for l in filtered_p if search_prof.lower() in l['name'].lower() or search_prof.lower() in l['assigned_teacher'].lower()]
        
    for l in filtered_p:
        with st.expander(f"👤 {l['name']} — Room {l['room']} ({l['assigned_teacher']})"):
            col_p1, col_p2 = st.columns(2)
            with col_p1:
                st.markdown("#### **Personal Details**")
                st.markdown(f"**Full Name:** {l['name']}")
                st.markdown(f"**Class/Grade:** {l['class_grade']}")
                st.markdown(f"**Assigned Teacher:** {l['assigned_teacher']}")
                st.markdown(f"**Room Number:** {l['room']}")
            with col_p2:
                st.markdown("#### **Emergency & Medical Contacts**")
                if authorized_view:
                    st.markdown(f"**Parent/Guardian:** {l['parent_name']}")
                    st.markdown(f"**Parent Phone/WhatsApp:** `{l['parent_phone']}`")
                    st.markdown(f"**Emergency Contact Phone:** `{l['emergency_contact']}`")
                    
                    # Highlight medical details in a red box if not empty
                    meds = l['medical_conditions']
                    if meds and meds != "None" and meds != "Not Provided":
                        st.markdown(f"<div style='background-color:#FDE8E8; padding:10px; border-radius:5px; border-left:4px solid #E02424; margin-bottom:10px;'>🚨 <b>MEDICAL ALLERGEN ALERT:</b> {meds}</div>", unsafe_allow_html=True)
                    else:
                        st.markdown(f"**Medical Conditions:** {meds}")
                        
                    other = l['other_info']
                    if other and other != "None" and other != "Not Provided":
                        st.markdown(f"💡 **Other Info:** {other}")
                else:
                    st.warning("⚠️ Details hidden. Check the authorization box above to view medical / emergency details.")
            
            # Show historical logs for this learner specifically
            st.markdown("---")
            st.markdown("#### **Historical Safety Logs**")
            l_updates = load_updates({'learner_id': l['id']})
            if not l_updates:
                st.write("No updates have been recorded yet for this learner.")
            else:
                for lu in l_updates:
                    c_badge = "🟢" if lu['status'] == "Present" else "🔴"
                    note_str = f" (*Note: {lu['note']}*)" if lu['note'] else ""
                    st.write(f"{lu['date']} {lu['time']} — {c_badge} Marked **{lu['status']}** at **{lu['event']}** (by {lu['recorded_by']}){note_str}")

elif nav_selection == "🔴 View All Updates Feed":
    # -------------------------------------------------------------
    # VIEW ALL UPDATES
    # -------------------------------------------------------------
    st.markdown("<div class='section-header'>🔴 REAL-TIME UPDATES FEED</div>", unsafe_allow_html=True)
    st.write("Chronological log of safety updates across all chaperones.")
    
    # Filtering tools
    col_uf1, col_uf2, col_uf3 = st.columns(3)
    with col_uf1:
        uf_event = st.selectbox("Filter by Event:", ["All", "Breakfast", "At the Hotel", "Boarding Bus", "Arrived at Venue", "Debate", "Lunch", "Dinner", "Bedtime", "Airport / Flight", "Picked by Parents"])
    with col_uf2:
        uf_teacher = st.selectbox("Recorded By Chaperone:", ["All", "Mr. Gideon", "Ms. Morgan", "Ms. Natasha", "Ms. Likando", "Mr. Saka"])
    with col_uf3:
        uf_status = st.selectbox("Status Filter:", ["All", "Present", "Absent"])
        
    filters = {}
    if uf_event != "All":
        # Match with wildcard or exact
        filters['event'] = uf_event
    if uf_teacher != "All":
        filters['recorded_by'] = uf_teacher
    if uf_status != "All":
        filters['status'] = uf_status
        
    updates = load_updates(filters)
    
    if not updates:
        st.write("No updates match the selected filters.")
    else:
        # Group by Date and Event for structured visualization as requested in Passage 25
        df_u = pd.DataFrame(updates)
        
        # Format the grouped log elegantly
        for (date_val, event_val), group in df_u.groupby(['date', 'event']):
            st.markdown(f"### 📍 {event_val.upper()} — {date_val}")
            
            for index, row in group.iterrows():
                badge_color = "🟢" if row['status'] == "Present" else "🔴"
                note_str = f" | Note: *{row['note']}*" if row['note'] else ""
                st.markdown(f" {badge_color} **{row['learner_name']}** — Marked **{row['status']}** by **{row['recorded_by']}** at {row['time']}{note_str}")
            st.markdown("<hr style='margin:10px 0;'>", unsafe_allow_html=True)

elif nav_selection == "🕘 Historical Event Log":
    # -------------------------------------------------------------
    # HISTORY
    # -------------------------------------------------------------
    st.markdown("<div class='section-header'>🕘 SYSTEM CHRONOLOGICAL HISTORY LOG</div>", unsafe_allow_html=True)
    st.write("Full historical audit log of the shared SQLite cloud database.")
    
    updates = load_updates()
    
    if not updates:
        st.info("The history database is currently empty. Roll calls recorded in the app will appear here permanently.")
    else:
        df_hist = pd.DataFrame(updates)
        st.dataframe(
            df_hist[['timestamp', 'learner_name', 'assigned_teacher', 'recorded_by', 'event', 'status', 'note', 'date', 'time']],
            column_config={
                'timestamp': 'System Timestamp',
                'learner_name': 'Learner Name',
                'assigned_teacher': 'Chaperone',
                'recorded_by': 'Recorded By',
                'event': 'Event',
                'status': 'Status',
                'note': 'Optional Note'
            },
            use_container_width=True,
            hide_index=True
        )

elif nav_selection == "⚙️ Settings & Administration":
    # -------------------------------------------------------------
    # SETTINGS / ADMINISTRATION
    # -------------------------------------------------------------
    st.markdown("<div class='section-header'>⚙️ CHAPERONE ADMINISTRATION PANEL</div>", unsafe_allow_html=True)
    st.write("Configure rosters, edit placeholders, update rooms, and review system audit trails.")
    
    tab_settings, tab_placeholders, tab_audit = st.tabs(["✏️ Edit Learners & Rooms", "🧩 Configure Placeholder Names", "📋 System Audit Trail"])
    
    with tab_settings:
        st.write("### **Update Learner Profiles**")
        learners = load_learners()
        sel_l_name = st.selectbox("Select Learner to Edit:", [l['name'] for l in learners])
        
        sel_l = next(l for l in learners if l['name'] == sel_l_name)
        
        if sel_l:
            with st.form(key=f"edit_learner_form_{sel_l['id']}"):
                st.write(f"Editing: **{sel_l['name']}**")
                
                new_room = st.text_input("Room Number:", value=sel_l['room'] if sel_l['room'] else "")
                new_grade = st.text_input("Class/Grade:", value=sel_l['class_grade'] if sel_l['class_grade'] else "")
                
                # Chaperone selector
                chaperones = ["Mr. Gideon", "Ms. Morgan", "Ms. Natasha", "Ms. Likando", "Mr. Saka", "Another Teacher"]
                current_chap_idx = chaperones.index(sel_l['assigned_teacher']) if sel_l['assigned_teacher'] in chaperones else 0
                new_teacher = st.selectbox("Assigned Chaperone:", chaperones, index=current_chap_idx)
                
                # Sensitive info
                new_parent = st.text_input("Parent/Guardian Full Name:", value=sel_l['parent_name'] if sel_l['parent_name'] else "")
                new_phone = st.text_input("Parent Phone:", value=sel_l['parent_phone'] if sel_l['parent_phone'] else "")
                new_emergency = st.text_input("Emergency Phone:", value=sel_l['emergency_contact'] if sel_l['emergency_contact'] else "")
                new_medical = st.text_area("Medical Conditions or Allergies:", value=sel_l['medical_conditions'] if sel_l['medical_conditions'] else "")
                new_other = st.text_area("Other Important Chaperone Info:", value=sel_l['other_info'] if sel_l['other_info'] else "")
                
                if st.form_submit_button("💾 Save Profile Changes"):
                    editor = st.session_state.current_teacher if st.session_state.current_teacher else "System Admin"
                    fields = {
                        'room': new_room if new_room else "Not Assigned",
                        'class_grade': new_grade if new_grade else "Not Assigned",
                        'assigned_teacher': new_teacher,
                        'parent_name': new_parent if new_parent else "Not Provided",
                        'parent_phone': new_phone if new_phone else "Not Provided",
                        'emergency_contact': new_emergency if new_emergency else "Not Provided",
                        'medical_conditions': new_medical if new_medical else "None",
                        'other_info': new_other if new_other else "None"
                    }
                    res, msg = update_learner_details(sel_l['id'], fields, editor)
                    if res:
                        st.success(f"🎉 {msg}")
                        st.rerun()
                    else:
                        st.error(f"❌ {msg}")
                        
    with tab_placeholders:
        st.write("### **Configure Temporary Placeholders**")
        st.write("Assign names to temporary placeholders: **Learner C**, **Learner D**, and **Learner E**.")
        
        placeholders = [l for l in learners if "Learner " in l['name']]
        if not placeholders:
            st.success("All placeholders have been renamed!")
        else:
            for p in placeholders:
                with st.form(key=f"rename_form_{p['id']}"):
                    st.write(f"Configure: **{p['name']}** (Assigned to: {p['assigned_teacher']})")
                    actual_name = st.text_input("Enter Actual Name:", placeholder="Firstname Lastname")
                    if st.form_submit_button(f"Save Name for {p['name']}"):
                        if not actual_name:
                            st.error("Please enter a valid name.")
                        else:
                            editor = st.session_state.current_teacher if st.session_state.current_teacher else "System Admin"
                            res, msg = update_learner_details(p['id'], {'name': actual_name, 'trip_info': 'Confirmed'}, editor)
                            if res:
                                st.success(f"🎉 Successfully renamed {p['name']} to {actual_name}!")
                                st.rerun()
                            else:
                                st.error(msg)
                                
    with tab_audit:
        st.write("### **System Activity Audit Trail**")
        st.write("Tracks all changes to learner room allocations, chaperone assignments, and personal records.")
        
        audit_trail = get_audit_trail()
        if not audit_trail:
            st.info("No profile modifications have been logged yet.")
        else:
            df_aud = pd.DataFrame(audit_trail)
            st.dataframe(
                df_aud[['timestamp', 'learner_name', 'field_changed', 'old_value', 'new_value', 'changed_by']],
                column_config={
                    'timestamp': 'Timestamp',
                    'learner_name': 'Learner Name',
                    'field_changed': 'Field Edited',
                    'old_value': 'Previous Value',
                    'new_value': 'New Value',
                    'changed_by': 'Modified By'
                },
                use_container_width=True,
                hide_index=True
            )
