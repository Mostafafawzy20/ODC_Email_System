import streamlit as st
import pandas as pd
from datetime import datetime
import time
from typing import Optional, Tuple, Any

from Email import (
    process_single_notification,
    process_bulk_notifications,
    send_to_all_hotels,
    load_and_prepare_data as load_email_data,
    EmailReport
)

def load_and_filter_data(missing_file, emails_file, exception_file=None, min_days_missing=None, max_days_missing=None):
    try:
        # Process missing files data
        raw_df = pd.read_excel(missing_file, sheet_name='Missing', skiprows=3)
        
        # Create missing_df with first 6 columns and assign names
        missing_df = raw_df.iloc[:, :6].copy()
        missing_df.columns = ['prop_code', 'hotel', 'file_name', 'business_date', 'dow', 'days_missing']
        
        # Force Days Missing to numeric, replacing any non-numeric values with NaN
        missing_df['days_missing'] = pd.to_numeric(missing_df['days_missing'], errors='coerce')
        
        # Remove any rows where days_missing is NaN
        missing_df = missing_df.dropna(subset=['days_missing'])
        
        # Convert to integer type
        missing_df['days_missing'] = missing_df['days_missing'].astype(int)
        
        # Process other columns
        missing_df['business_date'] = pd.to_datetime(missing_df['business_date'], errors='coerce')
        missing_df['prop_code'] = missing_df['prop_code'].astype(str)
        
        # Process emails data
        emails_df = pd.read_excel(emails_file)
        emails_df.columns = ['Hotels', 'Email']
        emails_df['Hotels'] = emails_df['Hotels'].astype(str)

        # Apply exception list filter if provided
        if exception_file is not None:
            try:
                exception_df = pd.read_excel(exception_file)
                potential_columns = exception_df.columns.tolist()
                found_column = None
                for col in potential_columns:
                    if 'prop' in col.lower() or 'code' in col.lower():
                        found_column = col
                        break
                
                if found_column:
                    exception_df[found_column] = exception_df[found_column].astype(str)
                    missing_df = missing_df[~missing_df['prop_code'].isin(exception_df[found_column])]
                else:
                    st.warning(f"Could not find property code column in exception file")
            except Exception as e:
                st.error(f"Exception file error: {str(e)}")
                return None, f"Error processing exception list: {str(e)}"

        # Apply days missing filter
        if min_days_missing is not None or max_days_missing is not None:
            original_count = len(missing_df)
            
            # Convert filter values to integers
            min_days = int(min_days_missing) if min_days_missing is not None else None
            max_days = int(max_days_missing) if max_days_missing is not None else None
            
            # Apply strict filtering
            if min_days == max_days:
                # Exact match
                missing_df = missing_df[missing_df['days_missing'] == min_days]
            else:
                # Range filtering
                if min_days is not None:
                    missing_df = missing_df[missing_df['days_missing'] >= min_days]
                if max_days is not None:
                    missing_df = missing_df[missing_df['days_missing'] <= max_days]
            
            filtered_count = len(missing_df)
            
            if filtered_count == 0:
                st.warning("No records match the filter criteria")

        # Create result object with processed data
        class ProcessedData:
            def __init__(self, missing, emails):
                self.missing = missing
                self.emails = emails

        return ProcessedData(missing_df, emails_df), None

    except Exception as e:
        import traceback
        st.error(f"Full error: {traceback.format_exc()}")
        return None, f"Error in data preparation: {str(e)}"

def render_email_section(data, additional_info: str = ""):
    st.markdown("### 📧 Email Notifications")
    
    # Get unique hotels from the filtered data
    filtered_hotels = data.missing['prop_code'].unique()
    st.info(f"Ready to send notifications to {len(filtered_hotels)} hotels")
    
    with st.expander("⚙️ Email Settings", expanded=False):
        # Add template selection at the top
        st.markdown("### 📝 Email Template")
        template_type = st.radio(
            "Select Email Template",
            options=["Standard Template (<28 days)", "Critical Template (>28 days)"],
            help="Choose which email template to use regardless of days missing"
        )
        
        st.markdown("---")
        
        # Add From and Bcc fields
        col1, col2 = st.columns(2)
        with col1:
            from_email = st.text_input(
                "From Email",
                value="",
                help="Enter the email address the notifications will be sent from"
            )
        with col2:
            bcc_email = st.text_input(
                "Bcc Email",
                value="",
                help="Enter email address(es) to be BCC'd (separate multiple emails with semicolons)"
            )
        
        col3, col4 = st.columns(2)
        with col3:
            batch_size = st.number_input(
                "Batch Size",
                min_value=1,
                max_value=20,
                value=10
            )
        with col4:
            delay = st.number_input(
                "Delay Between Emails (seconds)",
                min_value=1.0,
                max_value=10.0,
                value=2.0
            )
    
    # Pass the template choice to the email processing functions
    force_template = "critical" if template_type == "Critical Template (>28 days)" else "standard"
    
    email_tabs = st.tabs(["🎯 Single Hotel", "📨 Bulk Notifications", "🌐 Send to All"])
    
    # Single hotel notification tab
    with email_tabs[0]:
        col1, col2 = st.columns([3, 1])
        with col1:
            selected_hotel = st.selectbox(
                "Select Hotel",
                options=filtered_hotels,
                format_func=lambda x: f"{x} - {data.missing[data.missing['prop_code'] == x]['hotel'].iloc[0]}"
            )
        
        with col2:
            if st.button("📤 Send Email", key="single_email"):
                if not from_email.strip():
                    st.error("Please enter a From email address")
                else:
                    with st.spinner("Sending email..."):
                        result = process_single_notification(
                            missing_df=data.missing,
                            emails_df=data.emails,
                            hotel_code=selected_hotel,
                            additional_info=additional_info,
                            delay=delay,
                            from_email=from_email,
                            bcc_email=bcc_email if bcc_email.strip() else None,
                            force_template=force_template
                        )
                        if result.success:
                            st.success(result.message)
                        else:
                            st.error(result.message)
    
    # Bulk notifications tab
    with email_tabs[1]:
        selected_hotels = st.multiselect(
            "Select Hotels for Bulk Notification",
            options=filtered_hotels,
            format_func=lambda x: f"{x} - {data.missing[data.missing['prop_code'] == x]['hotel'].iloc[0]}"
        )
        
        if selected_hotels:
            if st.button("📤 Send Bulk Emails", key="bulk_email"):
                if not from_email.strip():
                    st.error("Please enter a From email address")
                else:
                    progress_bar = st.progress(0)
                    
                    def update_progress(progress):
                        progress_bar.progress(progress)
                    
                    with st.spinner("Processing bulk emails..."):
                        results, report = process_bulk_notifications(
                            missing_df=data.missing,
                            emails_df=data.emails,
                            hotel_codes=selected_hotels,
                            additional_info=additional_info,
                            batch_size=batch_size,
                            delay_between_emails=delay,
                            progress_callback=update_progress,
                            from_email=from_email,
                            bcc_email=bcc_email if bcc_email.strip() else None,
                            force_template=force_template
                        )
                        
                        if results:
                            st.success(f"Successfully processed {len(selected_hotels)} hotels")
                            summary = report.get_summary()
                            st.info(f"Summary: {summary['success']} succeeded, {summary['failed']} failed")
                        else:
                            st.error("Failed to process bulk emails")
    
    # Send to All tab
    with email_tabs[2]:
        st.warning(f"⚠️ This will send emails to all {len(filtered_hotels)} filtered hotels")
        
        if st.button("🚀 Send to All Hotels", key="send_all"):
            if not from_email.strip():
                st.error("Please enter a From email address")
            else:
                progress_bar = st.progress(0)
                
                def update_progress(progress):
                    progress_bar.progress(progress)
                
                with st.spinner("Processing all hotels..."):
                    results, report = send_to_all_hotels(
                        missing_df=data.missing,
                        emails_df=data.emails,
                        additional_info=additional_info,
                        batch_size=batch_size,
                        delay_between_emails=delay,
                        progress_callback=update_progress,
                        from_email=from_email,
                        bcc_email=bcc_email if bcc_email.strip() else None,
                        force_template=force_template
                    )
                    
                    if results:
                        st.success(f"Successfully processed all {len(filtered_hotels)} hotels")
                        summary = report.get_summary()
                        st.info(f"Summary: {summary['success']} succeeded, {summary['failed']} failed")
                    else:
                        st.error("Failed to process all hotels")

def render_filter_results(filtered_data: pd.DataFrame):
    """Render filtered results for user approval"""
    st.markdown("### 🔍 Filtered Hotels Preview")
    
    # Create a more readable preview table
    preview_df = filtered_data.copy()
    
    # Verify columns exist before sorting
    if 'days_missing' in preview_df.columns:
        preview_df = preview_df.sort_values(['prop_code', 'days_missing'], ascending=[True, False])
    
    # Group by hotel to show total files missing per hotel
    hotel_summary = preview_df.groupby(['prop_code', 'hotel']).agg({
        'file_name': 'count',
        'days_missing': 'max'
    }).reset_index()
    
    st.write("### Hotels to Receive Notifications")
    st.write(f"Total hotels that will receive notifications: {len(hotel_summary)}")
    
    # Show hotel summary
    st.dataframe(
        hotel_summary,
        column_config={
            'prop_code': "Property Code",
            'hotel': "Hotel Name",
            'file_name': "Number of Missing Files",
            'days_missing': "Max Days Missing"
        },
        height=400
    )
    
    # Add download button for filtered results
    csv = preview_df.to_csv(index=False)
    st.download_button(
        label="📥 Download Full Details",
        data=csv,
        file_name="filtered_hotels.csv",
        mime="text/csv"
    )
    
    # User approval
    st.markdown("### ✅ Confirm Selection")
    st.warning("⚠️ Only the hotels listed above will receive notifications. Please review carefully before proceeding.")
    proceed = st.button("✉️ Proceed to Email Notifications")
    
    return proceed

def main():
    st.set_page_config(
        page_title="ODC Email System",
        page_icon="📧",
        layout="wide",
    )
    
    st.title("📧 ODC Email System")
    st.markdown("Email Notification System for ODC Files")
    
    # Initialize session state for storing filtered data
    if 'filtered_data' not in st.session_state:
        st.session_state.filtered_data = None
    if 'show_email_section' not in st.session_state:
        st.session_state.show_email_section = False
    
    with st.sidebar:
        st.header("📝 Configuration")
        
        missing_file = st.file_uploader(
            "📊 Current Audit Report",
            type=['xlsx'],
            help="Upload the current audit report Excel file"
        )
        
        emails_file = st.file_uploader(
            "📧 Emails List",
            type=['xlsx'],
            help="Upload the email list Excel file"
        )
        
        exception_file = st.file_uploader(
            "⚠️ Exception List (Optional)",
            type=['xlsx'],
            help="Upload a list of property codes to exclude from processing"
        )
        
        st.markdown("---")
        st.markdown("### 🔍 Days Missing Filter")
        
        filter_type = st.radio(
            "Filter Type",
            options=["No Filter", "Exact Days", "Range of Days"],
            help="Choose how you want to filter by missing days"
        )
        
        min_days = None
        max_days = None
        
        if filter_type == "Exact Days":
            exact_days = st.selectbox(
                "Select Exact Days",
                options=[4, 10, 20, 28],
                help="Show only hotels with files missing for exactly this many days"
            )
            min_days = max_days = exact_days
            
        elif filter_type == "Range of Days":
            col1, col2 = st.columns(2)
            with col1:
                min_days = st.number_input(
                    "From Days",
                    min_value=0,
                    max_value=28,
                    value=4,
                    help="Minimum number of days missing"
                )
            with col2:
                max_days = st.number_input(
                    "To Days",
                    min_value=min_days,
                    max_value=28,
                    value=min(10, min_days),
                    help="Maximum number of days missing"
                )
        
        st.markdown("---")
        
        # Add Apply Filter button
        if st.button("🔍 Apply Filter", key="apply_filter"):
            if missing_file and emails_file:
                with st.spinner("🔄 Loading and processing data..."):
                    result, error = load_and_filter_data(
                        missing_file=missing_file,
                        emails_file=emails_file,
                        exception_file=exception_file,
                        min_days_missing=min_days,
                        max_days_missing=max_days
                    )
                    
                    if error:
                        st.error(f"❌ {error}")
                    else:
                        st.session_state.filtered_data = result
                        st.session_state.show_email_section = False  # Reset email section visibility
            else:
                st.error("Please upload the required files before applying the filter.")
        
        additional_info = st.text_area(
            "Additional Email Information",
            help="This text will be added to the email notifications",
            placeholder="Enter any additional information to include in the emails..."
        )

    # Main content area
    if st.session_state.filtered_data:
        if not st.session_state.show_email_section:
            # Show filter results and get user approval
            proceed = render_filter_results(st.session_state.filtered_data.missing)
            if proceed:
                st.session_state.show_email_section = True
                st.rerun()
        
        if st.session_state.show_email_section:
            # Remove the filter summary section and go straight to email section
            render_email_section(st.session_state.filtered_data, additional_info)
    else:
        st.info("Please upload the required files and apply the filter to get started.")

if __name__ == "__main__":
    main()
