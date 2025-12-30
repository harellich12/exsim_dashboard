"""
ExSim War Room - Bulk Upload Tab
Multi-file uploader for batch data import.
Expects exact filenames and shows summary panel before proceeding.
"""

import streamlit as st
import pandas as pd
from utils.state_manager import set_state
from utils.data_loader import (
    load_market_report, load_workers_balance, load_raw_materials,
    load_finished_goods, load_balance_statements, load_esg_report,
    load_production_data
)

# Expected filenames and their mappings
EXPECTED_FILES = {
    'market-report.xlsx': {
        'state_key': 'market_data',
        'loader': load_market_report,
        'tab': 'CMO (Marketing)',
        'description': 'Market share, segment data, competitive info'
    },
    'workers_balance_overtime.xlsx': {
        'state_key': 'workers_data',
        'loader': load_workers_balance,
        'tab': 'CPO (People)',
        'description': 'Worker counts, salaries, overtime by zone'
    },
    'raw_materials.xlsx': {
        'state_key': 'materials_data',
        'loader': load_raw_materials,
        'tab': 'Purchasing',
        'description': 'Raw materials inventory, costs'
    },
    'finished_goods_inventory.xlsx': {
        'state_key': 'finished_goods_data',
        'loader': load_finished_goods,
        'tab': 'Logistics',
        'description': 'Finished goods stock, warehouse capacity'
    },
    'production.xlsx': {
        'state_key': 'production_data',
        'loader': load_production_data,
        'tab': 'Production',
        'description': 'Machine capacity, production output'
    },
    'ESG.xlsx': {
        'state_key': 'esg_data',
        'loader': load_esg_report,
        'tab': 'ESG',
        'description': 'Emissions, energy consumption, tax rates'
    },
    'results_and_balance_statements.xlsx': {
        'state_key': 'balance_data',
        'loader': load_balance_statements,
        'tab': 'CFO (Finance)',
        'description': 'P&L, balance sheet, cash positions'
    }
}


def reset_tab_states():
    """Reset all tab initialization flags to force reload with new data."""
    for key in ['cmo_initialized', 'production_initialized', 'purchasing_initialized',
                'logistics_initialized', 'cpo_initialized', 'esg_initialized', 'cfo_initialized']:
        if key in st.session_state:
            del st.session_state[key]


def render_bulk_upload():
    """Render the Bulk Upload tab with multi-file uploader and summary panel."""
    st.header("📦 Bulk Upload - Data Import Center")
    
    st.markdown("""
    Upload all your ExSim Excel reports at once. The system expects **exact filenames** 
    as exported from ExSim. After upload, all dashboards will be automatically populated.
    """)
    
    # Show expected files
    with st.expander("📋 Expected File List", expanded=False):
        for filename, info in EXPECTED_FILES.items():
            st.markdown(f"- **`{filename}`** → {info['tab']}: {info['description']}")
    
    st.markdown("---")
    
    # Multi-file uploader
    uploaded_files = st.file_uploader(
        "Select all Excel files to upload",
        type=['xlsx', 'xls'],
        accept_multiple_files=True,
        key='bulk_uploader'
    )
    
    if uploaded_files:
        st.markdown("---")
        st.subheader("📊 Upload Summary")
        
        # Track results
        results = {
            'loaded': [],
            'skipped': [],
            'errors': []
        }
        
        # Process each file
        for file in uploaded_files:
            filename = file.name
            
            if filename in EXPECTED_FILES:
                config = EXPECTED_FILES[filename]
                try:
                    # Load data using appropriate loader
                    data = config['loader'](file)
                    
                    if data:
                        set_state(config['state_key'], data)
                        results['loaded'].append({
                            'file': filename,
                            'tab': config['tab'],
                            'status': '✅ Loaded'
                        })
                    else:
                        results['errors'].append({
                            'file': filename,
                            'error': 'Empty or invalid data'
                        })
                except Exception as e:
                    results['errors'].append({
                        'file': filename,
                        'error': str(e)
                    })
            else:
                results['skipped'].append({
                    'file': filename,
                    'reason': 'Unrecognized filename'
                })
        
        # Display results in columns
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.metric("✅ Loaded", len(results['loaded']))
        with col2:
            st.metric("⏭️ Skipped", len(results['skipped']))
        with col3:
            st.metric("❌ Errors", len(results['errors']))
        
        # Detailed results table
        if results['loaded']:
            st.success(f"**{len(results['loaded'])} files loaded successfully!**")
            
            loaded_df = pd.DataFrame(results['loaded'])
            st.dataframe(loaded_df, use_container_width=True, hide_index=True)
        
        if results['skipped']:
            st.warning("**Skipped files (unrecognized names):**")
            for item in results['skipped']:
                st.markdown(f"- `{item['file']}` - {item['reason']}")
        
        if results['errors']:
            st.error("**Errors:**")
            for item in results['errors']:
                st.markdown(f"- `{item['file']}` - {item['error']}")
        
        # Proceed button
        st.markdown("---")
        
        if results['loaded']:
            if st.button("🚀 Apply Data to All Tabs", type="primary", use_container_width=True):
                # Reset tab initialization to force reload with new data
                reset_tab_states()
                st.success("✅ Data applied! Navigate to any tab to see the populated data.")
                st.balloons()
        
        # Missing files warning
        loaded_filenames = [r['file'] for r in results['loaded']]
        missing = [f for f in EXPECTED_FILES.keys() if f not in loaded_filenames]
        
        if missing:
            st.info(f"**Optional:** {len(missing)} files not uploaded: {', '.join(missing)}")
    
    else:
        # No files uploaded yet - show drag & drop zone
        st.markdown("""
        <div style="
            border: 2px dashed #ccc;
            border-radius: 10px;
            padding: 40px;
            text-align: center;
            background: #f8f9fa;
            margin: 20px 0;
        ">
            <h3>📁 Drag & Drop Files Here</h3>
            <p>or click "Browse files" above</p>
        </div>
        """, unsafe_allow_html=True)
        
        st.caption("Tip: You can select multiple files at once by holding Ctrl/Cmd while clicking.")
