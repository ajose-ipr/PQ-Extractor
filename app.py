import streamlit as st
import subprocess
import os
import sys

# Set page config
st.set_page_config(
    page_title="Harmonic Analysis Toolkit",
    page_icon="📊",
    layout="wide"
)

# Custom CSS for styling
st.markdown("""
    <style>
    .main {
        background-color: #f5f5f5;
    }
    .title {
        color: #2c3e50;
        text-align: center;
        font-size: 2.5em;
        margin-bottom: 30px;
    }
    .card {
        background-color: white;
        border-radius: 10px;
        padding: 20px;
        margin: 10px 0;
        box-shadow: 0 4px 8px rgba(0,0,0,0.1);
        transition: transform 0.3s;
    }
    .card:hover {
        transform: translateY(-5px);
    }
    .button {
        background-color: #3498db;
        color: white;
        border: none;
        padding: 10px 20px;
        text-align: center;
        text-decoration: none;
        display: inline-block;
        font-size: 16px;
        margin: 4px 2px;
        cursor: pointer;
        border-radius: 5px;
    }
    .button:hover {
        background-color: #2980b9;
    }
    </style>
    """, unsafe_allow_html=True)

# App title
st.markdown("<h1 class='title'>Harmonic Analysis Toolkit</h1>", unsafe_allow_html=True)

# Define the modules
modules = [
    {
        "name": "7-Day Summary Analyzer",
        "description": "Extract THD/TDD & compliance data from harmonic analysis reports",
        "script": "7-Day Summary Analyzer.py",
        "icon": "📅"
    },
    {
        "name": "Graph Extractor",
        "description": "Extract and save graphs from harmonic analysis reports",
        "script": "Graph Extractor.py",
        "icon": "📈"
    },
    {
        "name": "Harmonic Table Analyzer",
        "description": "Extract harmonic tables from analysis reports",
        "script": "Harmonic Table Analyzer.py",
        "icon": "📋"
    }
]

# Create columns for the cards
cols = st.columns(3)

for idx, module in enumerate(modules):
    with cols[idx]:
        st.markdown(f"""
        <div class="card">
            <h2>{module['icon']} {module['name']}</h2>
            <p>{module['description']}</p>
        </div>
        """, unsafe_allow_html=True)
        
        if st.button(f"Run {module['name']}", key=f"btn_{idx}"):
            # Check if the script exists
            if os.path.exists(module["script"]):
                try:
                    # Run the script using subprocess
                    process = subprocess.Popen(["streamlit", "run", module["script"]])
                    st.success(f"{module['name']} launched successfully!")
                except Exception as e:
                    st.error(f"Error launching {module['name']}: {str(e)}")
            else:
                st.error(f"Script file '{module['script']}' not found. Please ensure it's in the same directory.")

# Footer
st.markdown("---")
st.markdown("""
    <div style="text-align: center; color: #7f8c8d;">
        <p>Harmonic Analysis Toolkit v1.0</p>
    </div>
    """, unsafe_allow_html=True)