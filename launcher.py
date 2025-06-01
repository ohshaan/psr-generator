import os
import sys
import subprocess

# Determine the base path (either the frozen temp folder or the script's folder)
if getattr(sys, "frozen", False):
    base_path = sys._MEIPASS
else:
    base_path = os.path.dirname(os.path.abspath(__file__))

# Build the path to app.py
app_path = os.path.join(base_path, "app.py")

# Launch Streamlit (omit --server.headless so the browser will open)
cmd = [
    sys.executable,
    "-m",
    "streamlit",
    "run",
    app_path
]

# Start Streamlit in a separate process; this EXE can then terminate if you prefer,
# or keep running as long as you need the server alive.
subprocess.Popen(cmd, shell=False)
