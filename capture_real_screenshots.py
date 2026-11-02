import subprocess
import os
import time

chrome = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
os.makedirs("real_screenshots", exist_ok=True)

routes = [
    ("login", "http://localhost:3000/login"),
    ("dashboard", "http://localhost:3000/dashboard"),
    ("map", "http://localhost:3000/map"),
    ("sos", "http://localhost:3000/sos"),
    ("ai_assistant", "http://localhost:3000/ai-assistant"),
    ("authority_dashboard", "http://localhost:3000/authority-dashboard"),
    ("shelters", "http://localhost:3000/shelters"),
    ("report_incident", "http://localhost:3000/report-incident"),
    ("hospital_dashboard", "http://localhost:3000/hospital-dashboard"),
    ("landing", "http://localhost:3000/"),
]

for name, url in routes:
    out_file = os.path.abspath(f"real_screenshots/{name}.png")
    if os.path.exists(out_file):
        os.remove(out_file)
    
    cmd = [
        chrome,
        "--headless=new",
        "--disable-gpu",
        "--no-sandbox",
        "--window-size=1280,820",
        "--virtual-time-budget=2500",  # Allow React components and Leaflet tiles to render
        f"--screenshot={out_file}",
        url
    ]
    print(f"Capturing {name} from {url}...")
    try:
        res = subprocess.run(cmd, capture_output=True, text=True, timeout=12)
        if os.path.exists(out_file):
            size = os.path.getsize(out_file)
            print(f" -> SUCCESS: {out_file} ({size} bytes)")
        else:
            print(f" -> FAILED (No file created): {res.stderr}")
    except Exception as e:
        print(f" -> ERROR: {e}")

print("\nDone capturing all screenshots.")
