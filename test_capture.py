import subprocess
import os

chrome = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
out_img = os.path.abspath("test_chrome_shot.png")

if os.path.exists(out_img):
    os.remove(out_img)

cmd = [
    chrome,
    "--headless=new",
    "--disable-gpu",
    "--no-sandbox",
    "--window-size=1280,800",
    f"--screenshot={out_img}",
    "http://localhost:3000"
]

print("Running command:", " ".join(cmd))
res = subprocess.run(cmd, capture_output=True, text=True, timeout=15)
print("Return code:", res.returncode)
print("Stdout:", res.stdout)
print("Stderr:", res.stderr)
print("Output file exists:", os.path.exists(out_img), os.path.getsize(out_img) if os.path.exists(out_img) else 0)
