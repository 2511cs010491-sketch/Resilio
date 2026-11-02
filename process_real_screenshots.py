from PIL import Image
import os

os.makedirs("processed_real_media", exist_ok=True)

files = {
    "login": "real_screenshots/login.png",
    "dashboard": "real_screenshots/dashboard.png",
    "ai_assistant": "real_screenshots/ai_assistant.png",
    "map": "real_screenshots/map.png",
    "sos": "real_screenshots/sos.png",
    "authority": "real_screenshots/authority_dashboard.png",
    "landing": "real_screenshots/landing.png",
    "shelters": "real_screenshots/shelters.png",
    "report_incident": "real_screenshots/report_incident.png",
    "hospital": "real_screenshots/hospital_dashboard.png",
}

for key, pth in files.items():
    out_jpg = f"processed_real_media/{key}.jpg"
    with Image.open(pth) as im:
        rgb = im.convert('RGB')
        rgb.save(out_jpg, "JPEG", quality=95)
        print(f"Processed {key}: {rgb.size} -> {out_jpg}")

print("All real screenshots processed to JPEG successfully.")
