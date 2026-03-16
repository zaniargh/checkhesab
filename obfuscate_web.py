import zlib
import base64
import os
import shutil

DIST_DIR = "dist_web"

print("Preparing web distribution folder...")
if os.path.exists(DIST_DIR):
    shutil.rmtree(DIST_DIR)
os.makedirs(DIST_DIR)

print("Reading and obfuscating app.py...")
with open("app.py", "r", encoding="utf-8") as f:
    source = f.read()

# Compress the source code to save space and hide content
compressed = zlib.compress(source.encode("utf-8"), level=9)
# Encode it in base64 to make it text-safe
encoded_script = base64.b64encode(compressed).decode("utf-8")

# Template for the secured app.py
obfuscated_code = f"""# Secured Checkhesab Application
import base64, zlib
_A="{encoded_script}"
exec(zlib.decompress(base64.b64decode(_A)).decode("utf-8"), globals())
"""

with open(f"{DIST_DIR}/app.py", "w", encoding="utf-8") as f:
    f.write(obfuscated_code)

print("Copying HTML templates and static files...")
shutil.copy("index.html", f"{DIST_DIR}/index.html")
shutil.copy("login.html", f"{DIST_DIR}/login.html")
shutil.copytree("static", f"{DIST_DIR}/static")

print("======================================================")
print(f"Done! Your web-ready secured application is in '{DIST_DIR}'.")
print("You can upload the contents of this folder to your web server.")
print("It works with standard Python exactly like the original app.py.")
print("======================================================")
