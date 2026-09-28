"""
Turns Certificate.xlsx into base64 text, for pasting into Render as the
Secret File "Certificate.xlsx.b64" (Render's secret files are text, and an
.xlsx isn't). Run it through make_certificate_secret.bat.

The output file contains the MTN certificate list - it's ignored by git,
and it's safe (and a good idea) to delete it once you've pasted it in.
"""
import base64
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
src = sys.argv[1] if len(sys.argv) > 1 else os.path.join(os.path.dirname(HERE), "Certificate.xlsx")
out = os.path.join(HERE, "Certificate.xlsx.b64.txt")

if not os.path.isfile(src):
    sys.exit(f"Couldn't find {src}")
with open(src, "rb") as f:
    text = base64.b64encode(f.read()).decode("ascii")
with open(out, "w") as f:
    # 76-character lines, like any base64 file - the app ignores line breaks
    f.write("\n".join(text[i:i + 76] for i in range(0, len(text), 76)) + "\n")
print(f"Wrote {out} ({len(text) // 1024} KB of text) from {src}")
