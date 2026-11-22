"""Build the iLA Outlook add-in package (unified manifest for Microsoft 365).

Usage:
    python build.py https://yyi921.github.io/ila-outlook-addin
    python build.py https://ila.works/ila2          # later, production

Writes package/manifest.json and dist/ila-outlook.zip (upload that zip).
Bump VERSION every time you change the manifest.
"""
import json, sys, zipfile
from pathlib import Path
from urllib.parse import urlparse

APP_ID = "46698586-e643-462a-95a0-5af6ff73ceca"   # keep fixed for this app
VERSION = "1.0.1"
ROOT = Path(__file__).parent

def build(base: str) -> Path:
    base = base.rstrip("/")
    if not base.startswith("https://"):
        sys.exit("Base URL must be HTTPS.")
    host = urlparse(base).netloc
    icons = [{"size": s, "url": f"{base}/assets/icon-{s}.png"} for s in (16, 32, 80)]
    mailbox = [{"name": "Mailbox", "minVersion": "1.8"}]   # 1.8: read attachment content

    manifest = {
        "$schema": "https://developer.microsoft.com/json-schemas/teams/v1.17/MicrosoftTeams.schema.json",
        "manifestVersion": "1.17",
        "id": APP_ID,
        "version": VERSION,
        "name": {"short": "iLA", "full": "iLA - iLoveAccounting for Outlook"},
        "description": {
            "short": "Turn emailed DDTAs into Lightyear bills.",
            "full": "iLA reads the open email, builds an AP invoice for each attached DDTA, "
                    "joins it to the DDTA scan and emails each one to Lightyear.",
        },
        "developer": {
            "name": "AFIS Consultancy Pty Ltd",
            "websiteUrl": "https://ila.works",
            "privacyUrl": "https://ila.works/privacy",
            "termsOfUseUrl": "https://ila.works/terms",
        },
        "icons": {"color": "color.png", "outline": "outline.png"},
        "accentColor": "#0F52BA",
        "localizationInfo": {"defaultLanguageTag": "en-au", "additionalLanguages": []},
        "validDomains": [host],
        "showLoadingIndicator": False,
        "isFullScreen": False,
        "defaultBlockUntilAdminAction": False,
        "authorization": {"permissions": {"resourceSpecific": [
            {"name": "MailboxItem.Read.User", "type": "Delegated"}
        ]}},
        "extensions": [{
            "requirements": {"scopes": ["mail"], "capabilities": mailbox},
            "runtimes": [{
                "requirements": {"capabilities": mailbox},
                "id": "TaskPaneRuntime",
                "type": "general",
                "code": {"page": f"{base}/taskpane.html"},
                "lifetime": "short",
                "actions": [{"id": "TaskPaneRuntimeShow", "type": "openPage", "pinnable": True}],
            }],
            "ribbons": [{
                "contexts": ["mailRead"],
                "tabs": [{
                    "builtInTabId": "TabDefault",
                    "groups": [{
                        "id": "ilaReadGroup",
                        "label": "iLA",
                        "icons": icons,
                        "controls": [{
                            "id": "ilaOpenPane",
                            "type": "button",
                            "label": "Open iLA",
                            "icons": icons,
                            "supertip": {"title": "Open iLA",
                                         "description": "Send the DDTAs attached to this email to Lightyear."},
                            "actionId": "TaskPaneRuntimeShow",
                        }],
                    }],
                }],
            }],
        }],
    }

    pkg = ROOT / "package"
    (pkg / "manifest.json").write_text(json.dumps(manifest, indent=2))
    dist = ROOT / "dist"; dist.mkdir(exist_ok=True)
    out = dist / "ila-outlook.zip"
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:   # manifest + 2 icons only, at zip root
        for f in ("manifest.json", "color.png", "outline.png"):
            z.write(pkg / f, f)
    return out

if __name__ == "__main__":
    print("Built", build(sys.argv[1] if len(sys.argv) > 1 else "https://yyi921.github.io/ila-outlook-addin"))
