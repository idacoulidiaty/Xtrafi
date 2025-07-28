import os

ENV = os.getenv("XTRAFI_ENV", "preprod")  # ou "prod"

def get_base_folder(username, orga):
    if ENV == "preprod":
        from pathlib import Path
        return Path.home() / "Documents/Xtrafi/Dataviz" / orga / username
    else:
        return f"https://api.xtrafi.com/files?user={username}&orga={orga}"  # endpoint API
