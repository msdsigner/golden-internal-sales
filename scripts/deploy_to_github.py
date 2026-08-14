import subprocess
import os
import json
import datetime

def deploy():
    log_path = "system_health_log.txt"
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    print("\n" + "="*60)
    print("[GitHub] Starting Deployment to GitHub Pages...")
    print("="*60)
    
    def log_message(msg):
        print(msg)
        with open(log_path, "a") as f:
            f.write(f"[{timestamp}] [GitHub] {msg}\n")

    webapp_dir = os.path.join(os.getcwd(), "webapp")
    config_path = os.path.join(os.getcwd(), "github_config.json")
    
    if not os.path.exists(config_path):
        log_message("[Error] github_config.json not found.")
        return False

    with open(config_path, "r") as f:
        config = json.load(f)
        repo_url = config.get("repo_url")
        branch = config.get("branch", "main")

    try:
        # Check if git is initialized in webapp
        if not os.path.exists(os.path.join(webapp_dir, ".git")):
            print("[Init] Initializing Git in webapp folder...")
            subprocess.run(["git", "init"], cwd=webapp_dir, check=True)
            subprocess.run(["git", "remote", "add", "origin", repo_url], cwd=webapp_dir, check=True)
            subprocess.run(["git", "checkout", "-b", branch], cwd=webapp_dir, check=True)

        log_message("[Git] Staging changes...")
        subprocess.run(["git", "add", "."], cwd=webapp_dir, check=True)
        
        commit_msg = f"Inventory Update {timestamp}"
        log_message(f"[Git] Committing: {commit_msg}")
        
        # Try to commit (if no changes, this may fail, so we catch it)
        try:
            subprocess.run(["git", "commit", "-m", commit_msg], cwd=webapp_dir, check=True, capture_output=True)
        except subprocess.CalledProcessError as e:
            if "nothing to commit" in e.stdout.decode() or "nothing to commit" in e.stderr.decode():
                print("[Info] No changes detected since last update.")
            else:
                raise e
        log_message(f"[Git] Pushing to GitHub ({branch})...")
        subprocess.run(["git", "push", "-u", "origin", branch], cwd=webapp_dir, check=True)

        log_message("[GitHub] Deployment Successful!")
        log_message(f"Site URL: https://msdsigner.github.io/golden-inventory/")
        return True
        
    except Exception as e:
        log_message(f"[GitHub] Deployment Failed: {e}")
        log_message("Tip: Make sure you have authorized Git to push to your account.")
        return False

if __name__ == "__main__":
    deploy()
