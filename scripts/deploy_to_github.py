import subprocess
import os
import json
import shutil
import datetime

def deploy():
    log_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), "system_health_log.txt")
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    print("\n" + "="*60)
    print("[GitHub] Starting GitHub Pages deployment...")
    print("="*60)

    def log_message(msg):
        print(msg)
        with open(log_path, "a") as f:
            f.write(f"[{timestamp}] [GitHub] {msg}\n")

    root_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
    webapp_dir = os.path.join(root_dir, "webapp")
    config_path = os.path.join(root_dir, "github_config.json")

    if not os.path.exists(config_path):
        log_message("[Error] github_config.json not found.")
        return False

    with open(config_path, "r") as f:
        config = json.load(f)
        repo_url = config.get("repo_url")
        branch = config.get("branch", "main")
        pages_branch = config.get("pages_branch", "gh-pages")
        site_url = config.get("site_url")

        if not site_url and repo_url:
            repo_name = repo_url.rstrip("/").split("/")[-1]
            if repo_name.endswith(".git"):
                repo_name = repo_name[:-4]
            owner = repo_url.rstrip("/").split("/")[-2]
            site_url = f"https://{owner}.github.io/{repo_name}/"

        if not site_url:
            site_url = "https://msdsigner.github.io/golden-internal-sales/"

    try:
        nested_git_dir = os.path.join(webapp_dir, ".git")
        if os.path.exists(nested_git_dir):
            print("[Cleanup] Removing stale nested git repo from webapp folder...")
            shutil.rmtree(nested_git_dir, ignore_errors=True)

        if not os.path.exists(os.path.join(root_dir, ".git")):
            print("[Init] Initializing git in project root...")
            subprocess.run(["git", "init"], cwd=root_dir, check=True)

        current_remote = subprocess.run(
            ["git", "remote", "get-url", "origin"],
            cwd=root_dir,
            capture_output=True,
            text=True,
        )
        if current_remote.returncode != 0:
            subprocess.run(["git", "remote", "add", "origin", repo_url], cwd=root_dir, check=True)
        else:
            subprocess.run(["git", "remote", "set-url", "origin", repo_url], cwd=root_dir, check=True)

        if not os.path.exists(os.path.join(webapp_dir, "index.html")):
            raise FileNotFoundError("webapp/index.html not found. GitHub Pages deployment requires a static site in webapp/")

        log_message(f"[Git] Publishing webapp/ to GitHub Pages branch '{pages_branch}'...")
        subprocess.run(["git", "subtree", "push", "--prefix", "webapp", "origin", pages_branch], cwd=root_dir, check=True)

        log_message("[GitHub] Deployment Successful!")
        log_message(f"Site URL: {site_url}")
        return True

    except Exception as e:
        log_message(f"[GitHub] Deployment Failed: {e}")
        log_message("Tip: Make sure the repository is public and GitHub Pages is enabled for the 'gh-pages' branch.")
        return False

if __name__ == "__main__":
    deploy()
