import subprocess
import sys

def run_script(script_name):
    try:
        subprocess.call([sys.executable, script_name])
    except Exception as e:
        print("Error running {}: {}".format(script_name, e))

if __name__ == "__main__":
    # Run extract_images.py
    print("Running extract_images.py...")
    run_script("extract_images.py")

    # Run selenium_automation.py
    print("Running selenium_automation.py...")
    run_script("selenium_automation.py")

    # Run add_footer.py
    print("Running add_footer.py...")
    run_script("add_footer.py")

    print("All scripts executed successfully.")
