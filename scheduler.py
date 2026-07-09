"""
Weekly Project Health Scheduler
================================
Runs the Project Health Reporting Agent automatically every Monday at 9:00 AM.
Keeps a log of all scheduled runs.

Usage:
    python scheduler.py
    python scheduler.py --now   # run immediately for testing

Requirements:
    pip install schedule python-dotenv
    Add your key to .env file: GEMINI_API_KEY=your_key_here

To stop: Press Ctrl+C
"""

import os
import sys
import time
import logging
from datetime import datetime
import schedule
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

# ─────────────────────────────────────────────
# LOGGING SETUP
# ─────────────────────────────────────────────
os.makedirs("logs", exist_ok=True)
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler("logs/scheduler.log", encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)
logger = logging.getLogger(__name__)


# ─────────────────────────────────────────────
# JOB FUNCTION
# ─────────────────────────────────────────────

def run_agent():
    """Execute the health reporting agent as a scheduled job."""
    api_key = os.environ.get("GEMINI_API_KEY")
    if not api_key:
        logger.error("GEMINI_API_KEY not set. Skipping this run.")
        return

    logger.info("=" * 55)
    logger.info("Scheduled run starting...")
    logger.info(f"Run time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    logger.info("=" * 55)

    try:
        # Import and run the agent's main function directly
        # so it shares the same Python environment and API key
        import importlib.util
        spec = importlib.util.spec_from_file_location("agent", "agent.py")
        agent_module = importlib.util.module_from_spec(spec)

        # Ensure GEMINI_API_KEY is available to agent module
        os.environ["GEMINI_API_KEY"] = api_key
        spec.loader.exec_module(agent_module)
        agent_module.main()

        logger.info("Scheduled run completed successfully.")
    except Exception as e:
        logger.error(f"Scheduled run failed: {e}", exc_info=True)


# ─────────────────────────────────────────────
# SCHEDULE SETUP
# ─────────────────────────────────────────────

def main():
    api_key = os.environ.get("GEMINI_API_KEY")
    if not api_key:
        print("ERROR: GEMINI_API_KEY environment variable not set.")
        print("  Run: set GEMINI_API_KEY=your_key_here")
        sys.exit(1)

    print("=" * 55)
    print("  Project Health Reporting Scheduler")
    print("  Runs every Monday at 09:00 AM")
    print("  Press Ctrl+C to stop")
    print("=" * 55)

    # Schedule: every Monday at 09:00
    schedule.every().monday.at("09:00").do(run_agent)

    # Also allow immediate one-off run with --now flag
    if "--now" in sys.argv:
        logger.info("--now flag detected: running agent immediately.")
        run_agent()

    logger.info("Scheduler started. Waiting for next Monday 09:00 run...")
    logger.info(f"Next run: {schedule.next_run()}")

    # Keep running
    while True:
        schedule.run_pending()
        time.sleep(60)  # check every minute


if __name__ == "__main__":
    main()
