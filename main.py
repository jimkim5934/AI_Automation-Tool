"""ONT Automation Tool — application entry point."""

import sys
import tkinter as tk

from ont_automation.gui_app import AutomationGUI


def main():
    root = tk.Tk()
    app = AutomationGUI(root)

    try:
        root.mainloop()
    except KeyboardInterrupt:
        print("\n[!] Program interrupted by user. Exiting safely...")
    finally:
        if app.olt:
            app.olt.disconnect()
        sys.exit(0)


if __name__ == "__main__":
    main()
