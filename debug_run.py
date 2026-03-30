import sys
import os
os.chdir(os.path.dirname(os.path.abspath(__file__)))

print("1. Starting...", flush=True)

# Execute everything from main.py but with debug
with open("main.py", "r", encoding="utf-8") as f:
    source = f.read()

# Remove the if __name__ guard
source = source.replace('if __name__ == "__main__":', 'if True:  # debug')

# Add debug prints
source = source.replace(
    'app = QApplication.instance() or QApplication(sys.argv)',
    'print("2. Creating QApp...", flush=True)\n    app = QApplication.instance() or QApplication(sys.argv)\n    print("3. QApp OK", flush=True)'
)
source = source.replace(
    'window = ScheduleApp()',
    'print("4. Creating window...", flush=True)\n    window = ScheduleApp()\n    print("5. Window created", flush=True)'
)
source = source.replace(
    'window.show()',
    'print("6. Showing...", flush=True)\n    window.show()\n    print("7. Visible!", flush=True)'
)

exec(compile(source, "main.py", "exec"))
