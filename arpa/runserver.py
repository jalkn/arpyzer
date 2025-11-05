import os
import sys
import webbrowser
from django.core.management import execute_from_command_line

# IMPORTANT: Set this to your actual project name ('arpa')
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'arpa.settings')

def main():
    port = '8000'
    server_address = f'http://127.0.0.1:{port}'

    # Check if we are running in the packaged executable
    if getattr(sys, 'frozen', False):
        print(f"Starting Django server at {{server_address}}...")

        # Automatically open the browser
        import threading
        def open_browser():
             import time
             time.sleep(1) # Give the server a moment to start
             webbrowser.open_new(server_address)

        threading.Thread(target=open_browser).start()

        # Run the Django server command, disabling the reloader
        execute_from_command_line(['manage.py', 'runserver', f'127.0.0.1:{{port}}', '--noreload'])

    else:
        # Standard development run
        print("Running Django in development mode...")
        execute_from_command_line(sys.argv)

if __name__ == '__main__':
    main()
