
import os
import subprocess
import sys

def convert_pptx_to_pdf(input_path, output_path):
    input_path = os.path.abspath(input_path)
    output_path = os.path.abspath(output_path)
    
    script = f'''
    tell application "Microsoft PowerPoint"
        set isRunning to running
        activate
        try
            open POSIX file "{input_path}"
            set activePres to active presentation
            save activePres in "{output_path}" as save as PDF
            close activePres saving no
        on error errMsg
            display dialog "Error: " & errMsg
        end try
        
        -- restore state
        if not isRunning then quit
    end tell
    '''
    
    try:
        process = subprocess.run(
            ['osascript', '-e', script], 
            capture_output=True, 
            text=True
        )
        if process.returncode != 0:
            print(f"Error: {process.stderr}")
            return False
        return True
    except Exception as e:
        print(f"Exception: {e}")
        return False

if __name__ == "__main__":
    # Create a dummy PPTX if needed, or just print the script for review
    print("Script prepared.")
