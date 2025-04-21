import os
import sys
import subprocess
import importlib.util
import shutil
from datetime import datetime

def ensure_directories():
    """Create the py and xl directories if they don't exist."""
    for directory in ['py', 'xl']:
        if not os.path.exists(directory):
            os.makedirs(directory)
            print(f"Created {directory} directory")

def run_all_scripts():
    """Run all Python scripts in the py folder and output to xl folder."""
    # Ensure directories exist
    ensure_directories()
    
    # Get list of all .py files in the py directory
    py_files = [f for f in os.listdir('py') if f.endswith('.py')]
    
    if not py_files:
        print("No Python scripts found in the /py directory.")
        return
    
    print(f"Found {len(py_files)} Python scripts to execute.")
    
    # Process each Python file
    success_count = 0
    failed_scripts = []
    
    for script_file in py_files:
        script_path = os.path.join('py', script_file)
        print(f"\n{'='*60}")
        print(f"Running: {script_file}")
        print(f"{'='*60}")
        
        try:
            # Create a modified environment with XL_OUTPUT_DIR set
            env = os.environ.copy()
            env['XL_OUTPUT_DIR'] = os.path.abspath('xl')
            
            # Run the script with the modified environment
            result = subprocess.run(
                [sys.executable, script_path],
                env=env,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True
            )
            
            # Print output
            if result.stdout:
                print("\nOutput:")
                print(result.stdout)
            
            # Check for errors
            if result.returncode != 0:
                print("\nError occurred:")
                print(result.stderr)
                failed_scripts.append(script_file)
            else:
                success_count += 1
                
                # Move any Excel files created in the current directory to the xl folder
                for file in os.listdir('.'):
                    if file.endswith(('.xlsx', '.xlsm', '.xls')):
                        shutil.move(file, os.path.join('xl', file))
                        print(f"Moved {file} to /xl directory")
            
        except Exception as e:
            print(f"Failed to execute {script_file}: {str(e)}")
            failed_scripts.append(script_file)
    
    # Print summary
    print(f"\n{'='*60}")
    print(f"Execution Summary ({datetime.now().strftime('%Y-%m-%d %H:%M:%S')})")
    print(f"{'='*60}")
    print(f"Total scripts: {len(py_files)}")
    print(f"Successfully executed: {success_count}")
    print(f"Failed: {len(failed_scripts)}")
    
    if failed_scripts:
        print("\nFailed scripts:")
        for script in failed_scripts:
            print(f"- {script}")

if __name__ == "__main__":
    run_all_scripts()