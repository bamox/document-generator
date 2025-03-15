#!/usr/bin/env python3
import os
import sys
import subprocess
import platform
import importlib.util
import pkg_resources

def check_python_version():
    """Check if Python version is compatible"""
    required_major = 3
    required_minor = 7
    
    current_major = sys.version_info.major
    current_minor = sys.version_info.minor
    
    if current_major < required_major or (current_major == required_major and current_minor < required_minor):
        print(f"Error: Python {required_major}.{required_minor} or higher is required.")
        print(f"Current Python version: {current_major}.{current_minor}")
        sys.exit(1)
    
    return True

def is_package_installed(package_name):
    """Check if a package is already installed"""
    try:
        pkg_resources.get_distribution(package_name)
        return True
    except pkg_resources.DistributionNotFound:
        return False

def check_and_install_requirements():
    """Check if required packages are installed and install only missing ones"""
    requirements = {
        "streamlit": "1.22.0",
        "pandas": "1.3.0",
        "python-docx": "0.8.11",
        "openpyxl": "3.0.9",
        "odfpy": "1.4.1", 
        "xlrd": "2.0.1",
        "pyxlsb": "1.0.9"
    }
    
    missing_packages = []
    
    # Check which packages need to be installed
    for package, min_version in requirements.items():
        if not is_package_installed(package):
            missing_packages.append(f"{package}>={min_version}")
        else:
            print(f"✅ {package} is already installed")
    
    # Install missing packages if any
    if missing_packages:
        print(f"Installing {len(missing_packages)} missing packages...")
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install"] + missing_packages)
            print("All packages installed successfully!")
        except subprocess.CalledProcessError as e:
            print(f"Error installing packages: {str(e)}")
            sys.exit(1)
    else:
        print("All required packages are already installed!")
    
    return True

def create_requirements_file():
    """Create a requirements.txt file if it doesn't exist"""
    if os.path.exists("requirements.txt"):
        print("requirements.txt already exists")
        return
        
    requirements = [
        "streamlit>=1.22.0",
        "pandas>=1.3.0",
        "python-docx>=0.8.11",
        "openpyxl>=3.0.9",
        "odfpy>=1.4.1",
        "xlrd>=2.0.1",
        "pyxlsb>=1.0.9",
    ]
    
    with open("requirements.txt", "w") as f:
        f.write("\n".join(requirements))
    
    print("Created requirements.txt file")

def check_generator_script():
    """Check if generate.py exists"""
    if not os.path.exists("generate.py"):
        print("Error: generate.py not found in the current directory.")
        print("Please make sure the file exists before running this script.")
        sys.exit(1)
    return True

def run_generator():
    """Run the document generator script"""
    print("\nStarting Document Generator application...")
    
    try:
        # Run streamlit directly using the module approach
        subprocess.run([sys.executable, "-m", "streamlit", "run", "generate.py"], check=True)
    except subprocess.CalledProcessError as e:
        print(f"Error running the application: {str(e)}")
        sys.exit(1)

def main():
    """Main function to set up and run the document generator"""
    print("Document Generator Setup Script")
    print("==============================\n")
    
    # Check Python version
    check_python_version()
    
    # Check if generate.py exists
    check_generator_script()
    
    # Check and install only missing requirements
    check_and_install_requirements()
    
    # Create requirements file for future use (only if it doesn't exist)
    create_requirements_file()
    
    # Run the generator
    run_generator()

if __name__ == "__main__":
    main()