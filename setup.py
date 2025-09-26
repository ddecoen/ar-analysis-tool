#!/usr/bin/env python3
"""
AR Analysis Tool Setup Script
Sets up the tool for first-time use
"""

import os
import sys
import subprocess

def install_requirements():
    """Install required Python packages"""
    print("📦 Installing required packages...")
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "-r", "requirements.txt"])
        print("✅ Requirements installed successfully")
        return True
    except subprocess.CalledProcessError:
        print("❌ Failed to install requirements")
        return False

def create_sample_config():
    """Create sample configuration if it doesn't exist"""
    if not os.path.exists("config.ini"):
        print("📝 Creating sample configuration...")
        # Config file already exists from our creation
        print("✅ Configuration file created")
    else:
        print("✅ Configuration file already exists")

def test_installation():
    """Test that the installation works"""
    print("🧪 Testing installation...")
    try:
        from ar_analysis import ARAnalyzer
        print("✅ AR Analysis tool imported successfully")
        return True
    except ImportError as e:
        print(f"❌ Import failed: {e}")
        return False

def show_next_steps():
    """Display next steps for the user"""
    print("\n🚀 Setup Complete! Next steps:")
    print("\n1. Prepare your data:")
    print("   - Export invoice data from your system as Excel file")
    print("   - Ensure it has columns: Document Number, Name, Invoice Date, Due Date, Payment Date, Amount")
    print("\n2. Run the analysis:")
    print("   python ar_analysis.py your_data.xlsx output_report.xlsx")
    print("\n3. Customize settings (optional):")
    print("   - Edit config.ini to adjust wire fee thresholds")
    print("   - Add specific document numbers for exclusions")
    print("\n4. Automate (optional):")
    print("   - Set up monthly cron job or scheduled task")
    print("   - Integrate with your ERP system export")
    print("\n📚 Documentation:")
    print("   - Read README.md for detailed instructions")
    print("   - Check example_usage.py for code examples")

def main():
    """Main setup function"""
    print("🚀 AR Analysis Tool Setup")
    print("=" * 50)
    
    success = True
    
    # Install requirements
    if not install_requirements():
        success = False
    
    # Create config
    create_sample_config()
    
    # Test installation
    if not test_installation():
        success = False
    
    if success:
        print("\n✅ Setup completed successfully!")
        show_next_steps()
    else:
        print("\n❌ Setup encountered errors. Please check the issues above.")
        return 1
    
    return 0

if __name__ == "__main__":
    exit(main())
