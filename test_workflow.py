"""
Quick test script to verify the ESG workflow is working correctly.
This will generate a sample Excel form without requiring PDF documents.
"""

import os
from dotenv import load_dotenv
from agents.excel_generator import ESGFormGenerator

# Load environment variables
load_dotenv()

def test_excel_generation():
    """Test the Excel generation without running the full CrewAI analysis."""
    
    print("Testing ESG Excel Form Generation...")
    print("=" * 50)
    
    # Sample metrics structure (what the crew would normally generate)
    sample_metrics = {
        "Environmental": {
            "Energy": [
                {
                    "name": "Total Energy Consumption",
                    "description": "Annual energy usage from all sources",
                    "unit": "kWh",
                    "frequency": "Annual",
                    "data_type": "numeric"
                },
                {
                    "name": "Renewable Energy %",
                    "description": "Percentage of energy from renewable sources",
                    "unit": "%",
                    "frequency": "Annual",
                    "data_type": "percentage"
                }
            ],
            "Emissions": [
                {
                    "name": "Scope 1 Emissions",
                    "description": "Direct GHG emissions",
                    "unit": "tCO2e",
                    "frequency": "Annual",
                    "data_type": "numeric"
                }
            ]
        },
        "Social": {
            "Employees": [
                {
                    "name": "Total Employees",
                    "description": "Number of employees",
                    "unit": "count",
                    "frequency": "Annual",
                    "data_type": "integer"
                }
            ]
        },
        "Governance": {
            "Board": [
                {
                    "name": "Board Diversity",
                    "description": "Percentage of diverse board members",
                    "unit": "%",
                    "frequency": "Annual",
                    "data_type": "percentage"
                }
            ]
        }
    }
    
    # Sample properties
    sample_properties = [
        {"name": "Test Property A", "type": "Office", "location": "New York"},
        {"name": "Test Property B", "type": "Retail", "location": "California"}
    ]
    
    # Create Excel generator
    generator = ESGFormGenerator()
    
    # Ensure output directory exists
    os.makedirs("outputs", exist_ok=True)
    
    # Generate the form
    output_path = "outputs/TEST_ESG_Form.xlsx"
    
    try:
        generator.create_comprehensive_form(
            metrics_data=sample_metrics,
            properties=sample_properties,
            investor_name="Test Investor",
            output_path=output_path
        )
        
        print(f"✅ Success! Test Excel form generated at: {output_path}")
        print("\nThe form includes:")
        print("  - Overview sheet with instructions")
        print("  - Metrics summary")
        print("  - Consolidated data entry sheet")
        print("  - Individual property sheets")
        print("  - Validation rules")
        
    except Exception as e:
        print(f"❌ Error generating Excel form: {str(e)}")
        raise

def test_crew_setup():
    """Test if CrewAI and agents can be initialized."""
    print("\nTesting CrewAI Setup...")
    print("=" * 50)
    
    try:
        from agents.esg_crew import ESGAnalysisCrew
        
        # Check if API key is set
        api_key = os.getenv("GOOGLE_API_KEY")
        if not api_key:
            print("❌ GOOGLE_API_KEY not found in .env file")
            print("   Please add your Gemini API key to the .env file")
            return False
        
        print("✅ Google API key found")
        
        # Try to initialize crew (without running it)
        crew = ESGAnalysisCrew(investor_name="Test")
        print("✅ ESG Analysis Crew initialized successfully")
        
        return True
        
    except Exception as e:
        print(f"❌ Error setting up CrewAI: {str(e)}")
        return False

if __name__ == "__main__":
    # Test Excel generation first (doesn't require API key)
    test_excel_generation()
    
    # Then test CrewAI setup
    print("\n")
    if test_crew_setup():
        print("\n🎉 All tests passed! The workflow is ready to use.")
        print("\nNext steps:")
        print("1. Add your ESG PDF documents to:")
        print("   - data/regulations/ (for regulatory documents)")
        print("   - data/frameworks/ (for investor frameworks)")
        print("2. Run main.py to generate your ESG data capture form")
    else:
        print("\n⚠️  Some tests failed. Please check the errors above.")