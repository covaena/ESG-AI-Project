import os
from dotenv import load_dotenv
from agents.esg_crew import ESGAnalysisCrew
from agents.excel_generator import ESGFormGenerator
import json

# Load environment variables
load_dotenv()

def main():
    """Main workflow orchestrator for ESG data capture form generation."""
    
    # Configuration
    investor_name = "BlackRock"  # Change this based on your needs
    properties = [
        {"name": "Property A", "type": "Office", "location": "New York"},
        {"name": "Property B", "type": "Retail", "location": "California"},
        {"name": "Property C", "type": "Industrial", "location": "Texas"}
    ]
    
    print("=" * 60)
    print(f"ESG Data Capture Workflow")
    print(f"Investor: {investor_name}")
    print(f"Properties: {len(properties)}")
    print("=" * 60)
    
    # Initialize the ESG Analysis Crew
    print("\n1. Initializing ESG Analysis Crew...")
    esg_crew = ESGAnalysisCrew(investor_name=investor_name)
    
    # Run the analysis
    print("\n2. Analyzing ESG requirements...")
    print("   - Scanning local regulations...")
    print("   - Reviewing investor frameworks...")
    print("   - Consolidating metrics...")
    
    try:
        # Execute the crew's analysis
        result = esg_crew.run(properties=properties)
        
        # Print intermediate results for debugging
        print("\n3. Analysis Complete!")
        print(f"   - Found metrics across {len(result.get('categories', {}))} categories")
        
        # Generate Excel form
        print("\n4. Generating Excel data capture form...")
        form_generator = ESGFormGenerator()
        
        output_filename = f"ESG_DataCapture_{investor_name}_{len(properties)}_properties.xlsx"
        output_path = os.path.join("outputs", output_filename)
        
        # Ensure outputs directory exists
        os.makedirs("outputs", exist_ok=True)
        
        # Create the Excel form
        form_generator.create_comprehensive_form(
            metrics_data=result,
            properties=properties,
            investor_name=investor_name,
            output_path=output_path
        )
        
        print(f"\n✅ Success! ESG data capture form generated:")
        print(f"   📁 {output_path}")
        print(f"\n   The form includes:")
        print(f"   - Overview and instructions")
        print(f"   - Consolidated metrics sheet")
        print(f"   - Individual property sheets")
        print(f"   - Data validation and formulas")
        
    except Exception as e:
        print(f"\n❌ Error during analysis: {str(e)}")
        print("Please check that:")
        print("  - Your PDF files are in the correct folders")
        print("  - Your Google API key is set correctly")
        print("  - All required packages are installed")
        raise

if __name__ == "__main__":
    main()