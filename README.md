Summarised by Gemini


ESG Data Capture Automation Agent
This project uses a multi-agent AI system to read ESG regulatory and investor documents (PDFs) and automatically generate a structured Excel data capture form.

Core Workflow & Architecture
The system operates as a sequential pipeline, orchestrating a crew of AI agents to process documents and generate a report.

Initiation: The process starts when main.py is executed. It imports and kicks off the ESGAnalysisCrew.


AI Crew Kickoff (agents/esg_crew.py):

Agent 1: Regulatory Compliance Specialist: Scans documents in the data/regulations folder. It uses a vector database (ChromaDB) to find and extract all regulatory metrics.

Agent 2: Investor Framework Analyst: Performs the same analysis for investor guidelines located in the data/frameworks folder.

Agent 3: ESG Metrics Architect: Receives the findings from the first two agents. It consolidates the lists, removes duplicate metrics, and formats the result into a clean JSON structure.

Handoff & Generation: The final JSON object is passed from the AI crew to the ESGFormGenerator (agents/excel_generator.py). This module uses the structured data to build a formatted, multi-tabbed Excel file.

Final Output: The generated .xlsx report is saved in the outputs/ directory.

File Structure Overview
agents/: The "AI Workforce" 🧠. Contains the core logic for the AI agent crew and the final Excel generator.

data/: The "In-Tray" 📥. This is where you place your input PDFs, sorted into regulations and frameworks subfolders.

chroma_db/: The "AI's Memory" 🗄️. The vector database where document embeddings are stored for fast, semantic searching.

outputs/: The "Out-Tray". The destination folder for the final, generated Excel spreadsheets.

main.py: The central control script that runs the entire end-to-end workflow.

