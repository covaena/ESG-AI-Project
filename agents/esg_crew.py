# /agents/esg_crew.py

import os
import json
from typing import List, Dict, Any

# Core CrewAI imports
from crewai import Agent, Task, Crew, Process

# Tool imports from their specific submodules
from crewai_tools.pdf_search_tool import PDFSearchTool
from crewai_tools.file_read_tool import FileReadTool
from crewai_tools.directory_read_tool import DirectoryReadTool

# The crucial LangChain integration for Google Generative AI
from langchain_google_genai import ChatGoogleGenerativeAI

class ESGAnalysisCrew:
    """CrewAI crew for analyzing ESG requirements and generating metrics."""
    
    def __init__(self, investor_name: str):
        self.investor_name = investor_name
        
        # Instantiate the Gemini LLM using the LangChain wrapper
        # This single object is reused by all agents and tools.
        self.llm = ChatGoogleGenerativeAI(
            model="gemini-1.5-flash-latest",
            api_key=os.getenv("GOOGLE_API_KEY")
        )
        
        self._setup_tools()
        self._create_agents()
        self._create_tasks()
        self._create_crew()
    
    def _setup_tools(self):
        """Initialize all tools for document processing."""
        self.regulation_search_tool = PDFSearchTool()
        self.file_reader = FileReadTool()
        self.regulation_dir_tool = DirectoryReadTool(directory='./data/regulations')
        self.framework_dir_tool = DirectoryReadTool(directory='./data/frameworks')
    
    def _create_agents(self):
        """Create specialized agents for ESG analysis."""
        
        # All agents now use the single, instantiated self.llm object
        self.regulatory_agent = Agent(
            role='ESG Regulatory Compliance Specialist',
            goal='Extract all mandatory ESG metrics and reporting requirements from local regulations',
            backstory="You are an expert in ESG regulatory compliance...",
            tools=[self.regulation_search_tool, self.regulation_dir_tool, self.file_reader],
            llm=self.llm,
            max_iter=5,
            verbose=True
        )
        
        self.framework_agent = Agent(
            role='Investor ESG Framework Analyst',
            goal=f'Identify all ESG metrics and KPIs required by {self.investor_name}...',
            backstory="You specialize in understanding investor ESG frameworks...",
            tools=[self.framework_dir_tool, self.regulation_search_tool, self.file_reader],
            llm=self.llm,
            max_iter=5,
            verbose=True
        )
        
        self.consolidation_agent = Agent(
            role='ESG Metrics Architect',
            goal='Consolidate and structure all ESG metrics into a comprehensive data collection framework',
            backstory="You are an expert at organizing complex ESG requirements...",
            llm=self.llm,
            max_iter=3,
            verbose=True
        )
    
    def _create_tasks(self):
        """Define tasks for the crew."""
        
        self.regulatory_task = Task(
            description="Analyze all PDF files in the regulations folder...",
            expected_output="A detailed list of all regulatory ESG metrics...",
            agent=self.regulatory_agent
        )
        
        self.framework_task = Task(
            description=f"Analyze {self.investor_name}'s ESG framework documents...",
            expected_output=f"A comprehensive list of {self.investor_name}'s ESG metrics...",
            agent=self.framework_agent
        )
        
        self.consolidation_task = Task(
            description="Consolidate all ESG metrics from regulations and investor requirements...",
            expected_output="A JSON-structured comprehensive ESG metrics framework...",
            agent=self.consolidation_agent,
            # Ask CrewAI to ensure the final output is a parsed JSON object
            output_json=True
        )
    
    def _create_crew(self):
        """Assemble the crew with defined agents and tasks."""
        self.crew = Crew(
            agents=[self.regulatory_agent, self.framework_agent, self.consolidation_agent],
            tasks=[self.regulatory_task, self.framework_task, self.consolidation_task],
            verbose=2,
            # Use the Process enum for sequential task execution
            process=Process.sequential,
        )
    
    def run(self, properties: List[Dict[str, str]]) -> Dict[str, Any]:
        """Execute the crew's analysis."""
        result = self.crew.kickoff(
            inputs={
                "investor_name": self.investor_name,
                "properties": properties,
            }
        )
        
        # With output_json=True, the result should be a dictionary directly
        if isinstance(result, dict):
            return result
        else:
            print(f"Warning: Crew output was not a dictionary. Raw output: {result}")
            # Fallback to a default structure if parsing fails
            return self._create_basic_metrics_structure()

    def _create_basic_metrics_structure(self) -> Dict[str, Any]:
        """Create a basic ESG metrics structure as a fallback."""
        # This method remains unchanged
        return {
            "Environmental": {"Energy": [{"name": "Total Energy Consumption", "unit": "kWh"}]},
            "Social": {"Employees": [{"name": "Total Employees", "unit": "count"}]},
            "Governance": {"Board": [{"name": "Board Diversity", "unit": "%"}]}
        }