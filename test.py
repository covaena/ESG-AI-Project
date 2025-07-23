from dotenv import load_dotenv
import os

load_dotenv()

api_key = os.getenv("OPENAI_API_KEY")
if api_key:
    print("✅ Environment loaded successfully!")
    print(f"API Key starts with: {api_key[:10]}...")
else:
    print("❌ No API key found. Check your .env file")