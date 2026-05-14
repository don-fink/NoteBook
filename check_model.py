import os

# Check the environment variable for the model
model = os.getenv('MODEL', 'Gemini')

if model == 'Qwen2.5-Coder:7b':
    print("Processing with local Qwen 2.5-Coder:7b model")
else:
    print(f"Processing through {model}")