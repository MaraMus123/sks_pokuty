import os
from dotenv import load_dotenv

load_dotenv(r"C:\Users\musin\IdeaProjects\email\.env")

variable_name = os.environ.get('VARIABLE_NAME')
print(variable_name)