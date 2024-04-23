import os, ast
from dotenv import load_dotenv

# Definições
load_dotenv()

PIPEFY_API_TOKEN = os.getenv('PIPEFY_API_TOKEN')
PIPEFY_GRAPHQL_ENDPOINT = os.getenv('PIPEFY_GRAPHQL_ENDPOINT')
PIPE_TO_FILE = ast.literal_eval(os.getenv('PIPE_TO_FILE'))

# IDs dos Pipes 
PIPE_IDS = ast.literal_eval(os.getenv('PIPE_IDS'))