import requests
from Definitions import PIPEFY_API_TOKEN, PIPEFY_GRAPHQL_ENDPOINT


# Função para consultar os dados do Pipefy
def fetch_pipefy_data(pipe_id, cursor=None, page_size=30,):
    # Cursor clause to handle pagination
    cursor_clause = f', after: "{cursor}"' if cursor else ""
    
    query = f"""
    query {{
      pipe(id: "{pipe_id}") {{
        phases {{
          name
          cards(first: {page_size}{cursor_clause}) {{
            edges {{
              node {{
                title
                fields {{
                  name
                  value
                }}
              }}
            }}
            pageInfo {{
              hasNextPage
              endCursor
            }}
          }}
        }}
      }}
    }}
    """

    headers = {'Authorization': f'Bearer {PIPEFY_API_TOKEN}'}
    response = requests.post(PIPEFY_GRAPHQL_ENDPOINT, json={'query': query}, headers=headers)
    data = response.json()
    return data

