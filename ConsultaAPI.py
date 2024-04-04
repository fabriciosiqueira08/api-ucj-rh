import requests

# Substitua 'SEU_TOKEN_AQUI' pelo seu token de API real
PIPEFY_API_TOKEN = "eyJhbGciOiJIUzUxMiJ9.eyJpc3MiOiJQaXBlZnkiLCJpYXQiOjE3MTE3MTY3MzgsImp0aSI6IjMwYzBiNjMzLTg5ZWEtNDM0ZC1hODMxLTZkYjc3OTM0MTBkNiIsInN1YiI6MjEyODA2LCJ1c2VyIjp7ImlkIjoyMTI4MDYsImVtYWlsIjoicmhAdWNqLmNvbS5iciIsImFwcGxpY2F0aW9uIjozMDAzMzgyMzIsInNjb3BlcyI6W119LCJpbnRlcmZhY2VfdXVpZCI6bnVsbH0.79e7athW43b4WrBvWOsxa4wsIEUbQlVRzdU6rlZ4pmjDB2ABiv8sOyPu0jv18Gj5HCkue4QIMavqAqE2CnMHiQ"
PIPEFY_GRAPHQL_ENDPOINT = 'https://api.pipefy.com/graphql'

# Substitua esta consulta pela sua consulta GraphQL
query = """
query {
  pipe(id: "301823995") {
    id
    name
    phases {
      id
      name
      cards_count
      cards {
        edges {
          node {
            id
            title
            fields {
              name
              value
            }
          }
        }
      }
    }
  }
}
"""

def fetch_pipefy_data():
    headers = {
        'Authorization': f'Bearer {PIPEFY_API_TOKEN}',
        'Content-Type': 'application/json'
    }

    response = requests.post(PIPEFY_GRAPHQL_ENDPOINT, json={'query': query}, headers=headers)

    if response.status_code == 200:
        # Converte a resposta JSON para um dicionário Python e retorna
        return response.json()
    else:
        # Trata erros de resposta da API, como token inválido ou problemas de rede
        print(f'Erro na consulta à API: {response.status_code}')
        print(response.text)
        return None

def main():
    data = fetch_pipefy_data()
    if data:
        print("Resposta da API do Pipefy:")
        print(data)
    else:
        print("Falha ao receber dados da API do Pipefy.")

if __name__ == "__main__":
    main()
