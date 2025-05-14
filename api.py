from dotenv import load_dotenv
import os
import re
import requests

load_dotenv()

apikey = os.getenv('APIKEY')

catalog = os.getenv('CATALOG')

def api_request(task):
    prompt = {
        "modelUri": f"gpt://{catalog}/yandexgpt-lite",
        "completionOptions": {
            "stream": False,
            "temperature": 0.3,
            "maxTokens": "2000"
        },
        "messages": [
            {
                "role": "system",
                "text": "Ты коротко объяснаяшь решение математических задач."
            },
            {
                "role": "user",
                "text": f"Напиши решение задачи: {task}"
            }
                ]
    }
    url = "https://llm.api.cloud.yandex.net/foundationModels/v1/completion"
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {apikey}"
    }
    response = requests.post(url, headers=headers, json=prompt)
    result = response.text
    find = re.compile(r'\"text\"\:\"(.+)\"\}\,\"status\"')
    return re.findall(find, result)[0]
