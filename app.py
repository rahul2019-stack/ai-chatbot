import requests
import time

# single_prompt_url = "http://localhost:11434/api/generate"
chat_url = "http://localhost:11434/api/chat"
msg_list = []
while True: 
    prompt = input("Enter the prompt for the model: ")
    input_msg_dict = {"role": "user", "content": prompt}
    
    msg_list.append(input_msg_dict)
    print(f"msg list is {msg_list}")
    if prompt.lower() == "exit":
        break
    payload = {
        "model": "gemma3:1b",
        "messages": msg_list,
        "stream": False
    }
    resp = requests.post(chat_url, json=payload)
    op_content = resp.json()['message']
    print(f"Output is {op_content}")
    msg_list.append(op_content)
    