import requests
import json
from models.memory_extraction_response import MemoryExtractionResponse

chat_url = "http://localhost:11434/api/chat"

# -----------------------------
# Long-term memory
# -----------------------------
long_term_memory = []

# -----------------------------
# Conversation history
# -----------------------------
msg_list = []

# -----------------------------
# System Prompt
# -----------------------------
system_prompt = """
You are a helpful AI assistant.

Use the known facts about the user whenever appropriate.
Do not make up facts.
"""

# -----------------------------
# Prompt used for memory extraction
# -----------------------------
memory_prompt = """
You are a memory extraction assistant.

Your job is to identify ONLY long-term facts about the user.

Examples of things to remember:
- Name
- Profession
- Skills
- Hobbies
- Favourite sport
- Favourite food
- Interests

DO NOT remember:
- Greetings
- Temporary requests
- Questions
- Current mood
- Small talk

Return ONLY valid JSON.

Example:

{
    "memories": [
        "User's favourite sport is Cricket.",
        "User works on OpenShift."
    ]
}
"""


# --------------------------------------
# Extract long-term memories
# --------------------------------------
def update_memory(chat_history):
    schema = MemoryExtractionResponse.model_json_schema()
    payload = {
        "model": "gemma3:1b",
        "messages": [
            {
                "role": "system",
                "content": memory_prompt
            }
        ] + chat_history,
        "format": schema,
        "stream": False
    }

    print(f"\nMemory extraction payload: {payload['messages']}")
    response = requests.post(chat_url, json=payload)
    print(f"\nMemory extraction response: {response.json()}")
    content = response.json()["message"]["content"]

    try:
        
        memory_json = json.loads(content)
        return memory_json.get("memories", [])

    except Exception:
        print("\nCould not parse memory response.")
        print(content)
        return []


# --------------------------------------
# Chat loop
# --------------------------------------
while True:

    prompt = input("\nYou: ")

    if prompt.lower() == "exit":
        break

    # Add user message
    msg_list.append({
        "role": "user",
        "content": prompt
    })

    # Build messages to send
    messages = [
        {
            "role": "system",
            "content": system_prompt
        }
    ]

    if long_term_memory:

        memory_text = "Known facts about the user:\n"

        for memory in long_term_memory:
            memory_text += f"- {memory}\n"

        messages.append({
            "role": "system",
            "content": memory_text
        })

    messages.extend(msg_list)

    payload = {
        "model": "gemma3:1b",
        "messages": messages,
        "stream": False
    }

    response = requests.post(chat_url, json=payload)

    assistant_msg = response.json()["message"]

    print(f"\nAI: {assistant_msg}") 

    msg_list.append(assistant_msg)

    # Count user messages
    user_messages = len(
        [m for m in msg_list if m["role"] == "user"]
    )

    # Every 3 user messages, update memory
    if user_messages % 3 == 0:

        print("\nUpdating long-term memory...\n")

        new_memories = update_memory(msg_list)

        for memory in new_memories:
            if memory not in long_term_memory:
                long_term_memory.append(memory)

        print("Current Long-Term Memory")

        for m in long_term_memory:
            print("-", m)