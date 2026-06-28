import ollama

def generate_excel(prompt):
    response = ollama.chat(
        model="deepseek-coder-v2:16b-lite-instruct-q4_K_M",
        messages=[{"role": "user", "content": f"Buatkan struktur Excel dan rumus untuk: {prompt}"}]
    )
    return response['message']['content']