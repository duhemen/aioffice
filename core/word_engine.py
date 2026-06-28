import ollama

def generate_word(prompt):
    response = ollama.chat(
        model="qwen2.5:14b-instruct-q4_K_M",
        messages=[{"role": "user", "content": f"Buatkan dokumen Word tentang: {prompt}. Format rapi dengan kop surat."}]
    )
    return response['message']['content']