import ollama

def generate_ppt(prompt):
    response = ollama.chat(
        model="mistral-nemo:12b",
        messages=[
            {"role": "system", "content": "Jawab SELALU dalam Bahasa Indonesia."},
            {"role": "user", "content": f"Buatkan outline PPT 5 slide tentang: {prompt}"}
        ]
    )
    return response['message']['content']