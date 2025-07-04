from openai import OpenAI

def Auto_write_admission_ChatGPT(prompt_text, OPENAI_API_KEY, path= "Replied.html"):
    client = OpenAI(api_key=OPENAI_API_KEY)
    completion = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=[{
            "role": "system", 
            "content": "You are a resident doctor, who needs to write admission notes based on ER or OPD notes."
        }, {
            "role": "user", 
            "content": prompt_text
        }]
    )
    replied_text = completion.choices[0].message.content

    # Save and process the reply
    with open(path, 'w', encoding="utf-8") as f:
        f.write(replied_text)
    return (replied_text)