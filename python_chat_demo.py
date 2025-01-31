from openai import OpenAI
import os   
import dotenv

dotenv.load_dotenv()

email_body = input("Ask a question:")

api_key = os.getenv('OPENAI_API_KEY')

client = OpenAI(api_key=api_key)

completion = client.chat.completions.create(
    model="gpt-4o",
    messages=[
        {"role": "system", "content": "you are a property management admin expert and your "},
        {"role": "user", "content": (email_body)},           
    ]
)

print(completion.choices[0].message.content)