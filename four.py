import openai
import pandas as pd
import time

openai.api_key = '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'

def get_completion(prompt, model="gpt-3.5-turbo"):
    # Create a new conversation for each prompt
    messages = [
        {"role": "system", "content": "You are a helpful assistant."},
        {"role": "user", "content": prompt}
    ]
    response = openai.ChatCompletion.create(
        model=model,
        messages=messages,
        temperature=0,
    )
    return response.choices[0].message["content"]

# Load the Excel file
excel_file = "C:/Users/npkar/OneDrive/Documents/Udemy/four.xlsx"  # Replace with the actual file path
df = pd.read_excel(excel_file)

# Create a new column 'Response' to store the responses
df['Response'] = ""

# Iterate through each row in the DataFrame
for index, row in df.iterrows():
    prompt = row['B']  # Replace with the actual column name containing your prompts
    
    # Retry loop for rate limit error
    while True:
        try:
            response = get_completion(prompt)
            df.at[index, 'C'] = response  # Store the response in the 'Response' column
            break  # Break out of the loop if successful
        except openai.error.RateLimitError as e:
            print(f"Rate limit exceeded. Waiting for 20 seconds. Error: {e}")
            time.sleep(20)  # Wait for 20 seconds before retrying

    # Save the updated DataFrame with the responses back to the Excel file after each row
    df.to_excel(excel_file, index=False)

print(f"Responses saved to {excel_file}")
