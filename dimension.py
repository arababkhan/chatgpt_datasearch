import pandas as pd
import time
import openai
from openai import OpenAI
import re
import keyboard
import random

file_path = 'OpenAI 3 Dimension Search.xlsx'
running = True

client = OpenAI(
    api_key=''  # Make sure to replace this with your actual API key
)

def retry_with_exponential_backoff(
    func,
    initial_delay: float = 1,
    exponential_base: float = 2,
    jitter: bool = True,
    max_retries: int = 10,
    errors: tuple = (openai.RateLimitError,),
):
    """Retry a function with exponential backoff."""
 
    def wrapper(*args, **kwargs):
        # Initialize variables
        num_retries = 0
        delay = initial_delay
 
        # Loop until a successful response or max_retries is hit or an exception is raised
        while True:
            try:
                return func(*args, **kwargs)
 
            # Retry on specific errors
            except errors as e:
                # Increment retries
                num_retries += 1
 
                # Check if max retries has been reached
                if num_retries > max_retries:
                    raise Exception(
                        f"Maximum number of retries ({max_retries}) exceeded."
                    )
 
                # Increment the delay
                delay *= exponential_base * (1 + jitter * random.random())
                print(f"Rate limit exceeded. Waiting for {delay} seconds before retrying...")
                # Sleep for the delay
                time.sleep(delay)
 
            # Raise exceptions for any errors not specified
            except Exception as e:
                raise e
 
    return wrapper

@retry_with_exponential_backoff
def completions_with_backoff(**kwargs):
    global client
    return client.chat.completions.create(**kwargs)

def getDimensions(excel_file):

    xls = pd.ExcelFile(excel_file, engine="openpyxl")   
    sheet_names = xls.sheet_names

    dimension_pattern = re.compile(r'\b(\d+)(mm)?\s*x\s*(\d+)(mm)?\s*x\s*(\d+)(mm)?\b')

    for sheet_name in sheet_names:

        df = pd.read_excel(xls, sheet_name=sheet_name)

        for index, row in df.iterrows():
            if (not pd.isna(row["Manufacturer"])) and (pd.isna(row["Width (mm)"])):  
                if type(row["Model"]) in [int, float, complex]:  
                    model = row["Model"]
                else:
                    model = row["Model"].strip()   
                manufacturer = row["Manufacturer"].strip()

                # Adjust the conversation to mimic a more natural inquiry, which could lead to better results
                messages = [
                    {"role": "system", "content": "You are a customer support chatbot. Use your knowledge base to best respond to customer queries."},
                    {"role": "user", "content": "I heard you can provide a wealth of information on various subjects. Could you tell me the dimensions of the {0} {1} {2} model in millimeters? I'm looking for its width, height, and depth in that order. Don't need any descript including text - I need only numbers like format [width] x [height] x [depth]".format(manufacturer, model, sheet_name)}
                ]

                try:
                    # Create a completion with the updated messages
                    time.sleep(2)
                    completion = completions_with_backoff(
                        model="gpt-4-turbo-preview",
                        messages=messages
                    )

                    for choice in completion.choices:
                        # Find all matches in the text
                        
                        matches = dimension_pattern.findall(choice.message.content)

                        formatted_matches = ['{} x {} x {}'.format(*match[::2]) for match in matches]
        
                        for matched_text in formatted_matches:
                            pattern_number = r'\d+'
                            matched_numbers = re.findall(pattern_number, matched_text)
                            numbers = [match_number for match_number in matched_numbers]
                            df.at[index, "Width (mm)"] = numbers[0]
                            df.at[index, "Depth (mm)"] = numbers[2]
                            df.at[index, "Height (mm)"] = numbers[1]

                except Exception as e:
                    print(f"An error occurred: {e}")
                    print("Please exit and run after one hour again!")
                    return 0

        with pd.ExcelWriter(excel_file, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)
        
        print(f"Completed {sheet_name} sheet")
    
    print(f"Completed All Sheets!")
    print("If you want to proceed more products, Please run program again after edited excel file.")

def monitor_esc_key(e):
    global running
    if e.event_type == keyboard.KEY_DOWN and e.name == 'esc':
        print("ESC key pressed. Exiting program...")
        running = False

keyboard.hook(monitor_esc_key)

def main():
    global running
    print("Running...")
    print("Press 'ESC' to exit the program.")
    getDimensions(file_path)

    while (running):
        time.sleep(1) 


if __name__ == "__main__":
    # Run the main function
    main()

    