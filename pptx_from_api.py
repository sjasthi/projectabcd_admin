import requests
import pandas as pd

# Make a GET request to the API
response = requests.get("https://jsonplaceholder.typicode.com/todos")

# Check if the request was successful
if response.status_code == 200:
    # Get the JSON data from the response
    data = response.json()

    # Create a DataFrame from the JSON data
    df = pd.DataFrame(data)

    # Display the DataFrame
    print(df)
else:
    print("Error occurred while accessing the API:", response.status_code)
