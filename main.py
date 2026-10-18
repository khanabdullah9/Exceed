import os
import pandas as pd
from langchain_core.messages import HumanMessage, ToolMessage
from pprint import pprint
from io import StringIO

from agent import agent
from xl import display, open_book

file_path = os.path.join("EXCEL_REPO", "file.xlsx")
if not os.path.exists(file_path):
    pd.DataFrame({}).to_excel(file_path)

data_frame = pd.read_excel(file_path)
workbook = open_book()

while True:
    prompt = input("(Press 'q' to quit): ")
    if prompt.strip().lower() == "q":
        break
    if prompt.strip().lower() == "cls":
        continue

    response = agent.invoke(dict(
        messages = [HumanMessage(content = prompt)],
        data_frame = data_frame
    ))
    # pprint(response)

    if response:
        for msg in response["messages"][::-1]:
            pprint(msg)
