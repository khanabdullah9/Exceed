import os
import pandas as pd
from langchain_core.messages import HumanMessage, ToolMessage, AIMessage
from pprint import pprint
from io import StringIO

from agent import agent
from da_agent import da_agent
from xl import display, open_book, close_book

# file_path = os.path.join("EXCEL_REPO", "file.xlsx")
# if not os.path.exists(file_path):
#     pd.DataFrame({}).to_excel(file_path)

# data_frame = pd.read_excel(file_path)
# workbook = open_book()

# while True:
#     prompt = input("(Press 'q' to quit): ")
#     if prompt.strip().lower() == "q":
#         close_book()
#         break
#     if prompt.strip().lower() == "cls":
#         continue

#     response = agent.invoke(dict(
#         messages = [HumanMessage(content = prompt)],
#         data_frame = data_frame,
#         file_path = os.path.join("EXCEL_REPO", "file.xlsx")
#     ))

    

#     # pprint(response)

#     response = da_agent.invoke(dict(
#             messages = [HumanMessage(content = "")],
#             file_path = os.path.join("EXCEL_REPO", "file.xlsx")
#         ))

#     if response:
#         for idx, msg in enumerate(response["messages"]):
#             if isinstance(msg, AIMessage):
#                 pprint(f"{idx}: AIMessage")
#                 pprint(msg.content)
#                 pprint(msg.additional_kwargs)
#             elif isinstance(msg, ToolMessage):
#                 pprint(f"{idx}: ToolMessage")
#                 pprint(msg.content)
#                 pprint(msg.name)
#             print("*"*100)


response = da_agent.invoke(dict(
            messages = [HumanMessage(content = "")],
            file_path = os.path.join("EXCEL_REPO", "Sales data.xlsx")
        ))

if response:
    for idx, msg in enumerate(response["messages"]):
        if isinstance(msg, AIMessage):
            pprint(f"{idx}: AIMessage")
            pprint(msg.content)
            pprint(msg.additional_kwargs)
        elif isinstance(msg, ToolMessage):
            pprint(f"{idx}: ToolMessage")
            pprint(msg.content)
            pprint(msg.name)
        print("*"*100)
