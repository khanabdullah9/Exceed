from dotenv import load_dotenv
import json
import os
import logging
from typing import Union
import pythoncom
from functools import wraps
import numpy as np
import pandas as pd

# 1. Setup the Root Logger to a high level (to silence the server/libraries)
logging.basicConfig(level=logging.WARNING) 

# 2. Create a specific logger for YOUR app
logger = logging.getLogger("my_app")
logger.setLevel(logging.INFO) # Your app will log INFO and above
logger.propagate = False # This prevents logs from "bubbling up" to the root logger

# 3. Add a FileHandler specifically to your logger
log_path = os.path.join(os.path.dirname(__file__), "app.log")
file_handler = logging.FileHandler(log_path, encoding="utf-8")
formatter = logging.Formatter("{asctime} - {levelname} - {message}", style="{", datefmt="%Y-%m-%d %H:%M")
file_handler.setFormatter(formatter)
logger.addHandler(file_handler)

def get_api_key():
    load_dotenv()
    try:
        return os.getenv("API_KEY")
    except Exception:
        return ""

def read_json(file_name: str):
    json_data = ""
    with open(file_name, "r") as data:
        json_data = json.load(data)
    return json_data

def read_app_config():
    return read_json("app_config.json")

def get_model_name():
    json_data = read_app_config()
    return json_data["model_name"]

def save_png(agent):
    png_data = agent.get_graph().draw_mermaid_png()
    with open(os.path.join("graph_pngs","graph.png"),"wb") as png:
        png.write(png_data)
    print("Image saved!")

def log_err(err_msg: Union[Exception|str]):
    logger.error(err_msg)

def log_agent_err(err_msg: Union[Exception|str]):
    logger.error(str(err_msg)+" [source: agent]")

def log_xl_err(err_msg: Union[Exception|str]):
    logger.error(str(err_msg)+" [source: xl]")

def log_info(txt: str):
    logger.info(txt)

def init_com(func):
    @wraps(func)
    def wrapper(*args, **kwargs):
        pythoncom.CoInitialize()
        try:
            return func(*args, **kwargs)
        finally:
            pythoncom.CoUninitialize()
    return wrapper

def parse_lst(str_lst: str) -> list[str]:
    """Convert agent's input str into a python list

    Args:
        str_lst (str): python str repr of list

    Returns:
        list[str]: parsed list
    """
    parsed_lst = []
    try:
        parsed_lst = str_lst.split(",")
    except Exception as err:
        pass

    try:
        parsed_lst = str_lst.remove(",").split("\n   ")
    except Exception as err:
        pass

    # print(f"parsed list {parsed_lst}")
    return parsed_lst

def parse_number_arr(lst: list[str]) -> np.array:
    try:
        # return np.array(lst, dtype = "float")
        return pd.Series(lst, dtype = "float")
    except Exception:
        pass

def calc_arr_mean(lst: list[any]) -> float:
    sum = 0
    count = 0
    for elem in lst:
        if isinstance(elem, int) or isinstance(elem, float):
            sum += elem
            count += 1
    return sum / count if count > 0 else 0

def replace_missing(lst: list[any], mean: float) -> list[any]:
    return list(map(lambda x: mean if x == None else x, lst))