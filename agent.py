from typing import TypedDict, Sequence, Annotated, Union
from langgraph.graph import StateGraph, START, END
from langgraph.graph.message import add_messages
from langchain_core.messages.utils import AnyMessage, SystemMessage, AIMessage
from langgraph.prebuilt import ToolNode, InjectedState
from langchain.tools import tool
from langchain_groq import ChatGroq
import xlwings as xw
import os
import pythoncom

import pandas as pd
from utils import get_api_key, get_model_name, save_png, log_agent_err, init_com
import xl

FILE_PATH = os.path.join("EXCEL_REPO", "file.xlsx")

llm = ChatGroq(
    model = get_model_name(),
    api_key = get_api_key()
)

class AgentState(TypedDict):
    messages: Annotated[Sequence[AnyMessage], add_messages]
    data_frame: pd.DataFrame

#tools section
## data cleaning
@tool
def handle_missing_values(col_name: str, state: Annotated[AgentState, InjectedState]) -> None:
    """Handle missing values in a column

    Args:
        col_name (str): column name,
        state (Annotated[AgentState, InjectedState]): agent's state
    """
    df = state["data_frame"]
    if df is not None:
        df[col_name] = df[col_name].fillna(df[col_name].mean())

@tool
def handle_date_format(col_name: str, format: str, state: Annotated[AgentState, InjectedState]) -> None:
    """Standardise the date format for the date related column

    Args:
        col_name (str): column name,
        format (str): python time strftime format
        state (Annotated[AgentState, InjectedState]): agent's state
    """
    df = state["data_frame"]
    try:
        if df is not None:
            df[col_name] = pd.to_datetime(df[col_name], format = format)
    except Exception:
        pass


## Analyse
@tool
def get_summary_stats(state: Annotated[AgentState, InjectedState], col_name: str = "") -> None:
    """Returns descriptive statistics of a column or the entire dataset

    Args:
        state (Annotated[AgentState, InjectedState]): agent's state
        col_name (str, optional): column name. Defaults to "".
    """
    df = state["data_frame"]
    if df is not None:
        xl.display(df[col_name].describe()) if col_name else xl.display(df.describe())


@tool
@init_com
def drop_column(
    col_name: str,
    state: Annotated[AgentState, InjectedState],
    sheet_name: str = ""
) -> bool:
    """Delete a column completely

    Args:
        sheet_name (str): Excel sheet name
        col_name (str): column name
        state (Annotated[AgentState, InjectedState]): agent's state

    Returns:
            bool: acknowledgement
    """

    try:
        df = state["data_frame"]

        if df is None or col_name not in df.columns:
            return False

        xl_col_idx = list(df.columns).index(col_name) + 1
        print(f"Column index: {xl_col_idx}")

        book = xl.open_book()

        sheet = (
            book.sheets[sheet_name]
            if sheet_name
            else book.sheets[0]
        )

        print(f"Sheet status: {sheet is not None}")

        xl.delete_column(sheet, xl_col_idx)
        return True
    except Exception as err:
        log_agent_err(err)
        return False


@tool
@init_com
def add_column(col_name: str,data: str, state: Annotated[AgentState, InjectedState], sheet_name: str = "") -> bool:
    """Add a new column

    Args:
        col_name (str): column name
        data (str): New data for the column (value at index 0 should be the column name). Keep the elements seperated by commas Eg: [1,2,3,4]
        state (Annotated[AgentState, InjectedState]): agent's state
        sheet_name (str): Name of the sheet (optional)

    Returns:
        bool: acknowledgement
    """
    try:
        df = state["data_frame"]
        if df is not None:
            # df[col_name] = data
            new_col_idx = len(df.columns)+1 # excel has 1 based indexing

            book = xl.open_book()
            sheet = book.sheets[sheet_name] if sheet_name else book.sheets[0]

            xl.add_column(sheet, data, new_col_idx)
            return True
        return False
    except Exception as err:
        log_agent_err(err)
        return False
            

@tool
def drop_row(idx: Union[int, float, str], sheet_name: str, state: Annotated[AgentState, InjectedState]) -> bool:
    """_summary_

    Args:
        idx (Union[int, float, str]): row index
        state (Annotated[AgentState, InjectedState]): agent's state

    Returns:
        bool: acknowledgement
    """
    try:
        book = xl.open_book()
        sheet = book.sheets[sheet_name] if sheet_name else book.sheets[0]

        return xl.del_row(sheet, idx)
    except Exception as err:
        log_agent_err(err)
        return False

@tool
def add_row(data: list[list], state: Annotated[AgentState, InjectedState]) -> bool:
    """Add a new row at the bottom

    Args:
        data (pd.DataFrame): data to be concatenated
        state (Annotated[AgentState, InjectedState]): agent's state

    Returns:
        bool: acknowledgement
    """
    df = state["data_frame"]
    if df is not None:
        state["data_frame"] = pd.concat(df, data)
        return True
    return False

## write content
@tool
def write_data(state: Annotated[AgentState, InjectedState]) -> bool:
    """Write the modified data back to the excel file

    Args:
        state (Annotated[AgentState, InjectedState]): agent's state

    Returns:
        bool: acknowledgement
    """
    df = state["data_frame"]
    if df is not None:
        df.to_excel(FILE_PATH)
        return True
    return False

tool_list = [add_row, 
             drop_row, 
             add_column, 
             drop_column, 
             get_summary_stats,
             handle_date_format, 
             handle_missing_values,
             write_data]
llm_with_tools = llm.bind_tools(tool_list)

def llm_invoke(state: AgentState) -> AgentState:
    sys_msg = SystemMessage(content = """
                You are an helpful assistant!
                You are designed to automate excel operations.
                Invoke the LLM and decide which tool to invoke.

                Tool list:
                1. add_row: add new row in the table, 
                2. drop_row: remove row in the table, 
                3. add_column: add new column in the table, 
                4. drop_column: remove column in the table, 
                5. get_summary_stats: get descriptive statistics,
                6. handle_date_format: standardise the date format, 
                7. handle_missing_values: fill missing values
                8. write_data: write the modified content back to the same file

                Never respond to 'out of context' questions/queries
                                      """)
    state["messages"].append(
        llm_with_tools.invoke([sys_msg] + state["messages"])
    )
    return state

def should_continue(state: AgentState) -> AgentState:
    last_msg = state["messages"][-1]
    if isinstance(last_msg, AIMessage):
        if last_msg.tool_calls:
            return "tool_node"
    return END

graph = StateGraph(AgentState)

graph.add_node("llm_invoke", llm_invoke)
graph.add_node("tool_node", ToolNode(tools = tool_list))

graph.set_entry_point("llm_invoke")
graph.add_conditional_edges("llm_invoke", 
                            should_continue,
                            {
                                "tool_node": "tool_node",
                                END: END
                            })
graph.add_edge("llm_invoke", END)
graph.add_edge("tool_node", END)

agent = graph.compile()


# save_png()