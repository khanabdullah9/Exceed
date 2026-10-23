from langgraph.graph import StateGraph, START, END
from typing import TypedDict, Annotated, Sequence, Union
from langchain_core.messages import AnyMessage, SystemMessage, AIMessage
from langgraph.graph.message import add_messages
from langchain_groq import ChatGroq
from langgraph.prebuilt import ToolNode, InjectedState
from langchain.tools import tool
import matplotlib.pyplot as plt
import xlwings as xw

from utils import init_com, log_agent_err, log_info, get_model_name, get_api_key, write_todo, read_todo, save_png, create_todo_path
from xl import open_book, display, plot_line
import data

llm = ChatGroq(
    model = get_model_name(),
    api_key = get_api_key()
)

class da_AgentState(TypedDict):
    messages: Annotated[Sequence[AnyMessage], add_messages]
    file_path: str

# DISPLAY_BOOK = xw.Book()

#tools
@tool
@init_com
def get_col_names(state: Annotated[da_AgentState, InjectedState]) -> dict[str, list[str]]:
    """Returns the column names of data frames in each sheet

    Args:
        state (Annotated[da_AgentState, InjectedState]): agent's state

    Returns:
        dict[str, list[str]]: sheet_name: list of column names
    """
    col_names = {}
    book = open_book(state["file_path"])

    for sheet_name in book.sheet_names:
        sheet = book.sheets[sheet_name]
        df = sheet.range("A1").expand("table").value
        df = data.create_df(df)
        col_names[sheet_name] = df.columns.to_list()

    return col_names

@tool
@init_com
def get_summary_stats(sheet_name: str, state: Annotated[da_AgentState, InjectedState], col_name: str = "") -> bool:
    """Returns the summary statistics of the dataset

    Args:
        sheet_name (str): sheet's name
        state (Annotated[da_AgentState, InjectedState]): agent's state
        col_name (str, optional): column name from the sheet. Defaults to "".

    Returns:
        bool: _description_
    """
    try:
        book = open_book(file_path = state["file_path"])
        sheet = book.sheets[sheet_name]

        df = sheet.range("A1").expand("table").value
        if not data:
            return False

        df = data.create_df(df)
        # if col_name:
        if False:
            display(df[col_name].describe(), book = book, sheet_name = "summary stats")
        else:
            display(df.describe(), book = book, sheet_name = "summary stats")

        return True
    except Exception as err:
        log_agent_err(err)
        return False

@tool
@init_com
def group_aggregate(sheet_name: str, state: Annotated[da_AgentState, InjectedState], group_col: Union[str|list[str]], agg_col: str, agg_type: str) -> bool:
    """Groups the dataframe on the basis of column/s and generates the aggregate for other column

    Args:
        sheet_name (str): excel sheet name
        state (Annotated[da_AgentState, InjectedState]): sheet' name
        group_col (str): column to be grouped on
        agg_col (str): column to be aggregated
        agg_type (str): mean, median, count, min, max

    Returns:
        bool: acknowledgement
    """
    try:
        book = open_book(file_path = state["file_path"])
        sheet = book.sheets[sheet_name]

        df = sheet.range("A1").expand("table").value
        if not data:
            return False

        df = data.create_df(df)
        grouped_df = df.groupby(by = group_col)

        match agg_type:
            case "median":
                display(grouped_df[agg_col].median(), book = book, sheet_name = f"median {agg_col} per {group_col}")
            case "count":
                display(grouped_df[agg_col].count(), book = book, sheet_name = f"count {agg_col} per {group_col}")
            case "min":
                display(grouped_df[agg_col].min(), book = book, sheet_name = f"min {agg_col} per {group_col}")
            case "max":
                display(grouped_df[agg_col].max(), book = book, sheet_name = f"max {agg_col} per {group_col}")
            case _:
                display(grouped_df[agg_col].mean(), book = book, sheet_name = f"mean {agg_col} per {group_col}")

        return True
    except Exception as err:
        log_agent_err(err)
        return False

@tool
@init_com
def plot_line_chart(sheet_name: str, state: Annotated[da_AgentState, InjectedState], col_x: str, col_y: str) -> bool:
    """Adds a line chart in the same sheet

    Args:
        sheet_name (str): excel sheet name
        state (Annotated[da_AgentState, InjectedState]): agent's state
        col_x (str): x-axis name
        col_y (str): y-axis name

    Returns:
        bool: acknowledgement
    """
    try:
        book = open_book()
        plot_line(col_x, col_y, book = book, sheet_name = f"plot-{col_x}-{col_y}")
    except Exception as err:
        log_agent_err(err)
        return False

tools = [
    get_col_names,
    get_summary_stats,
    group_aggregate,
    plot_line_chart]
llm_with_tools = llm.bind_tools(tools)
tool_desc = [
    "get_summary_stats: displays the summary statistics of column or entire dataset",
    "group_aggregate: groups dataset and calls aggregate functions",
    "plot_line_chart: displays a line chart",
    "get_col_names: Returns the column names of data frames in each sheet"
]

#Nodes
def create_todo(state: da_AgentState) -> da_AgentState:
    
    try:
        book = open_book(state["file_path"])
        todo_fname = create_todo_path(state["file_path"])
        content = []

        for sheet_name in book.sheet_names:
            df = data.create_df(book.sheets[sheet_name].range("A1").expand(mode = "table").value)
            content.append(f"sheet: {sheet_name} | columns: {df.columns.to_list()} \n")
        # write_todo(todo_fname, content)

        sys_msg = SystemMessage(content = f"""
            You are a data analysis planning agent. Based on the provided worksheet names and column names, create a concise, actionable plan for another agent to execute.
    
            - Return only bulleted steps in a logical sequence. 
            - Avoid giving headers to the tasks. Simple instruction that can be converted to tool call. Do not follow header-subheader format
            - Include relevant data validation, cleaning, EDA, statistical analysis, visualizations, and KPIs where applicable.
            - Specify the columns and operations involved whenever possible.
            - Consider relationships between worksheets only when supported by the available information.
            - Do not assume data values, invent columns, or make unsupported business assumptions.
            - If a column's meaning is unclear, instruct the agent to inspect its values before proceeding.
            - Prioritize meaningful, dataset-specific analysis over generic or redundant steps.
    
            Here are the sheet name/s and column name/s:
            {"\n".join(content)}

            Here are the list of available tools:
            {"\n".join(tool_desc)}

            Output only the analysis steps, with no introduction or conclusion.
                                                      """)

        state["messages"].append(llm.invoke([sys_msg]))
        if state["messages"]:
            for msg in state["messages"]:
                if isinstance(msg, AIMessage):
                        write_todo(todo_fname, msg.content)
            
        return state
    except Exception as err:
        log_agent_err(err)
        return state

def llm_infer(state: da_AgentState) -> da_AgentState:
    sys_msg = SystemMessage(content = f"""
             You are a data analysis agent. Execute the analysis steps provided below using only the available tools.

            **Instructions:**
            - Follow the steps in order.
            - For each step, call the most appropriate available tool.
            - If no suitable tool exists for a step, skip it.
            - Do not invent tool names, arguments, data, or results.
            - Use tool outputs to guide subsequent tool calls when relevant.
            - Do not answer questions, explain your reasoning, or generate text outside tool calls.
            - Stop when all executable steps have been completed.

            **Analysis Steps:**
            {read_todo(state["file_path"])}

            **Available Tools:**
            {"\n".join(tool_desc)}
                                           """)

    try:
        state["messages"].append(llm_with_tools.invoke([sys_msg]+state["messages"]))
        return state
    except Exception as err:
        log_agent_err(err)
        return state

def should_continue(state: da_AgentState) -> da_AgentState:
    if not state["messages"]:
        return END
    
    last_message = state["messages"][-1]
    return "tool_node" if last_message.tool_calls else END

graph = StateGraph(da_AgentState)

graph.add_node("create_todo", create_todo)
graph.add_node("llm_infer", llm_infer)
graph.add_node("tool_node", ToolNode(tools = tools))

graph.set_entry_point("create_todo")
graph.add_edge("create_todo", "llm_infer")
graph.add_edge("llm_infer", "tool_node")
graph.add_edge("tool_node", "llm_infer")
graph.add_edge("llm_infer", END)
# graph.add_edge("create_todo", END)

da_agent = graph.compile()

# save_png(da_agent)

