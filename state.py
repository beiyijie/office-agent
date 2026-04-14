from __future__ import annotations

from typing import Any, Dict, List, Optional, TypedDict


class AgentState(TypedDict):
    task: str
    messages: List[Dict[str, Any]]
    context: Dict[str, Any]
    current_node: str
    tools_results: Dict[str, Any]
    error: Optional[str]
    retry_count: int
    next_node: str
    response: str
    pending_actions: List[str]
    completed_actions: List[str]
    intermediate_results: Dict[str, Any]


def create_initial_state(user_message: str) -> AgentState:
    return AgentState(
        task=user_message,
        messages=[{"role": "user", "content": user_message}],
        context={},
        current_node="router",
        tools_results={},
        error=None,
        retry_count=0,
        next_node="router",
        response="",
        pending_actions=[],
        completed_actions=[],
        intermediate_results={},
    )
