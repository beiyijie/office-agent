from __future__ import annotations

from office_agent.state import AgentState


def route_to_node(state: AgentState) -> str:
    return state["next_node"]
