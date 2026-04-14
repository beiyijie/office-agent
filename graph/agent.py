from __future__ import annotations

from typing import Any

from office_agent.graph.edges import route_to_node
from office_agent.graph.nodes import document_node, email_node, file_ops_node, general_node, meeting_node, router_node
from office_agent.state import AgentState

try:
    from langgraph.graph import END, StateGraph
except ImportError:  # pragma: no cover
    END = "__end__"
    StateGraph = None


class OfficeAgent:
    def __init__(self) -> None:
        self.max_iterations = 10
        self._graph = self._build_graph() if StateGraph is not None else None

    def _build_graph(self) -> Any:
        workflow = StateGraph(AgentState)
        workflow.add_node("router", router_node)
        workflow.add_node("email_node", email_node)
        workflow.add_node("document_node", document_node)
        workflow.add_node("file_ops_node", file_ops_node)
        workflow.add_node("general_node", general_node)
        workflow.add_node("meeting_node", meeting_node)
        workflow.set_entry_point("router")

        workflow.add_conditional_edges(
            "router",
            route_to_node,
            {
                "email_node": "email_node",
                "document_node": "document_node",
                "file_ops_node": "file_ops_node",
                "general_node": "general_node",
                "meeting_node": "meeting_node",
                "END": END,
            },
        )

        for node in ("email_node", "document_node", "file_ops_node", "general_node", "meeting_node"):
            workflow.add_edge(node, "router")

        return workflow.compile()

    def invoke(self, state: AgentState) -> AgentState:
        state = router_node(state)
        iterations = 0

        while iterations < self.max_iterations:
            iterations += 1
            next_node = state.get("next_node", "END")
            if next_node == "END":
                break

            handlers = {
                "email_node": email_node,
                "document_node": document_node,
                "file_ops_node": file_ops_node,
                "general_node": general_node,
                "meeting_node": meeting_node,
            }
            handler = handlers.get(next_node)
            if not handler:
                break

            state = handler(state)
            state["completed_actions"] = [*state.get("completed_actions", []), next_node]
            state = router_node(state)

        return state
