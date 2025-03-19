

from abc import ABC, abstractmethod

from openai import BaseModel

from agent.agent import Agent
from agent.schema import FlowInput


class BaseFlow(BaseModel, ABC):
    def __init__(self, agents: dict[str, Agent]):
        """
        Args:
            agents: A dictionary of agents with their names as keys and Agent objects as values, the primary agent should be the key "primary"
        >>> agent_list = {
        ...    "primary": Agent() # primary agent
        ...    "other": Agent(), # other agents
        ... }
        >>> flow = AgentFlow(agents)
        """
        if "primary" not in agents:
            raise ValueError("Primary agent is required")
        self.agents = agents

    @abstractmethod
    def run(self, input: FlowInput):
        """Run the flow"""
