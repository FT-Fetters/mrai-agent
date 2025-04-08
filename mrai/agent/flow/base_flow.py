

from abc import ABC, abstractmethod

from mrai.agent.agent import Agent
from mrai.agent.schema import FlowInput
from mrai.agent.tool.assign_agent_tool import AssignAgent
from mrai.agent.tool.terminate_tool import Terminate


class BaseFlow(ABC):
    def __init__(self, agents: dict[str, Agent]):
        """
        Args:
            agents: A dictionary of agents with their names as keys and Agent objects as values, the primary agent should be the key "primary"
        >>> agent_list = {
        ...    "primary": Agent() # primary agent
        ...    "other": Agent(), # other agents
        ... }
        >>> flow = BaseFlow(agents)
        """
        if "primary" not in agents:
            raise ValueError("Primary agent is required")
        self.agents = agents
        for agent in self.agents.values():
            # add terminate tool to all agents
            agent.tools.append(Terminate())
            # add assign agent tool to all agents
            agent.tools.append(AssignAgent())

    @abstractmethod
    def run(self, input: FlowInput):
        """Run the flow"""
