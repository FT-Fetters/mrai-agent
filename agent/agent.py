from abc import ABC, abstractmethod
from typing import Optional

from pydantic import BaseModel, Field

from agent.schema import Memory, Message


class Agent(BaseModel, ABC):
    memory: Memory = Field(default_factory=Memory)

    def __init__(self, memory: Optional[Memory] = None):
        if memory:
            self.memory = memory
        else:
            self.memory = Memory()

    @abstractmethod
    def action(self):
        """"""

    def add_observation(
            self,
            observation: str
    ):
        """Add an observation to the memory"""
        self.memory.add_message(Message(role="system", content=observation))
