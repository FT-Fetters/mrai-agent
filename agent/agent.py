from abc import ABC
from typing import Optional

from pydantic import Field

from agent.llm.llm import LLM
from agent.schema import Callback, LLMResponse, Memory, Message, Tool
from agent.tool.terminate_tool import Terminate

from loguru import logger

class Agent(ABC):
    memory: Memory = Field(default_factory=Memory)
    llm: LLM = Field(..., description="The LLM of the agent")
    tools: list[Tool] = Field(default=[], description="The tools of the agent")
    callbacks: list[Callback] = Field(default=[], description="The callbacks of the agent")
    
    def __init__(
        self,
        llm: LLM,
        prompt: str,
        tools: Optional[list[Tool]] = None,
        memory: Optional[Memory] = None,
        callbacks: Optional[list[Callback]] = None,
        name: Optional[str] = None,
    ):
        self.llm = llm
        
        if memory:
            self.memory = memory
        else:
            self.memory = Memory()
        # add system prompt
        self.memory.add_message(Message(role="system", content=prompt))
        
        if tools:
            self.tools = tools
        else:
            self.tools = []
        
        if callbacks:
            self.callbacks = callbacks
        else:
            self.callbacks = []
        
        self.name = name if name else "Agent"
        

    async def action(self):
        """The action of the agent"""
        messages_for_llm = self.memory.messages.copy()
        response: LLMResponse = await self.llm.chat(messages=messages_for_llm, tools=self.tools)
        if response.content:
            self.memory.add_message(Message(role="assistant", content=response.content))
            logger.info(f"🤔 「{self.name}」 thought: {response.content}")
        
        if response.tool_calls:
            for tool_call in response.tool_calls:
                logger.info(f"🔧 「{self.name}」 called tool: {tool_call.function.name}")
                # tool_call.function.execute(**tool_call.arguments)
        
        return response.content, response.tool_calls

    def add_observation(
            self,
            observation: str
    ):
        """Add an observation to the memory"""
        self.memory.add_message(Message(role="system", content=observation))

    def add_user_message(
            self,
            message: str
    ):
        """Add a user message to the memory"""
        self.memory.add_message(Message(role="user", content=message))
