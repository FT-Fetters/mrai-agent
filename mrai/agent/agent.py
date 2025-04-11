from abc import ABC, abstractmethod
from typing import Generator, Iterator, Optional, Any, Union, AsyncIterator

from pydantic import Field, BaseModel, ConfigDict

from mrai.agent.llm.llm import LLM
from mrai.agent.schema import Callback, LLMResponse, Memory, Message, Tool, ToolCall

from loguru import logger

class Agent(ABC, BaseModel):
    model_config = ConfigDict(arbitrary_types_allowed=True)

    memory: Memory = Field(default_factory=Memory)
    llm: LLM = Field(..., description="The LLM of the agent")
    tools: list[Tool] = Field(default=[], description="The tools of the agent")
    callbacks: list[Callback] = Field(default=[], description="The callbacks of the agent")
    name: str = Field(default="Agent")
    
    @abstractmethod
    async def action(self) -> Union[tuple[Optional[str], list[ToolCall]], AsyncIterator[str]]:
        """The action of the agent"""
        pass
    
    @abstractmethod
    def add_observation(self, observation: str) -> None:
        """Add an observation to the memory"""
        pass
    
    @abstractmethod
    def add_user_message(self, message: str) -> None:
        """Add a user message to the memory"""
        pass


class SimpleAgent(Agent):
    def __init__(
        self,
        llm: LLM,
        prompt: str,
        tools: Optional[list[Tool]] = None,
        memory: Optional[Memory] = None,
        callbacks: Optional[list[Callback]] = None,
        name: Optional[str] = None,
    ):
        super().__init__(
            llm=llm,
            memory=memory or Memory(),
            tools=tools or [],
            callbacks=callbacks or [],
            name=name or "Agent"
        )
        
        # add system prompt
        self.memory.add_message(Message(role="system", content=prompt))

    async def action(self) -> tuple[Optional[str], list[ToolCall]]:
        """The action of the agent"""
        messages_for_llm = self.memory.messages.copy()
        assistant_message: Message = await self.llm.chat(messages=messages_for_llm, tools=self.tools)
        # add the response to the memory
        self.memory.add_message(assistant_message)
        if assistant_message.content:
            logger.info(f"🤔 「{self.name}」 thought: {assistant_message.content}")
        
        if assistant_message.tool_calls:
            for tool_call in assistant_message.tool_calls:
                logger.info(f"🔧 「{self.name}」 called tool: {tool_call.function.name}")
        
        return assistant_message.content, assistant_message.tool_calls

    def add_observation(self, observation: str) -> None:
        """Add an observation to the memory"""
        self.memory.add_message(Message(role="system", content=observation))

    def add_user_message(self, message: str) -> None:
        """Add a user message to the memory"""
        self.memory.add_message(Message(role="user", content=message))


class RealtimeCallAgent(Agent):
    
    user_input: str = Field(default="")
    prompt: str = Field(default="")
    
    def __init__(
        self,
        llm: LLM,
        prompt: str,
        tools: Optional[list[Tool]] = None,
        memory: Optional[Memory] = None,
        callbacks: Optional[list[Callback]] = None,
        name: Optional[str] = None,
    ):
        super().__init__(
            llm=llm,
            memory=memory or Memory(),
            tools=tools or [],
            callbacks=callbacks or [],
            name=name or "Agent"
        )
        self.prompt = prompt
        
    async def action(self) -> AsyncIterator[str]:
        """
        ### Realtime call agent returns a stream of chunks, each chunk is a string

        the chunk is formatted as follows:
            <chunk_type>::<chunk_content>
            
            the chunk_type can be one of the following:
            * content: the content of the chunk
            * reasoning_content: the reasoning content of the chunk
        """
        messages_for_llm = self.memory.messages.copy()
        async for chunk in self.llm.stream_chat(
            messages=messages_for_llm,
            tools=self.tools,
            flag=True
        ):
            yield chunk

    
    def add_observation(self, observation: str) -> None:
        # forbid add observation to realtime call agent
        raise NotImplementedError("Realtime call agent does not support add observation")
    
    def add_user_message(self, message: str) -> None:
        self.user_input = message
        
    def set_memory(self, memory: Memory) -> None:
        """Set the memory of the agent"""
        self.memory = memory
