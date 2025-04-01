from openai import AsyncOpenAI
from openai.types.chat import ChatCompletionMessageParam
from agent.schema import Message, LLMResponse, ToolCall, Tool
from typing import List, Union, Sequence, cast
from agent.llm.llm_config import LLMConfig
import json


class LLM:

    def __init__(self, config: LLMConfig):
        # validate config
        if config.api_key is None or config.api_key == "":
            raise ValueError("api_key is required")
        if config.base_url is None or config.base_url == "":
            config.base_url = "https://api.openai.com/v1"
        if config.model is None or config.model == "":
            raise ValueError("model is required")
        self.config = config

        self.client = AsyncOpenAI(
            api_key=config.api_key,
            base_url=config.base_url,
        )

    @staticmethod
    def format_messages(messages: Sequence[Union[str, dict, Message]]) -> List[ChatCompletionMessageParam]:
        formatted_messages: List[ChatCompletionMessageParam] = []
        for message in messages:
            if isinstance(message, str):
                formatted_messages.append({"role": "user", "content": message})
            elif isinstance(message, dict):
                if "role" in message and "content" in message:
                    formatted_messages.append(cast(ChatCompletionMessageParam, message))
                else:
                    raise ValueError("Invalid message format, missing role or content")
            elif isinstance(message, Message):
                formatted_messages.append(cast(ChatCompletionMessageParam, message.to_dict()))
            else:
                raise ValueError("Invalid message format")
        return formatted_messages

    @staticmethod
    def _process_tool_call(tool_call, tools: List[Tool]) -> ToolCall:
        """Process a single tool call and return a ToolCall object"""
        tool = next((tool for tool in tools if tool.name == tool_call.function.name), None)
        if not tool:
            raise ValueError(f"Tool {tool_call.function.name} not found")
        try:
            arguments = json.loads(tool_call.function.arguments)
        except json.JSONDecodeError:
            arguments = {}
        return ToolCall(
            id=tool_call.id,
            type=tool_call.type,
            function=tool,
            arguments=arguments
        )

    async def chat(self, messages: Sequence[Union[str, dict, Message]], tools: list[Tool] = []) -> LLMResponse:
        dict_messages: List[ChatCompletionMessageParam] = self.format_messages(messages)
        response = await self.client.chat.completions.create(
            model=self.config.model,
            messages=dict_messages,
            temperature=self.config.temperature,
            max_tokens=self.config.max_tokens,
            tools=[tool.to_dict() for tool in tools]
        )
        if not response.choices:
            raise ValueError("No response from OpenAI")
        
        tool_calls: list[ToolCall] = []
        content: str = ""
        for choice in response.choices:
            if choice.message.content:
                content = choice.message.content
            if choice.message.tool_calls and tools:
                tool_calls.extend(
                    self._process_tool_call(tool_call, tools)
                    for tool_call in choice.message.tool_calls
                )
        
        return LLMResponse(
            content=content,
            tool_calls=tool_calls,
        )
