from openai import AsyncOpenAI
from openai.types.chat import ChatCompletionMessageParam, ChatCompletionToolParam
from mrai.agent.schema import Message, LLMResponse, ToolCall, Tool
from typing import List, Union, Sequence, cast, Iterable
from mrai.agent.llm.llm_config import LLMConfig
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
            return ToolCall(
                id=tool_call.id,
                type=tool_call.type,
                function=ToolCall.ToolCallFunction(
                    name=tool_call.function.name,
                    arguments=json.loads(tool_call.function.arguments),
                ),
                tool=tool
            )
        except json.JSONDecodeError as e:
            print(f"Invalid JSON in tool call arguments: {tool_call.function.arguments}")
            raise ValueError(f"Invalid JSON in tool call arguments: {e}")

    async def chat(self, messages: Sequence[Union[str, dict, Message]], tools: list[Tool] = []) -> Message:
        dict_messages: List[ChatCompletionMessageParam] = self.format_messages(messages)
        response = await self.client.chat.completions.create(
            model=self.config.model,
            messages=dict_messages,
            temperature=self.config.temperature,
            max_tokens=self.config.max_tokens,
            tools=cast(Iterable[ChatCompletionToolParam], [tool.to_dict() for tool in tools]) if tools else [],
            tool_choice="auto"
        )
        if not response.choices:
            raise ValueError("No response from OpenAI")
        
        tool_calls: list[ToolCall] = []
        content: str = ""
        if not response.choices[0]:
            raise ValueError("No response from OpenAI")
        # default use the first choice
        choice = response.choices[0]
        if choice.message.content:
            content = choice.message.content
        if choice.message.tool_calls and tools:
            tool_calls.extend(
                self._process_tool_call(tool_call, tools)
                for tool_call in choice.message.tool_calls
            )
        
        assistant_message = Message(
            role=choice.message.role,
            content=content,
            tool_calls=tool_calls
        )
        return assistant_message