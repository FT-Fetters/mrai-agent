
from openai import AsyncOpenAI
from agent.schema import Message, LLMResponse
from typing import List, Union
from agent.llm.llm_config import LLMConfig


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
    def format_messages(messages: List[Union[str, dict, Message]]):
        formatted_messages = []
        for message in messages:
            if isinstance(message, str):
                formatted_messages.append({"role": "user", "content": message})
            elif isinstance(message, dict):
                if "role" in message and "content" in message:
                    formatted_messages.append(message)
                else:
                    raise ValueError("Invalid message format, missing role or content")
            elif isinstance(message, Message):
                formatted_messages.append(message.to_dict())
            else:
                raise ValueError("Invalid message format")
        return formatted_messages


    async def chat(self, messages: List[Union[str, dict, Message]]):
        messages = self.format_messages(messages)
        response = await self.client.chat.completions.create(
            model=self.config.model,
            messages=messages,
            temperature=self.config.temperature,
            max_tokens=self.config.max_tokens,
        )
        if not response.choices or not response.choices[0].message.content:
            raise ValueError("No response from OpenAI")
        return LLMResponse(
            content=response.choices[0].message.content,
            tool_calls=[],
        )
