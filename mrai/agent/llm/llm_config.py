from pydantic import Field
from typing import Literal
from openai import BaseModel


class LLMConfig(BaseModel):

    base_url: str = Field(default="https://api.openai.com/v1", description="The base URL for the OpenAI API")
    api_key: str = Field(..., description="The API key for the OpenAI API")
    model: str = Field(..., description="The model to use for the OpenAI API")
    temperature: float = Field(default=0.5, description="The temperature for the OpenAI API")
    max_tokens: int = Field(default=1000, description="The maximum number of tokens for the OpenAI API")
    top_p: float = Field(default=1.0, description="The top P for the OpenAI API")
    frequency_penalty: float = Field(default=0.0, description="The frequency penalty for the OpenAI API")
    api_type: Literal["openai", "azure"] = Field(default="openai", description="The type of API to use")
    reasoning_effort: Literal["low", "medium", "high"] = Field(default="medium", description="The effort for the OpenAI API")
    
    
