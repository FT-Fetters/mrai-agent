
from pydantic import Field
from typing import List, Literal, Union
from openai import BaseModel

from agent.agent import Agent


class Message(BaseModel):

    role: Literal["system", "user", "assistant", "tool"] = Field(..., description="The role of the message")
    content: str = Field(..., description="The content of the message")



    def to_dict(self, **kwargs):
        return {
            "role": self.role,
            "content": self.content
        }

class FlowInput(BaseModel):
    """The input of the flow"""

    text: str = Field(..., description="The text to be processed")

    context: Union[list[Message], str] = Field(..., description="The context of the flow")

    def get_valid_input(self) -> Union[str, None]:
        """Get the valid input of the flow"""

        if self.text:
            return self.text
        else:
            return None


class Tool(BaseModel):

    class ToolParameter(BaseModel):
        name: str = Field(..., description="The name of the parameter")
        description: str = Field(..., description="The description of the parameter")
        type: Literal["string", "number", "boolean", "array", "object", "enum"] = Field(..., description="The type of the parameter")
        enum: list[str] = Field(..., description="The enum of the parameter")
        required: bool = Field(..., description="Whether the parameter is required")

    name: str = Field(..., description="The name of the tool")
    description: str = Field(..., description="The description of the tool")
    parameters: dict[str, ToolParameter] = Field(..., description="The parameters of the tool")



class ToolCall(BaseModel):
    """Represents a tool/function call in a message"""

    id: str = Field(..., description="The id of the tool call")
    type: str = Field(..., description="The type of the tool call")
    function: Tool = Field(..., description="The tool of the tool call")


class LLMResponse(BaseModel):
    """The response of the LLM"""

    content: str = Field(..., description="The content of the response")
    tool_calls: list[ToolCall] = Field(default=[], description="The tool calls of the response")

    
class FlowStepContext(BaseModel):
    """The context of the flow step"""

    step_input: Union[str, List[Message]] = Field(..., description="The input of the flow step")
    handler: Agent = Field(..., description="The handler of the flow step")
    message_context: List[Message] = Field(..., description="The message context of the flow step")


class Memory(BaseModel):

    messages: List[Message] = Field(..., description="The messages of the memory")

    def add_message(self, message: Message):
        self.messages.append(message)