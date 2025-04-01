from abc import ABC, abstractmethod
from pydantic import Field
from typing import List, Literal, Union, TYPE_CHECKING, Any, Dict
from openai import BaseModel

# 移除循环导入
# from agent.agent import Agent

# 如果在类型检查时，导入 Agent 类型
if TYPE_CHECKING:
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

    # context: Union[list[Message], str] = Field(..., description="The context of the flow")

    def get_valid_input(self) -> Union[str, None]:
        """Get the valid input of the flow"""

        if self.text:
            return self.text
        else:
            return None


class Tool(BaseModel, ABC):

    class ToolParameter(BaseModel):
        name: str = Field(..., description="The name of the parameter")
        description: str = Field(..., description="The description of the parameter")
        type: Literal["string", "number", "boolean", "array", "object", "enum"] = Field(..., description="The type of the parameter")
        enum: list[str] = Field(default=[], description="The enum of the parameter")
        required: bool = Field(..., description="Whether the parameter is required")

    name: str = Field(..., description="The name of the tool")
    description: str = Field(..., description="The description of the tool")
    parameters: dict[str, ToolParameter] = Field(default={}, description="The parameters of the tool")
    
    @abstractmethod
    def execute(self, **kwargs):
        """Execute the tool"""
        pass
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert the tool to a dictionary that follows the OpenAI tool format standard"""
        # 创建参数模式对象
        properties: Dict[str, Any] = {}
        required_params: List[str] = []
        
        for param_key, param in self.parameters.items():
            param_schema: Dict[str, Any] = {
                "type": param.type,
                "description": param.description,
            }
            
            # 只有当类型为enum时才添加enum字段
            if param.type == "enum":
                param_schema["enum"] = param.enum
                
            properties[param_key] = param_schema
            
            # 收集必需参数
            if param.required:
                required_params.append(param_key)
        
        # 构建符合OpenAI工具格式的字典
        parameters_dict: Dict[str, Any] = {
            "type": "object",
            "properties": properties
        }
        
        # 只有当有必需参数时才添加required字段
        if required_params:
            parameters_dict["required"] = required_params
            
        return {
            "type": "function",
            "function": {
                "name": self.name,
                "description": self.description,
                "parameters": parameters_dict
            }
        }



class ToolCall(BaseModel):
    """Represents a tool/function call in a message"""

    id: str = Field(..., description="The id of the tool call")
    type: str = Field(..., description="The type of the tool call")
    function: Tool = Field(..., description="The tool of the tool call")
    arguments: dict = Field(default={}, description="The arguments of the tool call")


class LLMResponse(BaseModel):
    """The response of the LLM"""

    content: str = Field(..., description="The content of the response")
    tool_calls: list[ToolCall] = Field(default=[], description="The tool calls of the response")

    
class FlowStepContext(BaseModel):
    """The context of the flow step"""

    step_input: Union[str, List[Message]] = Field(..., description="The input of the flow step")
    handler: 'Agent' = Field(..., description="The handler of the flow step")
    message_context: List[Message] = Field(..., description="The message context of the flow step")


class Memory(BaseModel):

    messages: List[Message] = Field(default=[], description="The messages of the memory")

    def add_message(self, message: Message):
        self.messages.append(message)
        

class Callback(ABC):
    
    @abstractmethod
    def on_action(self):
        pass