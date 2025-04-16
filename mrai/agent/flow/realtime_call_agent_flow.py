from abc import ABC, abstractmethod
import json
import re
from typing import AsyncIterator, Callable, Literal
from mrai.agent.agent import Agent, RealtimeCallAgent
from mrai.agent.flow.base_flow import BaseFlow
from mrai.agent.schema import FlowInput, Memory, Message
from loguru import logger


class MemoryOrganizer(ABC):
    
    @abstractmethod
    async def organize(self, content_cache: str, observation: dict, memory: dict, flow_input: str):
        pass


class RealtimeCallAgentFlow(BaseFlow):
    
    memory: dict
    
    def __init__(
        self, agents: dict[str, Agent],
        memory_organizer: MemoryOrganizer,
        tool_call: bool = True,
        memory_build_type: Literal["auto", "manual"] = "auto"
    ):
        self.memory = {}
        self.memory_organizer = memory_organizer
        super().__init__(agents)
        # realtime call agent flow can not assign agent to other agents
        for agent in agents.values():
            agent.tools = [tool for tool in agent.tools if tool.name != "assign_agent"]
        if not tool_call:
            # if tool_call is False, remove all tools from all agents, including terminate tool
            for agent in agents.values():
                agent.tools = []
        
        self.memory_build_type = memory_build_type

    async def run(self, flow_input: FlowInput):
        valid_input = flow_input.get_valid_input()
        if valid_input is None:
            raise ValueError("Flow input is not valid")
        
        primary_agent = self.agents.get("primary")
        if primary_agent is not None:
            if not isinstance(primary_agent, RealtimeCallAgent):
                raise ValueError("Primary agent must be a RealtimeCallAgent")
            primary_agent.add_user_message(valid_input)
            # init agent prompt
            await self.rebuild_memory(primary_agent)
            await self.step(primary_agent)
        else:
            raise ValueError("Primary agent is not provided")
        
    async def step(self, agent: RealtimeCallAgent):

        print("-" * 100)

        stream_generator = agent.action()
        if not isinstance(stream_generator, AsyncIterator):
            # if the agent action result is not an AsyncIterator, raise an error
            raise ValueError("Agent action must return a AsyncIterator")
        
        # append each chunk to the content_cache
        other_content_cache = ""
        content_cache = ""
        observation = {}
        async for chunk in stream_generator:
            try:
                formatted_chunk = await self.handle_chunk(chunk)
                if formatted_chunk["type"] == "content":
                    content_cache += formatted_chunk["content"]
                else:
                    other_content_cache += formatted_chunk["content"]
                await self.handle_formatted_chunk(formatted_chunk)
                interruption, tool_call = await self.after_new_chunk(content_cache, other_content_cache)
                if interruption:
                    terminate, tool_call_result = await self.handle_tool_call(tool_call)
                    observation = {
                        "tool_call": tool_call,
                        "tool_call_result": tool_call_result
                    }
                    if terminate:
                        return
                    break
            except Exception as e:
                logger.exception(f"Error handling chunk: {e}")
                observation = {
                    "error": str(e)
                }
                break
            
        terminate = await self.memory_organizer.organize(content_cache, observation, self.memory, agent.user_input)
        if terminate == True:
            return
        await self.rebuild_memory(agent)
        await self.step(agent)
        
    async def rebuild_memory(self, agent: RealtimeCallAgent):
        """
        Based on the developer's reorganized Memory, construct prompts for the large model.
        Special keys:
            - system_prompt: The system prompt will be added to the beginning of the prompt.
        """
        new_system_prompt = ""
        new_memory = Memory()
        copy_memory = self.memory.copy()
        if self.memory_build_type == "auto":
            sections = []
            if agent.prompt:
                sections.append(agent.prompt)
            if copy_memory.get("system_prompt"):
                sections.append(copy_memory.get("system_prompt"))
            if agent.user_input:
                sections.append(f"<user_input>{agent.user_input}</user_input>")
            for key, value in copy_memory.items():
                if key == "system_prompt" or key == "user_input":
                    continue
                if isinstance(value, str):
                    sections.append(f"<{key}>{value}</{key}>")
                elif isinstance(value, dict | list):
                    sections.append(f"<{key}>{json.dumps(value, ensure_ascii=False, indent=2)}</{key}>")
                else:
                    sections.append(f"<{key}>{str(value)}</{key}>")
            
            new_system_prompt = "\n\n".join(sections)
        elif self.memory_build_type == "manual":
            sections = []
            if agent.prompt:
                sections.append(agent.prompt)
            if copy_memory.get("system_prompt"):
                sections.append(copy_memory.get("system_prompt"))
            new_system_prompt = "\n\n".join(sections)
                
        new_memory.add_message(
            Message(
                role="system",
                content=new_system_prompt
            )
        )
        agent.set_memory(new_memory)
        

    async def handle_chunk(self, chunk: str) -> dict:
        if chunk.startswith("content::"):
            return {
                "type": "content",
                "content": chunk[len("content::"):]
            }
        else:
            return {
                "type": chunk.split("::")[0],
                "content": chunk.split("::")[1]
            }
            
    async def handle_formatted_chunk(self, formatted_chunk: dict):
        # TODO: handle the formatted chunk
        if formatted_chunk.get("type") == "content":
            print("\033[92m" + formatted_chunk.get("content","") + "\033[0m", flush=True, end="")
        else:
            print("\033[90m" + formatted_chunk.get("content", "") + "\033[0m", flush=True, end="")
            
            
            
    async def after_new_chunk(
        self,
        content_cache: str,
        other_content_cache: str
    ) -> tuple[bool, dict]:
        
        # use regex to find the <tool_call>...</tool_call>
        tool_call_regex = r"<tool_call>(.*?)</tool_call>"
        tool_call_match = re.search(tool_call_regex, content_cache, re.DOTALL)
        if not tool_call_match:
            return False, {}
        tool_call = tool_call_match.group(1)
        # print(tool_call)
        # 去掉 // 和 /* */ 注释
        # Remove /* */ comments first
        tool_call = re.sub(r'/\*.*?\*/', '', tool_call, flags=re.DOTALL)  # 去掉 /* */
        # Remove // comments only if they appear at the start of a line (ignoring leading whitespace)
        tool_call = re.sub(r'^\s*//.*', '', tool_call, flags=re.MULTILINE)
        # Remove empty lines that might result from comment removal
        tool_call = "\n".join(line for line in tool_call.splitlines() if line.strip())
        
        try:
            tool_call = json.loads(tool_call)
        except json.JSONDecodeError as e:
            logger.error(f"Failed to parse tool_call JSON after comment removal: {e}")
            logger.debug(f"Original content with comments:\n{tool_call_match.group(1)}")
            logger.debug(f"Content after comment removal attempt:\n{tool_call}")
            # Return error instead of raising exception or returning False, {}
            # Let the handle_tool_call function decide what to do with the parse failure
            return True, {"error": f"Failed to parse tool call JSON: {e}", "raw_content": tool_call_match.group(1)}

        return True, tool_call
    
    async def handle_tool_call(self, tool_call: dict) -> tuple[bool, dict]:
        """
        Handle the tool call.
        Return:
            - bool: Is terminate
            - dict: The result of the tool call.
        """
        # Check if the tool_call itself indicates a parsing error from after_new_chunk
        if isinstance(tool_call, dict) and "error" in tool_call and "Failed to parse tool call JSON" in tool_call["error"]:
             return False, {
                "success": False,
                "error": tool_call["error"],
                "raw_content": tool_call.get("raw_content")
            }

        if tool_call.get("name") == "terminate":
            return True, {}
        logger.info(f"Tool call: {tool_call}")
        tools = self.agents["primary"].tools
            
        tool = next((tool for tool in tools if tool.name == tool_call.get("name", "")), None)
        if tool is None:
            return False, {
                "success": False,
                "error": f"Tool {tool_call.get('name', '')} not found"
            }
        arguments = tool_call.get("arguments", {})
        result = tool.execute(**arguments)
        if isinstance(result, dict):
            return False, result
        else:
            return False, {
                "success": True,
                "result": str(result)
            }
