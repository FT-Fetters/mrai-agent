
import json
from mrai.agent.agent import Agent
from mrai.agent.flow.base_flow import BaseFlow
from mrai.agent.schema import FlowInput, FlowStepContext, Message
from loguru import logger


class AgentFlow(BaseFlow):

    def __init__(self, agents: dict[str, Agent]):
        super().__init__(agents)


    async def run(self, flow_input: FlowInput):
        valid_input = flow_input.get_valid_input()
        if valid_input is None:
            raise ValueError("Flow input is not valid")
        #
        primary_agent = self.agents.get("primary")
        if primary_agent is not None:
            primary_agent.add_user_message(valid_input)
            await self.step(primary_agent)
        else:
            raise ValueError("Primary agent is not provided")
        
    
    async def step(self, agent: Agent):
        _, tool_calls = await agent.action()

        # check if any tool call is a terminate tool
        terminated = any(tool_call.function.name == "terminate" for tool_call in tool_calls)
        # remove the terminate tool call from the tool calls
        tool_calls = [tool_call for tool_call in tool_calls if tool_call.function.name != "terminate"]

        assign_agent_tool_calls = [tool_call for tool_call in tool_calls if tool_call.function.name == "assign_agent"]
        for assign_agent_tool_call in assign_agent_tool_calls:
            assign_agent_name = assign_agent_tool_call.function.arguments.get("agent")
            if assign_agent_name is None:
                raise ValueError("Assign agent name is not provided")
            # assign the agent to the task
        # remove the assign agent tool call from the tool calls
        tool_calls = [tool_call for tool_call in tool_calls if tool_call.function.name != "assign_agent"]

        # execute the tool calls
        for tool_call in tool_calls:
            tool_call_result = tool_call.tool.execute(**tool_call.function.arguments)
            if tool_call_result:
                # add the tool call result to the agent's memory
                agent.memory.add_message(Message(role="tool", content=json.dumps({
                    "name": tool_call.function.name,
                    "result": json.dumps(tool_call_result, ensure_ascii=False),
                    "arguments": tool_call.function.arguments
                })))
                logger.info(f"📄 「{agent.name}」 called tool 「{tool_call.function.name}」 result > \n {json.dumps(tool_call_result, ensure_ascii=False, indent=2)}")
            else:
                logger.info(f"📄 「{agent.name}」 called tool 「{tool_call.function.name}」 but have no result")
                
        if terminated:
            return
        
        # next step
        await self.step(agent)
        
        
