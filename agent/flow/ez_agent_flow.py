
from agent.agent import Agent
from agent.flow.base_flow import BaseFlow
from agent.schema import FlowInput, FlowStepContext, Message


class EzAgentFlow(BaseFlow):

    def __init__(self, agents: dict[str, Agent]):
        super().__init__(agents)


    def run(self, flow_input: FlowInput):
        valid_input = flow_input.get_valid_input()
        if valid_input is None:
            raise ValueError("Flow input is not valid")
        # 
    
    def flow_step(self, flow_step_context: FlowStepContext):
        step_input = flow_step_context.step_input
        if isinstance(step_input, str):
            flow_step_context.message_context.append(
                Message(
                    role="user",
                    content=step_input,
                )
            )
        elif isinstance(step_input, list):
            flow_step_context.message_context.extend(step_input)
        else:
            raise ValueError("Invalid step input")
        handler = flow_step_context.handler
        