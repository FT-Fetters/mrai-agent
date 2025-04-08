
from mrai.agent.schema import Tool


class Terminate(Tool):
    def __init__(self):
        super().__init__(
            name="terminate",
            description="Terminate the current agent execution flow, used to end a conversation or task",
            parameters={}
        )
    
    def execute(self, **kwargs):
        """Terminate the agent"""
        pass