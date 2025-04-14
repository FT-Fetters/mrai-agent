
from mrai.agent.schema import Tool


class Terminate(Tool):
    def __init__(self):
        super().__init__(
            name="terminate",
            description="When the tasks of the Agent or workflow have been completed, call terminate to end this process",
            parameters={}
        )
    
    def execute(self, **kwargs):
        """Terminate the agent"""
        pass