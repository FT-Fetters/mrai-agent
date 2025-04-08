
from mrai.agent.schema import Tool

class AssignAgent(Tool):
    def __init__(self):
        super().__init__(
            name="assign_agent",
            description="Assign an agent to a task",
            parameters={
                "agent": Tool.ToolParameter(
                    name="agent",
                    description="The name of the agent to assign",
                    type="string",
                    enum=[],
                    required=True
                )
            }
        )

    def execute(self, agent: str):
        """Assign an agent to a task, no need to execute anything, call by flow"""
        pass
