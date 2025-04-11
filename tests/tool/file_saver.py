

from mrai.agent.schema import Tool


class FileSaver(Tool):
    def __init__(self):
        super().__init__(
            name="file_saver",
            description="Save a file to the local filesystem",
            parameters={
                "file_path": Tool.ToolParameter(
                    name="file_path",
                    type="string",
                    description="The path to save the file to",
                    required=True
                ),
                "content": Tool.ToolParameter(
                    name="content",
                    type="string",
                    description="The content to save to the file",
                    required=True
                )
            }
        )

    def execute(self, file_path: str, content: str):
        with open(file_path, "w") as f:
            f.write(content)
        return {
            "success": True,
            "message": f"File {file_path} saved successfully"
        }


