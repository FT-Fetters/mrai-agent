TOOL_CALL_RULE = \
"""
## This is the usage of the tool and some rules

### Tool Usage

```
<tool_call>
{{
    "name": "<tool_name>",
    "arguments": {{
        "<argument_name>": "<argument_value>"
    }}
}}
</tool_call>
```


### Rules

1. Must start with <tool_call> and end with </tool_call>.
2. <tool_call>The tag contains a JSON formatted string that describes the tool call.

### Tools

{tools}
"""