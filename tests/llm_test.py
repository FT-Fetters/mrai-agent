import asyncio
from agent.llm.llm import LLM
from agent.llm.llm_config import LLMConfig


llm_config = LLMConfig(
    base_url="https://yunwu.ai/v1",
    api_key="sk-PlosMpk7twm2PMuYQLu8r3pIuSmKDHVbnIcGPZBBvflKLIHN",
    model="gpt-4o"
)


llm = LLM(llm_config)


async def test_llm():
    response = await llm.chat([
        "i wanne build a llm agent system, how can i do?"
    ])
    print(response)


if __name__ == "__main__":
    asyncio.run(test_llm())
