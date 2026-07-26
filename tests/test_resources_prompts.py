import json
import pytest
from server import mcp


async def test_registered_resources():
    """Verify that site-info template and configured-sites resource are registered."""
    resources = await mcp.list_resources()
    resource_uris = [str(r.uri) for r in resources]
    assert "sharepoint://configured-sites" in resource_uris

    templates = await mcp.list_resource_templates()
    template_uris = [str(t.uriTemplate) for t in templates]
    assert "sharepoint://site-info" in template_uris


async def test_registered_prompts():
    """Verify that search_and_analyze and site_audit prompts are registered."""
    prompts = await mcp.list_prompts()
    prompt_names = [p.name for p in prompts]
    
    assert "search_and_analyze" in prompt_names
    assert "site_audit" in prompt_names


async def test_configured_sites_resource_output():
    """Verify that configured-sites resource returns the expected JSON data."""
    contents = await mcp.read_resource("sharepoint://configured-sites")
    assert len(contents) > 0
    
    data = json.loads(contents[0].content)
    assert isinstance(data, list)
    assert len(data) > 0
    assert "name" in data[0]
    assert "url" in data[0]


async def test_prompts_templates():
    """Verify that prompt handlers return templates containing the expected text."""
    prompt_result = await mcp.get_prompt("site_audit", {"site_name": "HR"})
    assert len(prompt_result.messages) == 1
    
    message_content = prompt_result.messages[0].content
    if hasattr(message_content, "text"):
        text = message_content.text
    else:
        text = str(message_content)
        
    assert "HR" in text
    assert "list_available_sites" in text
