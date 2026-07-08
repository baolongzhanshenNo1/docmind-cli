"""Unit tests for llm_client.py — LLM client with mock HTTP."""
import json
import os
import pytest
import respx
from httpx import Response

from generator.llm_client import LLMClient
from generator.models import AnalysisDocument, Section


# ── Sample LLM JSON responses ──

SAMPLE_LLM_JSON_RESPONSE = {
    'title': '基于深度学习的图像识别系统设计与实现',
    'author': '张三',
    'sections': [
        {'type': 'cover', 'title': '封面', 'content': ''},
        {'type': 'statement', 'title': '郑重声明', 'content': ''},
        {'type': 'abstract_cn', 'title': '摘要', 'content': '本文研究了...'},
        {'type': 'abstract_en', 'title': 'ABSTRACT', 'content': 'This paper...'},
        {'type': 'toc', 'title': '目录', 'content': ''},
        {'type': 'body', 'title': '1 绪论', 'content': '介绍研究背景...'},
        {'type': 'body', 'title': '2 相关技术', 'content': '深度学习基础...'},
        {'type': 'body', 'title': '3 系统设计', 'content': '架构设计...'},
        {'type': 'body', 'title': '4 系统实现', 'content': '关键代码...'},
        {'type': 'body', 'title': '5 实验与分析', 'content': '实验结果...'},
        {'type': 'references', 'title': '参考文献', 'content': ''},
        {'type': 'conclusion', 'title': '结论', 'content': '总结...'},
        {'type': 'acknowledgments', 'title': '致谢', 'content': ''},
    ],
}

SAMPLE_LLM_RESPONSE_WITH_MARKDOWN = """
```json
{
  "title": "Test Paper",
  "author": "Test Author",
  "sections": [
    {"type": "body", "title": "1 Intro", "content": "intro content"}
  ]
}
```
"""

SAMPLE_LLM_RESPONSE_PLAIN = """
{
  "title": "Plain Paper",
  "author": "Plain Author",
  "sections": [
    {"type": "body", "title": "Chapter 1", "content": "content"}
  ]
}
"""

# ── Fixtures ──


@pytest.fixture
def llm_client():
    """Create LLMClient with test API key."""
    return LLMClient(
        api_key='test-key',
        api_base='https://api.test.com/v1',
        model='test-model',
    )


@pytest.fixture
def mock_api(respx_mock):
    """Mock the OpenAI-compatible API endpoint."""
    return respx_mock


# ── LLMClient Initialization ──


class TestLLMClientInit:
    def test_init_with_params(self):
        client = LLMClient(
            api_key='my-key',
            api_base='https://custom.api.com/v1',
            model='gpt-4',
        )
        assert client.api_key == 'my-key'
        assert client.api_base == 'https://custom.api.com/v1'
        assert client.model == 'gpt-4'

    def test_init_from_env(self, monkeypatch):
        monkeypatch.setenv('DOCMIND_API_KEY', 'env-key')
        monkeypatch.setenv('DOCMIND_API_BASE', 'https://env.api.com/v1')
        monkeypatch.setenv('DOCMIND_MODEL', 'env-model')

        client = LLMClient()
        assert client.api_key == 'env-key'
        assert client.api_base == 'https://env.api.com/v1'
        assert client.model == 'env-model'

    def test_init_missing_api_key_raises(self, monkeypatch):
        # Ensure no env var is set
        monkeypatch.delenv('DOCMIND_API_KEY', raising=False)
        with pytest.raises(ValueError, match='API key not configured'):
            LLMClient()

    def test_init_defaults(self, monkeypatch):
        monkeypatch.setenv('DOCMIND_API_KEY', 'test-key')
        monkeypatch.delenv('DOCMIND_API_BASE', raising=False)
        monkeypatch.delenv('DOCMIND_MODEL', raising=False)

        client = LLMClient()
        assert client.api_key == 'test-key'
        assert client.api_base == 'https://api.openai.com/v1'
        assert client.model == 'gpt-4o-mini'


# ── Prompt Building ──


class TestBuildPrompt:
    def test_prompt_includes_requirements(self, llm_client):
        prompt = llm_client._build_prompt(
            '写一篇关于AI的论文', [{'type': 'body'}]
        )
        assert '写一篇关于AI的论文' in prompt
        assert 'body' in prompt

    def test_prompt_includes_section_types(self, llm_client):
        template = [
            {'type': 'cover'},
            {'type': 'body'},
            {'type': 'references'},
        ]
        prompt = llm_client._build_prompt('需求', template)
        assert 'cover' in prompt
        assert 'body' in prompt
        assert 'references' in prompt

    def test_prompt_handles_string_sections(self, llm_client):
        template = ['cover', 'body', 'conclusion']
        prompt = llm_client._build_prompt('需求', template)
        assert 'cover' in prompt
        assert 'body' in prompt
        assert 'conclusion' in prompt


# ── JSON Extraction ──


class TestExtractJson:
    def test_extract_plain_json(self, llm_client):
        result = llm_client._extract_json(
            '{"title": "test", "sections": []}'
        )
        assert result == {'title': 'test', 'sections': []}

    def test_extract_json_from_markdown_fence(self, llm_client):
        result = llm_client._extract_json(
            '```json\n{"title": "test", "sections": []}\n```'
        )
        assert result == {'title': 'test', 'sections': []}

    def test_extract_json_from_code_fence_no_lang(self, llm_client):
        result = llm_client._extract_json(
            '```\n{"title": "test", "sections": []}\n```'
        )
        assert result == {'title': 'test', 'sections': []}

    def test_extract_json_with_surrounding_text(self, llm_client):
        text = 'Here is the result:\n{"title": "test", "sections": []}\nDone.'
        result = llm_client._extract_json(text)
        assert result == {'title': 'test', 'sections': []}

    def test_extract_json_nested_with_markdown(self, llm_client):
        result = llm_client._extract_json(SAMPLE_LLM_RESPONSE_WITH_MARKDOWN)
        assert result['title'] == 'Test Paper'
        assert result['author'] == 'Test Author'
        assert len(result['sections']) == 1

    def test_extract_json_raises_on_invalid(self, llm_client):
        with pytest.raises(ValueError, match='Failed to parse JSON'):
            llm_client._extract_json('not json at all')


# ── generate_analysis (async) ──


class TestGenerateAnalysis:
    @pytest.mark.asyncio
    async def test_successful_generation(self, llm_client, mock_api):
        # Mock the API endpoint
        mock_api.post('https://api.test.com/v1/chat/completions').mock(
            return_value=Response(
                200,
                json={
                    'choices': [
                        {
                            'message': {
                                'content': json.dumps(SAMPLE_LLM_JSON_RESPONSE)
                            }
                        }
                    ]
                },
            )
        )

        template = [
            {'type': 'cover'},
            {'type': 'body'},
            {'type': 'references'},
        ]
        doc = await llm_client.generate_analysis(
            '写一篇关于深度学习的论文', template
        )

        assert isinstance(doc, AnalysisDocument)
        assert doc.title == '基于深度学习的图像识别系统设计与实现'
        assert doc.author == '张三'
        assert len(doc.sections) == 13

        # Verify section types
        types = [s.type for s in doc.sections]
        assert 'cover' in types
        assert 'body' in types
        assert 'references' in types

    @pytest.mark.asyncio
    async def test_handles_markdown_response(self, llm_client, mock_api):
        mock_api.post('https://api.test.com/v1/chat/completions').mock(
            return_value=Response(
                200,
                json={
                    'choices': [
                        {
                            'message': {
                                'content': SAMPLE_LLM_RESPONSE_WITH_MARKDOWN
                            }
                        }
                    ]
                },
            )
        )

        doc = await llm_client.generate_analysis(
            'test', [{'type': 'body'}]
        )
        assert doc.title == 'Test Paper'
        assert doc.author == 'Test Author'

    @pytest.mark.asyncio
    async def test_handles_plain_response(self, llm_client, mock_api):
        mock_api.post('https://api.test.com/v1/chat/completions').mock(
            return_value=Response(
                200,
                json={
                    'choices': [
                        {
                            'message': {
                                'content': SAMPLE_LLM_RESPONSE_PLAIN
                            }
                        }
                    ]
                },
            )
        )

        doc = await llm_client.generate_analysis(
            'test', [{'type': 'body'}]
        )
        assert doc.title == 'Plain Paper'

    @pytest.mark.asyncio
    async def test_api_error_raises(self, llm_client, mock_api):
        mock_api.post('https://api.test.com/v1/chat/completions').mock(
            return_value=Response(401, json={'error': 'Unauthorized'})
        )

        with pytest.raises(Exception):
            await llm_client.generate_analysis(
                'test', [{'type': 'body'}]
            )

    @pytest.mark.asyncio
    async def test_uses_correct_headers(self, llm_client, mock_api):
        mock_api.post('https://api.test.com/v1/chat/completions').mock(
            return_value=Response(
                200,
                json={
                    'choices': [
                        {'message': {'content': '{"title":"t","sections":[]}'}}
                    ]
                },
            )
        )

        await llm_client.generate_analysis('test', [{'type': 'body'}])

        request = mock_api.calls[-1].request
        assert request.headers['Authorization'] == 'Bearer test-key'
        assert request.headers['Content-Type'] == 'application/json'

    @pytest.mark.asyncio
    async def test_uses_correct_model(self, llm_client, mock_api):
        mock_api.post('https://api.test.com/v1/chat/completions').mock(
            return_value=Response(
                200,
                json={
                    'choices': [
                        {'message': {'content': '{"title":"t","sections":[]}'}}
                    ]
                },
            )
        )

        await llm_client.generate_analysis('test', [{'type': 'body'}])

        request_body = json.loads(mock_api.calls[-1].request.content)
        assert request_body['model'] == 'test-model'


# ── generate_analysis_sync ──


class TestGenerateAnalysisSync:
    def test_sync_wrapper(self, llm_client, mock_api):
        mock_api.post('https://api.test.com/v1/chat/completions').mock(
            return_value=Response(
                200,
                json={
                    'choices': [
                        {
                            'message': {
                                'content': json.dumps(SAMPLE_LLM_JSON_RESPONSE)
                            }
                        }
                    ]
                },
            )
        )

        doc = llm_client.generate_analysis_sync(
            'test', [{'type': 'body'}]
        )
        assert isinstance(doc, AnalysisDocument)
        assert doc.title == '基于深度学习的图像识别系统设计与实现'
