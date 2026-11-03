"""LLM 보완: 가짜 SDK 로 호출 흐름·안전 검사·검증 후 채택을 테스트한다."""

import json
from types import SimpleNamespace

import pytest

from sheetproof import run
from sheetproof.llm import (LLMError, build_prompt, complete_with_llm, llm_candidates, make_caller,
                            validate_expr)

GOOD = {"Checks!E32": 'xl.concat(xl.to_text(g("Checks!B9")).title(), "!!")'}


class FakeAnthropic:
    """anthropic.Anthropic 대역."""

    calls = []

    def __init__(self, api_key):
        self.api_key = api_key
        self.messages = self

    def create(self, **kw):
        FakeAnthropic.calls.append(kw)
        return SimpleNamespace(stop_reason="end_turn", content=[
            SimpleNamespace(type="thinking", thinking=""),
            SimpleNamespace(type="text", text=json.dumps({"expr": self.reply, "note": "테스트"})),
        ])


class FakeOpenAI:
    def __init__(self, api_key, base_url=None):
        self.base_url = base_url
        self.chat = SimpleNamespace(completions=self)

    def create(self, model, messages):
        msg = SimpleNamespace(content='```json\n{"expr": "%s"}\n```' % FakeOpenAI.reply.replace('"', '\\"'))
        return SimpleNamespace(choices=[SimpleNamespace(message=msg)])


@pytest.fixture
def errors_sample(samples_dir):
    return run(samples_dir / "errors_text.xlsx")


def test_candidates_exclude_volatile(errors_sample):
    a, _, _ = errors_sample
    assert llm_candidates(a) == ["Checks!E32"]


def test_prompt_has_no_cell_values(errors_sample):
    a, _, _ = errors_sample
    prompt = build_prompt(a, "Checks!E32")
    assert "PROPER" in prompt and "Checks!B9" in prompt
    assert "sheetproof" not in prompt  # B9 의 실제 값은 전송하지 않는다


@pytest.mark.parametrize("expr,ok", [
    ('xl.SUM(xl.rng(v, "S!A1:A3"))', True),
    ("[x * 2 for x in range(3)]", True),
    ("__import__('os').system('x')", False),
    ("xl.re.compile('x')", False),
    ("xl._safe", False),
    ("open('f')", False),
    ("(lambda: 1)()", False),
    ("g('A1').__class__", False),
])
def test_validate_expr(expr, ok):
    assert (validate_expr(expr) is None) == ok


def test_anthropic_fake_flow_accepts_verified(errors_sample):
    a, _, _ = errors_sample
    FakeAnthropic.reply = GOOD["Checks!E32"]
    FakeAnthropic.calls.clear()
    call = make_caller("anthropic", "sk-test", client_factory=FakeAnthropic)
    res = complete_with_llm(a, call)
    assert FakeAnthropic.calls[0]["model"] == "claude-opus-5-5"
    assert res.candidates[0].accepted, res.candidates[0].reason
    assert res.overrides == GOOD


def test_wrong_llm_answer_is_rejected(errors_sample):
    a, _, _ = errors_sample
    FakeOpenAI.reply = 'xl.UPPER(g("Checks!B9"))'
    call = make_caller("openai", "sk-test", client_factory=FakeOpenAI)
    res = complete_with_llm(a, call)
    assert not res.candidates[0].accepted
    assert "검증 실패" in res.candidates[0].reason
    assert res.overrides == {}


def test_unsafe_llm_answer_is_rejected(errors_sample):
    a, _, _ = errors_sample
    FakeOpenAI.reply = "__import__('os').getcwd()"
    call = make_caller("openai", "sk-test", client_factory=FakeOpenAI)
    res = complete_with_llm(a, call)
    assert "안전 검사 실패" in res.candidates[0].reason


def test_openai_compatible_requires_base_url():
    with pytest.raises(LLMError):
        make_caller("openai_compatible", "k", "m", None, client_factory=FakeOpenAI)
    make_caller("openai_compatible", "k", "m", "http://localhost:1/v1", client_factory=FakeOpenAI)
