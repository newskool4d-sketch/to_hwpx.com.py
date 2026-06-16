from __future__ import annotations

import hwp_writer


HwpParameterValue = str | int | bool


class RejectedSizeFieldError(RuntimeError):
    pass


class FakeProgrammerBugError(ValueError):
    pass


class FakeSet:
    def __init__(self) -> None:
        self.items: dict[str, HwpParameterValue] = {}

    def SetItem(self, key: str, value: HwpParameterValue) -> None:
        self.items[key] = value


class SizeRejectingSet(FakeSet):
    def SetItem(self, key: str, value: HwpParameterValue) -> None:
        if key in {"WidthValue", "HeightValue"}:
            raise RejectedSizeFieldError(f"reject {key}")
        super().SetItem(key, value)


class FakeInsertText:
    def __init__(self) -> None:
        self.HSet = FakeSet()
        self.Text = ""


class FakeParameterSet:
    def __init__(self) -> None:
        self.HInsertText = FakeInsertText()


class FakeAction:
    def __init__(self, hwp: FakeHwp, name: str) -> None:
        self.hwp = hwp
        self.name = name

    def CreateSet(self) -> FakeSet:
        return FakeSet()

    def GetDefault(self, parameter_set: FakeSet) -> None:
        self.hwp.action_defaults.append(self.name)

    def Execute(self, parameter_set: FakeSet) -> None:
        self.hwp.executed_actions.append(self.name)
        self.hwp.executed_parameters.append((self.name, dict(parameter_set.items)))


class FakeHAction:
    def __init__(self, hwp: FakeHwp) -> None:
        self.hwp = hwp

    def GetDefault(self, action_name: str, parameter_set: FakeSet) -> None:
        self.hwp.action_defaults.append(action_name)

    def Execute(self, action_name: str, parameter_set: FakeSet) -> None:
        self.hwp.executed_actions.append(action_name)
        if action_name == "InsertText":
            self.hwp.inserted_texts.append(self.hwp.HParameterSet.HInsertText.Text)

    def Run(self, command: str) -> None:
        self.hwp.run_commands.append(command)


class FakeHwp:
    def __init__(self) -> None:
        self.HParameterSet = FakeParameterSet()
        self.HAction = FakeHAction(self)
        self.created_actions: list[str] = []
        self.action_defaults: list[str] = []
        self.executed_actions: list[str] = []
        self.executed_parameters: list[tuple[str, dict[str, HwpParameterValue]]] = []
        self.run_commands: list[str] = []
        self.inserted_texts: list[str] = []

    def CreateAction(self, name: str) -> FakeAction:
        self.created_actions.append(name)
        return FakeAction(self, name)


class ValueErrorColumnWidthAction(FakeAction):
    def Execute(self, parameter_set: FakeSet) -> None:
        if self.name == "TableColWidth":
            raise FakeProgrammerBugError("programmer bug")
        super().Execute(parameter_set)


class ValueErrorColumnWidthHwp(FakeHwp):
    def CreateAction(self, name: str) -> FakeAction:
        self.created_actions.append(name)
        return ValueErrorColumnWidthAction(self, name)


class RuntimeErrorColumnWidthAction(FakeAction):
    def Execute(self, parameter_set: FakeSet) -> None:
        if self.name == "TableColWidth":
            raise RejectedSizeFieldError("recoverable width failure")
        super().Execute(parameter_set)


class RuntimeErrorColumnWidthHwp(FakeHwp):
    def CreateAction(self, name: str) -> FakeAction:
        self.created_actions.append(name)
        return RuntimeErrorColumnWidthAction(self, name)


class SizeRejectingTableCreateAction(FakeAction):
    def CreateSet(self) -> FakeSet:
        if self.name == "TableCreate":
            return SizeRejectingSet()
        return super().CreateSet()


class SizeRejectingTableCreateHwp(FakeHwp):
    def CreateAction(self, name: str) -> FakeAction:
        self.created_actions.append(name)
        return SizeRejectingTableCreateAction(self, name)


def test_build_doc_renders_current_block_types_without_hwp() -> None:
    hwp = FakeHwp()
    blocks = [
        {"type": "h", "level": 1, "text": "제목"},
        {"type": "p", "text": "본문"},
        {"type": "li", "depth": 1, "marker": "1.", "content": "항목", "text": "1. 항목"},
        {"type": "bq", "text": "인용"},
        {"type": "code", "text": "print('hi')"},
        {"type": "hr"},
        {"type": "table", "header": ["열1", "열2"], "rows": [["A", "B"]]},
        {"type": "official_header", "key": "수신", "value": "홍길동"},
    ]

    hwp_writer.build_doc(hwp, blocks)

    for expected_text in ("제목", "본문", "1. 항목", "인용", "print('hi')", "열1", "열2", "A", "B"):
        assert expected_text in hwp.inserted_texts
    assert any(text.startswith("─" * 10) for text in hwp.inserted_texts)
    assert any(text.startswith("수신") and text.endswith("홍길동") for text in hwp.inserted_texts)
    assert "TableCreate" in hwp.created_actions
    assert "TableColWidth" in hwp.created_actions
    assert "BreakPara" in hwp.run_commands
    assert "MoveDocEnd" in hwp.run_commands


def test_build_doc_ignores_unknown_block_type_without_hwp() -> None:
    hwp = FakeHwp()

    hwp_writer.build_doc(hwp, [{"type": "unknown", "text": "ignored"}])

    assert hwp.created_actions == []
    assert hwp.action_defaults == []
    assert hwp.executed_actions == []
    assert hwp.run_commands == []
    assert hwp.inserted_texts == []


def test_insert_table_does_not_swallow_programmer_errors() -> None:
    hwp = ValueErrorColumnWidthHwp()
    error: ValueError | None = None

    try:
        hwp_writer.insert_table(hwp, ["열1", "열2"], [["A", "B"]])
    except ValueError as exc:
        error = exc

    assert error is not None
    assert "programmer bug" in str(error)


def test_insert_table_sends_column_width_warning_to_stderr(capsys) -> None:
    hwp = RuntimeErrorColumnWidthHwp()

    hwp_writer.insert_table(hwp, ["열1", "열2"], [["A", "B"]])

    captured = capsys.readouterr()
    assert "열 너비 조정 실패" in captured.err
    assert "recoverable width failure" in captured.err
    assert "열 너비 조정 실패" not in captured.out
    assert "열1" in hwp.inserted_texts
    assert "A" in hwp.inserted_texts


def test_insert_table_warns_when_optional_size_fields_are_rejected(capsys) -> None:
    hwp = SizeRejectingTableCreateHwp()

    hwp_writer.insert_table(hwp, ["열1", "열2"], [["A", "B"]])

    captured = capsys.readouterr()
    assert "표 크기 설정 실패" in captured.err
    assert "WidthValue" in captured.err
    assert "HeightValue" in captured.err
    assert "TableCreate" in hwp.executed_actions
    assert "열1" in hwp.inserted_texts


def test_insert_table_applies_mvp_width_and_alignment_metadata() -> None:
    hwp = FakeHwp()

    hwp_writer.insert_table(
        hwp,
        ["시작", "종료", "분", "내용", "담당"],
        [["10:00", "10:05", "5", "인사", "담당자"]],
        table_role="schedule",
        column_widths=[14, 14, 9, 49, 14],
    )

    width_sets = [items for name, items in hwp.executed_parameters if name == "TableColWidth"]
    assert [items["Width"] for items in width_sets] == [1960, 1960, 1260, 6860, 1960]
    assert "CellBorderFill" not in hwp.created_actions

    para_sets = [items for name, items in hwp.executed_parameters if name == "ParagraphShape"]
    assert para_sets[0]["Align"] == 3
    assert para_sets[5]["Align"] == 1


def test_insert_table_accepts_explicit_total_width_for_renderer_page_width() -> None:
    # Given
    hwp = FakeHwp()

    # When
    hwp_writer.insert_table(
        hwp,
        ["용어", "정의"],
        [["위탁", "외부 기관에 맡김"]],
        total_width=24000,
    )

    # Then
    width_sets = [items for name, items in hwp.executed_parameters if name == "TableColWidth"]
    assert [items["Width"] for items in width_sets] == [7200, 16800]
