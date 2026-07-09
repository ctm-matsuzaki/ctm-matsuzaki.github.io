from pathlib import Path

from minutes_tool.excel_writer import MinutesExcelWriter
from minutes_tool.parser import MinutesParser


SAMPLE_TEXT = """
基本情報
会議名
ルーツ支援体制・各施設初期対応方針確認会議
日時
不明
場所
オンライン
出席者
名倉さん、長谷川さん、松崎さん

議題① ルーツ側の支援体制・窓口整理
内容
・ルーツ側は3名体制で支援する。
・基本窓口は名倉さんとする。
決定事項
・ルーツ側は3名体制で対応する。

議題② 定例会・報告連絡ルール
内容
・月次定例と週次確認を分ける。
決定事項
・課題管理表を作成・共有する。

ToDo
担当 期限 内容
先方 不明 ルーツ3名をLINE WORKSへ招待する
名倉さん 7月中 大阪・埼玉の現場訪問日程を調整する

次回会議
日時
不明
場所
大阪施設予定
"""


def test_parse_sample_text() -> None:
    minutes = MinutesParser().parse(SAMPLE_TEXT)

    assert minutes.meeting_name == "ルーツ支援体制・各施設初期対応方針確認会議"
    assert len(minutes.agendas) == 2
    assert minutes.agendas[0].title == "ルーツ側の支援体制・窓口整理"
    assert len(minutes.todos) == 2
    assert minutes.todos[1].assignee == "名倉さん"
    assert minutes.next_meeting.place == "大阪施設予定"


def test_write_excel(tmp_path: Path) -> None:
    template = Path("minutes_tool/resources/会議議事録テンプレ.xlsx")
    output = tmp_path / "output.xlsx"
    minutes = MinutesParser().parse(SAMPLE_TEXT)

    MinutesExcelWriter(template).write(minutes, output)

    assert output.exists()
    assert output.stat().st_size > 0

