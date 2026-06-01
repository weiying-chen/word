from daai_program_extract import (
    build_episodes,
    extract_episode_rows_from_html,
    extract_youtube_fields,
    enrich_with_youtube,
)


def test_extract_episode_rows_falls_back_when_episode_json_is_broken() -> None:
    html = """
    <div class="item" id="episode-0">
      <div class="item-description">
        <div class="title">判讀骨鬆</div>
        <div class="date">2026-02-05</div>
      </div>
    </div>
    <script>
      document.getElementById('episode-0').addEventListener('click', function(){
        var episodeJson = '{\\"EpID\\":\\"6723\\",\\"EpPremiere\\":\\"2026-02-05 20:49:50\\",\\"EpTitle\\":\\"判讀骨鬆\\",\\"YTID\\":\\"XOzUmSNOdO8\\"}'
      });
    </script>

    <div class="item" id="episode-1">
      <div class="item-description">
        <div class="title">肝癌不開刀</div>
        <div class="date">2026-05-06</div>
      </div>
    </div>
    <script>
      document.getElementById('episode-1').addEventListener('click', function(){
        var episodeJson = '{\\"EpID\\":\\"6787\\",\\"EpPremiere\\":\\"2026-05-06 20:50:10\\",\\"EpTitle\\":\\"肝癌不開刀\\",\\"Description\\":\\"BROKEN
      });
    </script>
    """

    rows = extract_episode_rows_from_html(html)
    assert len(rows) == 2

    assert rows[0]["episode_index"] == 0
    assert rows[0]["epid"] == "6723"
    assert rows[0]["date"] == "2026-02-05"
    assert rows[0]["title"] == "判讀骨鬆"
    assert rows[0]["ytid"] == "XOzUmSNOdO8"

    assert rows[1]["episode_index"] == 1
    assert rows[1]["epid"] == "6787"
    assert rows[1]["date"] == "2026-05-06"
    assert rows[1]["title"] == "肝癌不開刀"
    assert rows[1]["ytid"] == ""


def test_extract_youtube_fields_parses_title_description_and_last_timestamp() -> None:
    html = """
    <meta property="og:title" content="【大愛醫生館】 腰椎連環「扁」 20260521">
    <script>
    var ytInitialPlayerResponse = {"videoDetails":{"shortDescription":"line1\\n00:11｜片頭\\n07:18｜腰椎連環「扁」"}}
    </script>
    """
    fields = extract_youtube_fields(html)
    assert fields["youtubeTitle"] == "【大愛醫生館】 腰椎連環「扁」 20260521"
    assert "00:11｜片頭" in fields["youtubeDescription"]
    assert fields["descriptionLastTimestampLine"] == "07:18｜腰椎連環「扁」"


def test_build_episodes_uses_required_output_keys() -> None:
    rows = [
        {
            "episode_index": 8,
            "epid": "6797",
            "date": "2026-05-20",
            "title": "肺腺癌先禮後兵",
            "ytid": "P0uiRM2no18",
        }
    ]
    episodes = build_episodes(rows)
    assert episodes[0]["episodeIndex"] == 8
    assert episodes[0]["epId"] == "6797"
    assert episodes[0]["titleZh"] == "肺腺癌先禮後兵"
    assert episodes[0]["ytId"] == "P0uiRM2no18"
    assert episodes[0]["youtubeUrl"] == "https://www.youtube.com/watch?v=P0uiRM2no18"


def test_extract_youtube_fields_handles_missing_og_title() -> None:
    html = """
    <script>
    var ytInitialPlayerResponse = {"videoDetails":{"shortDescription":"00:31｜片頭\\n02:18｜主題"}}
    </script>
    """
    fields = extract_youtube_fields(html)
    assert fields["youtubeTitle"] == ""
    assert "00:31｜片頭" in fields["youtubeDescription"]
    assert fields["descriptionLastTimestampLine"] == "02:18｜主題"


def test_extract_youtube_fields_handles_missing_short_description() -> None:
    html = '<meta property="og:title" content="【大愛醫生館】 測試標題 20260601">'
    fields = extract_youtube_fields(html)
    assert fields["youtubeTitle"] == "【大愛醫生館】 測試標題 20260601"
    assert fields["youtubeDescription"] == ""
    assert fields["descriptionLastTimestampLine"] == ""


def test_extract_episode_rows_handles_multiple_broken_episode_json_blocks() -> None:
    html = """
    <div class="item" id="episode-0">
      <div class="item-description"><div class="title">A</div><div class="date">2026-05-01</div></div>
    <script>
      document.getElementById('episode-0').addEventListener('click', function(){
        var episodeJson = '{\\"EpID\\":\\"7001\\",\\"EpPremiere\\":\\"2026-05-01 20:00:00\\",\\"EpTitle\\":\\"A\\",\\"Description\\":\\"BROKEN
      });
    </script>

    <div class="item" id="episode-1">
      <div class="item-description"><div class="title">B</div><div class="date">2026-05-02</div></div>
    <script>
      document.getElementById('episode-1').addEventListener('click', function(){
        var episodeJson = '{\\"EpID\\":\\"7002\\",\\"EpPremiere\\":\\"2026-05-02 20:00:00\\",\\"EpTitle\\":\\"B\\",\\"Description\\":\\"BROKEN
      });
    </script>
    """
    rows = extract_episode_rows_from_html(html)
    assert len(rows) == 2
    assert rows[0]["epid"] == "7001"
    assert rows[1]["epid"] == "7002"
    assert rows[0]["title"] == "A"
    assert rows[1]["title"] == "B"


def test_enrich_with_youtube_skips_episode_without_ytid(monkeypatch) -> None:
    episodes = [
        {
            "episodeIndex": 0,
            "epId": "7001",
            "date": "2026-05-01",
            "titleZh": "A",
            "ytId": "",
            "youtubeUrl": "",
            "youtubeTitle": "",
            "youtubeDescription": "",
            "descriptionLastTimestampLine": "",
        }
    ]

    called = {"value": False}

    def _boom(_: str) -> str:
        called["value"] = True
        raise AssertionError("fetch_html should not be called")

    monkeypatch.setattr("daai_program_extract.fetch_html", _boom)
    enrich_with_youtube(episodes)
    assert called["value"] is False
