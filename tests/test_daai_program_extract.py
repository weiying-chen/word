from daai_program_extract import extract_episode_rows_from_html


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
