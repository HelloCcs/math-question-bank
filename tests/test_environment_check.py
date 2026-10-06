from pathlib import Path
from check_environment import run_check

def test_environment_check_creates_healthy_sqlite(tmp_path, monkeypatch):
    (tmp_path/".env").write_text("DB_DRIVER=sqlite\nSQLITE_PATH=data/test.db\n",encoding="utf-8")
    code,lines=run_check(tmp_path)
    assert code in {0,2}
    assert any("SQLite 可读写且完整" in line for line in lines)
    assert (tmp_path/"data/test.db").exists()
