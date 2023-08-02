from collections.abc import Iterator
from pathlib import Path

import comtypes.client
import pytest

import adotypes
from adotypes import com_dlls


class Test_Create:
    def test_create_db_file(self, tmp_path: Path):
        db_fullpath = tmp_path / "sample.mdb"
        db_fspath = db_fullpath.__fspath__()
        CATALOG_CONN_STR = (
            "Provider=Microsoft.ACE.OLEDB.12.0;"
            f"Data Source={db_fspath};"
            "Jet OLEDB:Engine Type=5"
        )
        assert not db_fullpath.exists()
        adotypes.connect(create=CATALOG_CONN_STR).close()
        assert db_fullpath.exists()


class Test_ConnectionFailure:
    def test_takes_create_and_open(self):
        with pytest.raises(TypeError):
            adotypes.connect(create="create", open="open")  # type: ignore

    def test_not_takes_create_or_open(self):
        with pytest.raises(TypeError):
            adotypes.connect()  # type: ignore

    def test_takes_invalid_open_conn_str(self):
        with pytest.raises(adotypes.ProgrammingError):
            adotypes.connect(open="hoge")

    def test_takes_invalid_create_conn_str(self):
        with pytest.raises(adotypes.ProgrammingError):
            adotypes.connect(create="hoge")


@pytest.fixture
def db_fspath(tmp_path_factory) -> Iterator[str]:
    fullpath = tmp_path_factory.mktemp("sample") / "sample.mdb"
    fspath = fullpath.__fspath__()
    conn_str = (
        "Provider=Microsoft.ACE.OLEDB.12.0;"
        f"Data Source={fspath};"
        "Jet OLEDB:Engine Type=5"
    )
    catalog = comtypes.client.CreateObject(
        com_dlls.adox.Catalog, interface=com_dlls.adox._Catalog
    )
    catalog.Create(conn_str)
    catalog.ActiveConnection.Close()
    yield fspath


class Test_Open:
    def _connect(self, fspath: str) -> adotypes.Connection:
        return adotypes.connect(
            open=f"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={fspath}"
        )

    def test_commit_creating_table(self, db_fspath):
        conn = self._connect(db_fspath)
        orig = len(conn.adox_catalog.Tables)
        conn.cursor().execute("CREATE TABLE MyTable (Id INT, Name TEXT)")
        assert orig + 1 == len(conn.adox_catalog.Tables)
        conn.commit()
        assert orig + 1 == len(conn.adox_catalog.Tables)
        conn.close()
        with self._connect(db_fspath) as conn:
            assert orig + 1 == len(conn.adox_catalog.Tables)

    def test_rollback_creating_table(self, db_fspath):
        conn = self._connect(db_fspath)
        orig = len(conn.adox_catalog.Tables)
        conn.cursor().execute("CREATE TABLE MyTable (Id INT, Name TEXT)")
        assert orig + 1 == len(conn.adox_catalog.Tables)
        conn.rollback()
        assert orig == len(conn.adox_catalog.Tables)
        conn.close()
        with self._connect(db_fspath) as conn:
            assert orig == len(conn.adox_catalog.Tables)

    def test_commit_when_exiting(self, db_fspath):
        with self._connect(db_fspath) as conn:
            orig = len(conn.adox_catalog.Tables)
            conn.cursor().execute("CREATE TABLE MyTable (Id INT, Name TEXT)")
            assert orig + 1 == len(conn.adox_catalog.Tables)
        with self._connect(db_fspath) as conn:
            assert orig + 1 == len(conn.adox_catalog.Tables)

    def test_rollback_when_exiting(self, db_fspath):
        with self._connect(db_fspath) as conn:
            orig = len(conn.adox_catalog.Tables)
        with pytest.raises(ZeroDivisionError):
            with self._connect(db_fspath) as conn:
                assert orig == len(conn.adox_catalog.Tables)
                conn.cursor().execute("CREATE TABLE MyTable (Id INT, Name TEXT)")
                assert orig + 1 == len(conn.adox_catalog.Tables)
                _ = 1 / 0
        with self._connect(db_fspath) as conn:
            assert orig == len(conn.adox_catalog.Tables)

    def test_rollback_when_execute_is_failed(self, db_fspath):
        with self._connect(db_fspath) as conn:
            orig = len(conn.adox_catalog.Tables)
        with pytest.raises(adotypes.DatabaseError):
            with self._connect(db_fspath) as conn:
                assert orig == len(conn.adox_catalog.Tables)
                conn.cursor().execute("hoge")
        with self._connect(db_fspath) as conn:
            assert orig == len(conn.adox_catalog.Tables)


class Test_CRUD:
    @pytest.fixture
    def conn(self, db_fspath: Path) -> Iterator[adotypes.Connection]:
        with adotypes.connect(
            open=f"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={db_fspath}"
        ) as conn:
            new_tbl = comtypes.client.CreateObject(
                com_dlls.adox.Table, interface=com_dlls.adox._Table
            )
            cols: com_dlls.adox.Columns = new_tbl.Columns
            new_tbl.Name = "MyTable"
            cols.Append("Id", com_dlls.adodb.adInteger)
            cols.Append("Name", com_dlls.adodb.adLongVarWChar)
            tables: com_dlls.adox.Tables = conn.adox_catalog.Tables
            tables.Append(new_tbl)
            yield conn

    def test_crud(self, conn: adotypes.Connection):
        with conn.cursor() as c:
            c.execute("SELECT Id, Name FROM MyTable")
            assert c.fetchone() is None
            assert c.fetchmany(2) == []
            assert c.fetchall() == []
        conn.cursor().execute("INSERT INTO MyTable (Id, Name) VALUES (1, 'John')")
        with conn.cursor() as c:
            c.execute("SELECT Id, Name FROM MyTable")
            assert c.fetchone() == (1, "John")
            assert c.fetchone() is None
        conn.cursor().execute("INSERT INTO MyTable (Id, Name) VALUES (2, 'Ringo')")
        with conn.cursor() as c:
            c.execute("SELECT Id, Name FROM MyTable")
            assert c.fetchone() == (1, "John")
            assert c.fetchone() == (2, "Ringo")
            assert c.fetchone() is None
        conn.cursor().execute("INSERT INTO MyTable (Id, Name) VALUES (3, 'Paul')")
        with conn.cursor() as c:
            c.execute("SELECT Id, Name FROM MyTable")
            assert c.fetchmany() == [(1, "John")]
            assert c.fetchmany(2) == [(2, "Ringo"), (3, "Paul")]
            assert c.fetchmany(3) == []
            assert c.fetchone() is None
        conn.cursor().execute("INSERT INTO MyTable (Id, Name) VALUES (4, 'George')")
        with conn.cursor() as c:
            c.execute("SELECT Id, Name FROM MyTable")
            assert c.fetchone() == (1, "John")
            c.arraysize = 2
            assert c.fetchmany() == [(2, "Ringo"), (3, "Paul")]
            assert c.fetchall() == [(4, "George")]
            assert c.fetchall() == []
            assert c.fetchone() is None
        conn.cursor().execute("UPDATE MyTable SET Name = 'Sean' WHERE Id = 1")
        with conn.cursor() as c:
            c.execute("SELECT Id, Name FROM MyTable")
            assert c.fetchall() == [
                (1, "Sean"),
                (2, "Ringo"),
                (3, "Paul"),
                (4, "George"),
            ]
        conn.cursor().execute("DELETE FROM MyTable WHERE Id = 1")
        with conn.cursor() as c:
            c.execute("SELECT Id, Name FROM MyTable WHERE Id = 1")
            assert c.fetchone() is None
        with conn.cursor() as c:
            c.execute("SELECT Id, Name FROM MyTable")
            assert c.fetchall() == [(2, "Ringo"), (3, "Paul"), (4, "George")]
