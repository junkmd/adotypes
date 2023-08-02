from collections.abc import Iterable, MutableMapping, Sequence
import enum
import logging
import types
from typing import Any, Optional, SupportsIndex, Union
import weakref

import comtypes.client

from adotypes import com_dlls, api_exceptions as exc
from adotypes._hints import _Parameters, _ColumnDescription, _InputSizes, _OutputSize

LOGGER = logging.getLogger(__name__)


class Connection:
    def __init__(self, connector: com_dlls.adodb._Connection, **kwargs: Any) -> None:
        self._trns_lv = 0
        self._connector = connector
        # THIS MUST BE `weakref.WeakValueDictionary`!
        # If this were a built-in list or dictionary, COM objects would cause
        # a serious and tragic memory leak!
        self._cursors: MutableMapping[int, "Cursor"] = weakref.WeakValueDictionary()
        self._trns_lv: int = self._connector.BeginTrans()

    def close(self) -> None:
        if not self._trns_lv:
            LOGGER.debug(f"{self!r} has been closed")
            return
        self._rollback()
        cursors = iter(self._cursors.values())
        for c in cursors:
            c.close()
        del cursors
        self._connector.Close()
        del self._connector
        self._trns_lv = 0
        LOGGER.debug(f"complete closing {self!r}")

    def commit(self) -> None:
        try:
            self._connector.CommitTrans()
        except Exception as e:
            LOGGER.error(str(e), stack_info=True)
            raise exc.ProgrammingError from e
        LOGGER.debug("commit is done")
        self._trns_lv = self._connector.BeginTrans()

    def rollback(self) -> None:
        if not self._trns_lv:
            LOGGER.debug("transaction has not started")
            return
        self._rollback()
        self._trns_lv = self._connector.BeginTrans()

    def cursor(self) -> "Cursor":
        return Cursor(self)

    def _rollback(self) -> None:
        try:
            self._connector.RollbackTrans()
        except Exception as e:
            LOGGER.error(str(e), stack_info=True)
            raise exc.ProgrammingError from e
        LOGGER.debug("rollback is done")

    def __del__(self) -> None:
        self.close()

    def __enter__(self) -> "Connection":
        return self

    def __exit__(
        self,
        exc_type: Optional[type[Exception]],
        exc_val: Optional[Exception],
        exc_tb: Optional[types.TracebackType],
    ) -> None:
        if exc_type:
            self.rollback()
        else:
            self.commit()

    def register_cursor(self, cursor: "Cursor") -> None:
        self._cursors[hash(cursor)] = cursor
        LOGGER.debug(f"{cursor!r} is registered")

    def unregister_cursor(self, cursor: "Cursor") -> None:
        repr_ = repr(cursor)
        del self._cursors[hash(cursor)]
        LOGGER.debug(f"{repr_} is unregistered")

    @property
    def ado_connection(self) -> com_dlls.adodb._Connection:
        return self._connector

    @property
    def adox_catalog(self) -> com_dlls.adox._Catalog:
        catalog = comtypes.client.CreateObject(
            com_dlls.adox.Catalog, interface=com_dlls.adox._Catalog
        )
        catalog.ActiveConnection = self._connector
        return catalog

    def __repr__(self) -> str:
        return f"<Connection object at {id(self):#016x}>"


class Cursor:
    arraysize: int

    def __init__(self, connection: Connection) -> None:
        self.arraysize = 1
        self._connection = connection
        self._connection.register_cursor(self)

    @property
    def description(self) -> list[_ColumnDescription]:
        raise NotImplementedError

    @property
    def rowcount(self) -> int:
        raise NotImplementedError

    def close(self) -> None:
        if not hasattr(self, "_connection"):
            LOGGER.debug(f"{self!r} has been closed")
            return
        if hasattr(self, "_rs"):
            if self._rs.State != com_dlls.adodb.adStateClosed:
                self._rs.Close()
            del self._rs
        self._connection.unregister_cursor(self)
        del self._connection
        LOGGER.debug(f"complete closing {self!r}")

    def execute(
        self, operation: str, parameters: Optional[_Parameters[Any]] = None
    ) -> None:
        cmd = self._create_command(operation)
        try:
            ptr_records_affected, _rs = cmd.Execute()
        except Exception as e:
            msg = (
                "failure;\n"
                f"msg: {e};\n"
                f"cmd: {operation!r};\n"
                f"params: {parameters!r}"
            )
            LOGGER.error(msg, stack_info=True)
            raise exc.DatabaseError(msg) from e
        rs: com_dlls.adodb._Recordset = _rs.QueryInterface(com_dlls.adodb._Recordset)
        ra: int = ptr_records_affected[0].value
        LOGGER.debug(f"success; cmd: {operation!r}; params: {parameters!r}")
        LOGGER.debug(f"records affected is {ra}")
        self._rs = rs

    def executemany(
        self, operation: str, seq_of_parameters: Iterable[_Parameters[Any]]
    ) -> None:
        raise NotImplementedError

    def fetchone(self) -> Optional[tuple[Any, ...]]:
        if self._rs.EOF:
            return None
        result = tuple(f.Value for f in self._rs.Fields)
        self._rs.MoveNext()
        return result

    def fetchmany(self, size: Optional[int] = None) -> list[tuple[Any, ...]]:
        size = self.arraysize if size is None else size
        result = []
        cnt = 0
        while cnt < size and (r := self.fetchone()) is not None:
            result.append(r)
            cnt += 1
        return result

    def fetchall(self) -> list[tuple[Any, ...]]:
        result = []
        while (r := self.fetchone()) is not None:
            result.append(r)
        return result

    def setinputsizes(self, sizes: _InputSizes) -> None:
        raise NotImplementedError  # maybe does nothing

    def setoutputsize(
        self, size: _OutputSize, column: Union[None, int, str] = None
    ) -> None:
        raise NotImplementedError  # maybe does nothing

    def __del__(self) -> None:
        self.close()

    def __enter__(self) -> "Cursor":
        return self

    def __exit__(
        self,
        exc_type: Optional[type[Exception]],
        exc_val: Optional[Exception],
        exc_tb: Optional[types.TracebackType],
    ) -> None:
        self.close()

    @property
    def connection(self) -> Connection:
        return self._connection

    def _create_command(self, text: str) -> com_dlls.adodb._Command:
        cmd = comtypes.client.CreateObject(
            com_dlls.adodb.Command, interface=com_dlls.adodb._Command
        )
        cmd.ActiveConnection = self.connection.ado_connection
        cmd.CommandTimeout = self.connection.ado_connection.CommandTimeout
        cmd.CommandType = com_dlls.adodb.adCmdText
        cmd.CommandText = text
        cmd.Prepared = False
        return cmd

    def __repr__(self) -> str:
        return f"<Cursor object at {id(self):#016x}>"


# Type Objects and Constructors
# https://peps.python.org/pep-0249/#type-objects-and-constructors
class TypeConstants(tuple[int], enum.Enum):
    INTEGER = (
        com_dlls.adodb.adInteger,
        com_dlls.adodb.adSmallInt,
        com_dlls.adodb.adTinyInt,
        com_dlls.adodb.adUnsignedInt,
        com_dlls.adodb.adUnsignedSmallInt,
        com_dlls.adodb.adUnsignedTinyInt,
        com_dlls.adodb.adBoolean,
        com_dlls.adodb.adError,
    )
    ROWID = (com_dlls.adodb.adChapter,)
    LONG = (
        com_dlls.adodb.adBigInt,
        com_dlls.adodb.adFileTime,
        com_dlls.adodb.adUnsignedBigInt,
    )
    EXACT_NUMERIC = (
        com_dlls.adodb.adDecimal,
        com_dlls.adodb.adNumeric,
        com_dlls.adodb.adVarNumeric,
        com_dlls.adodb.adCurrency,
    )
    APPROX_NUMERIC = (
        com_dlls.adodb.adDouble,
        com_dlls.adodb.adSingle,
    )
    STRING = (
        com_dlls.adodb.adBSTR,
        com_dlls.adodb.adChar,
        com_dlls.adodb.adLongVarChar,
        com_dlls.adodb.adLongVarWChar,
        com_dlls.adodb.adVarChar,
        com_dlls.adodb.adVarWChar,
        com_dlls.adodb.adWChar,
    )
    BYNARY = (
        com_dlls.adodb.adBinary,
        com_dlls.adodb.adLongVarBinary,
        com_dlls.adodb.adVarBinary,
    )
    DATETIME = (
        com_dlls.adodb.adDBTime,
        com_dlls.adodb.adDBTimeStamp,
        com_dlls.adodb.adDate,
        com_dlls.adodb.adDBDate,
    )
    OTHER = (
        com_dlls.adodb.adEmpty,
        com_dlls.adodb.adIDispatch,
        com_dlls.adodb.adIUnknown,
        com_dlls.adodb.adPropVariant,
        com_dlls.adodb.adArray,
        com_dlls.adodb.adUserDefined,
        com_dlls.adodb.adVariant,
        com_dlls.adodb.adGUID,
    )


class _DbApiColumnType:
    """Uses to describe columns in a database that are specified type.

    See also:
    - https://peps.python.org/pep-0249/#type-objects-and-constructors
    - https://peps.python.org/pep-0249/#string
    - https://peps.python.org/pep-0249/#binary
    - https://peps.python.org/pep-0249/#number
    - https://peps.python.org/pep-0249/#datetime
    - https://peps.python.org/pep-0249/#rowid
    """

    def __init__(self, values: Sequence[int]):
        self.values = frozenset(values)

    def __eq__(self, other: int) -> bool:
        return other in self.values


STRING = _DbApiColumnType(TypeConstants.STRING)
BINARY = _DbApiColumnType(TypeConstants.BYNARY)
NUMBER = _DbApiColumnType(
    TypeConstants.INTEGER
    + TypeConstants.LONG
    + TypeConstants.EXACT_NUMERIC
    + TypeConstants.APPROX_NUMERIC
)
DATETIME = _DbApiColumnType(TypeConstants.DATETIME)
ROWID = _DbApiColumnType(TypeConstants.ROWID)


def Date(year: SupportsIndex, month: SupportsIndex, day: SupportsIndex) -> Any:
    raise NotImplementedError


def Time(hour: SupportsIndex, minute: SupportsIndex, second: SupportsIndex) -> Any:
    raise NotImplementedError


def Timestamp(
    year: SupportsIndex,
    month: SupportsIndex,
    day: SupportsIndex,
    hour: SupportsIndex,
    minute: SupportsIndex,
    second: SupportsIndex,
) -> Any:
    raise NotImplementedError


def DateFromTicks(ticks: float) -> Date:
    raise NotImplementedError


def TimeFromTicks(ticks: float) -> Time:
    raise NotImplementedError


def TimestampFromTicks(ticks: float) -> Timestamp:
    raise NotImplementedError
