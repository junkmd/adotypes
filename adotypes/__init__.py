import logging

from adotypes.api_constructors import connect  # noqa
from adotypes import com_dlls  # noqa
from adotypes.api_objects import (  # noqa
    Connection,
    Cursor,
    STRING,
    BINARY,
    NUMBER,
    DATETIME,
    ROWID,
    Date,
    Time,
    Timestamp,
    DateFromTicks,
    TimeFromTicks,
    TimestampFromTicks,
)
from adotypes.api_globals import (  # noqa
    apilevel,
    threadsafety,
    paramstyle,
)
from adotypes.api_exceptions import (  # noqa
    DatabaseError,
    DataError,
    Error,
    IntegrityError,
    InterfaceError,
    InternalError,
    NotSupportedError,
    OperationalError,
    ProgrammingError,
    Warning,
)


LOGGER = logging.getLogger(__name__)


class NullHandler(logging.Handler):
    def emit(self, record: logging.LogRecord) -> None:
        pass


LOGGER.addHandler(NullHandler())
