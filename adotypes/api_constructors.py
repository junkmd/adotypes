import logging
from typing import Any, NoReturn, Optional, overload

import comtypes.client

from adotypes import com_dlls, api_exceptions as exc
from adotypes.api_objects import Connection

LOGGER = logging.getLogger(__name__)


# fmt: off
@overload
def connect(*, create: str, **kwargs: Any) -> Connection: ...  # noqa
@overload
def connect(*, open: str, user_id: str = ..., passward: str = ..., options: str = ..., **kwargs: Any) -> Connection: ...  # noqa
@overload
def connect(*, create: str, open: str, **kwargs: Any) -> NoReturn: ...  # noqa
# fmt: on


def connect(**kwargs: Any) -> Connection:
    conn, kw = _new_adodb_connection(**kwargs)
    return Connection(conn, **kw)


def _new_adodb_connection(
    create: Optional[str] = None, open: Optional[str] = None, **kwargs: Any
) -> tuple[com_dlls.adodb._Connection, dict[str, Any]]:
    if create and open:
        raise TypeError
    if create:
        LOGGER.debug("start creating a new db using adox")
        try:
            catalog = comtypes.client.CreateObject(
                com_dlls.adox.Catalog, interface=com_dlls.adox._Catalog
            )
            catalog.Create(create)
            conn = catalog.ActiveConnection.QueryInterface(com_dlls.adodb._Connection)
            LOGGER.debug("complete creating a new db using adox")
            return (conn, kwargs)
        except Exception as e:
            LOGGER.error(str(e), stack_info=True)
            raise exc.ProgrammingError from e
    if open:
        LOGGER.debug("start opening connection to an existing db using adodb")
        try:
            conn = comtypes.client.CreateObject(
                com_dlls.adodb.Connection, interface=com_dlls.adodb._Connection
            )
            user_id = kwargs.pop("user_id", "")
            passward = kwargs.pop("passward", "")
            options = kwargs.pop("options", com_dlls.adodb.adConnectUnspecified)
            conn.Open(open, user_id, passward, options)
            LOGGER.debug("complete opening connection to an existing db using adodb")
            return (conn, kwargs)
        except Exception as e:
            LOGGER.error(str(e), stack_info=True)
            raise exc.ProgrammingError from e
    raise TypeError
