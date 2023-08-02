from collections.abc import Mapping, Sequence
from typing import Optional, Protocol, TypeVar, Union


_T_co = TypeVar("_T_co", covariant=True)


class SizedSupportsGetItem(Protocol[_T_co]):
    # fmt: off
    def __len__(self) -> int: ...  # noqa
    def __getitem__(self, __k: int) -> _T_co: ...  # noqa
    # fmt: on


_Parameters = Union[SizedSupportsGetItem[_T_co], Mapping[str, _T_co]]

_ColumnDescription = tuple[
    Optional[str],  # name
    Optional[int],  # type_code
    Optional[int],  # display_size
    Optional[int],  # internal_size
    Optional[int],  # precision
    Optional[int],  # scale
    Optional[int],  # null_ok
]

_InputSizes = Union[int, Sequence[Optional[int]]]
_OutputSize = Union[int, tuple[int, Optional[int]]]
