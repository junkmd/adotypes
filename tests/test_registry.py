from unittest.mock import call, MagicMock
from winreg import HKEYType

import pytest
from pytest_mock import MockerFixture as _Mocker

from adotypes import registry


_GUID = "{xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx}"


class Test:
    @pytest.fixture
    def _winreg(self, mocker: _Mocker) -> MagicMock:
        return mocker.patch.object(registry, "winreg")

    @pytest.fixture
    def open_key(self, mocker: _Mocker, _winreg: MagicMock) -> MagicMock:
        m = mocker.MagicMock(spec=HKEYType)
        _winreg.OpenKey.return_value.__enter__.return_value = m
        return m

    @pytest.fixture
    def qi_key(self, _winreg: MagicMock) -> MagicMock:
        m = _winreg.QueryInfoKey
        m.return_value = (3, ..., ...)
        return m

    @pytest.fixture
    def enum_key(self, _winreg: MagicMock) -> MagicMock:
        m = _winreg.EnumKey
        m.side_effect = ["10.b", "2.11", "10.a"]
        return m

    def _assert_mocks_called(self, open_key, qi_key, enum_key) -> None:
        qi_key.assert_called_once_with(open_key)
        call_args = [call(open_key, 0), call(open_key, 1), call(open_key, 2)]
        assert enum_key.call_args_list == call_args

    def test_not_takes_order(self, open_key, qi_key, enum_key):
        expected = [(_GUID, 16, 11), (_GUID, 2, 17), (_GUID, 16, 10)]
        assert expected == registry.get_all_typelib_versions(_GUID)
        self._assert_mocks_called(open_key, qi_key, enum_key)

    @pytest.mark.parametrize(
        "order, expected",
        [
            ("default", [(_GUID, 16, 11), (_GUID, 2, 17), (_GUID, 16, 10)]),
            ("asc", [(_GUID, 2, 17), (_GUID, 16, 10), (_GUID, 16, 11)]),
            ("desc", [(_GUID, 16, 11), (_GUID, 16, 10), (_GUID, 2, 17)]),
        ],
    )
    def test_takes_order(self, open_key, qi_key, enum_key, order, expected):
        assert expected == registry.get_all_typelib_versions(_GUID, order=order)
        self._assert_mocks_called(open_key, qi_key, enum_key)

    def test_takes_invalid_order(self, open_key, qi_key, enum_key):
        with pytest.raises(TypeError):
            registry.get_all_typelib_versions(_GUID, order="foo")  # type: ignore
