from typing import Literal
import winreg


def get_all_typelib_versions(
    libid: str, *, order: Literal["default", "asc", "desc"] = "default"
) -> list[tuple[str, int, int]]:
    result: list[tuple[str, int, int]] = []
    with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, rf"TypeLib\{libid}") as key:
        cnt, _, _ = winreg.QueryInfoKey(key)
        for i in range(cnt):
            mj, mn, *_ = (int(v, base=16) for v in winreg.EnumKey(key, i).split("."))
            result.append((libid, mj, mn))
    if order == "asc":
        return sorted(result, key=lambda r: (r[1], r[2]))
    elif order == "desc":
        return sorted(result, key=lambda r: (r[1], r[2]), reverse=True)
    elif order == "default":
        return result
    raise TypeError
