import comtypes.client

from adotypes.registry import get_all_typelib_versions


comtypes.client.GetModule(
    get_all_typelib_versions("{B691E011-1797-432E-907A-4D8C69339129}", order="desc")[0]
)
from comtypes.gen import ADODB as adodb  # noqa


comtypes.client.GetModule(
    get_all_typelib_versions("{00000600-0000-0010-8000-00AA006D2EA4}", order="desc")[0]
)
from comtypes.gen import ADOX as adox  # noqa
