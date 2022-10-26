import re
NORMALIZED_INTERFACES = (
    "Fa",
    "Gi",
    "Te",
    "Fo",
    "Et",
    "Lo",
    "Se",
    "Vlan",
    "Tunnel",
    "Portchannel",
    "Management",
)

INTERFACE_NAME_RE = re.compile(
    r"(?P<interface_type>[a-zA-Z\-_ ]*)(?P<interface_num>[\d.\/]*)"
)


