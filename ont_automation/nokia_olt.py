"""Nokia OLT connector facade — combines all OLT/ONT operation mixins."""

from ont_automation.olt_connection import OLTConnectionMixin
from ont_automation.ont_cleanup import ONTCleanupMixin
from ont_automation.ont_lookup import ONTLookupMixin
from ont_automation.ont_provisioning import ONTProvisioningMixin
from ont_automation.ont_registration import ONTRegistrationMixin


class NokiaOLTConnector(
    OLTConnectionMixin,
    ONTLookupMixin,
    ONTRegistrationMixin,
    ONTProvisioningMixin,
    ONTCleanupMixin,
):
    """Unified Nokia OLT client for connection, lookup, registration, provisioning, and cleanup."""

    def __init__(self, ip: str, user: str, pw: str, proto: str):
        OLTConnectionMixin.__init__(self, ip, user, pw, proto)
