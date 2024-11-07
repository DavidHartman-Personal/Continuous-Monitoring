import datetime
from mongoengine import *

class Comment(EmbeddedDocument):
    content = StringField()
    date_added = DateTimeField(default=datetime.datetime.now)

class SystemResourceIdentifier(EmbeddedDocument):
    identifier = StringField()
    identifier_type = StringField()

class SystemResource(Document):
    name = StringField(required=True, primary_key=True)
    primary_ip_address = StringField()
    status = StringField()
    additional_ip_addresses = StringField()
    environment = StringField()
    function = StringField()
    system_administrator_owner = StringField()
    application_administrator_owner = StringField()
    unique_asset_identifier = StringField()
    ipv4_or_ipv6_address = StringField()
    virtual = StringField()
    public = StringField()
    dns_name_or_url = StringField()
    netbios_name = StringField()
    mac_address = StringField()
    authenticated_scan = StringField()
    baseline_configuration_name = StringField()
    os_name_and_version = StringField()
    location = StringField()
    asset_type = StringField()
    hardware_make_model = StringField()
    in_latest_scan = StringField()
    software_database_vendor = StringField()
    software_database_name_version = StringField()
    patch_level = StringField()
    comments = StringField()
    serial_number_asset_tag_number = StringField()
    vlan_network_id = StringField()
    system_identifiers = ListField(EmbeddedDocumentField(SystemResourceIdentifier))
    history_comments = ListField(EmbeddedDocumentField(Comment))
    current_flag = BooleanField(default=True)
    delete_flag = BooleanField(default=False)
