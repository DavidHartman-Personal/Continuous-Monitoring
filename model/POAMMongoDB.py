import datetime
from mongoengine import *

class Comment(EmbeddedDocument):
    content = StringField()
    date_added = DateTimeField(default=datetime.datetime.now)

class SystemResourceIdentifier(EmbeddedDocument):
    identifier = StringField()
    identifier_type = StringField()

class POAM(Document):
    poam_id = StringField(required=True, primary_key=True)
    controls = StringField()
    weakness_name = StringField()
    weakness_description = StringField()
    weakness_detector_source = StringField()
    status = StringField()
    status_date = DateTimeField()
    comments = StringField()
    system_identifiers = ListField(EmbeddedDocumentField(SystemResourceIdentifier))
    history_comments = ListField(EmbeddedDocumentField(Comment))
    date_added = DateTimeField(default=datetime.datetime.now())
    date_updated = DateTimeField(default=datetime.datetime.now())
    current_flag = BooleanField(default=True)
    delete_flag = BooleanField(default=False)
