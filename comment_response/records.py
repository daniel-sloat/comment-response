"""Provides access to grouped record data."""

from comment_response.comment import Comments
from comment_response.response import Response
from xlsx.sheets.record import Record


class Records:
    """Group of records."""

    def __init__(self, records: list[Record], config: dict):
        self.records = records
        self.config = config

    def __repr__(self):
        return f"Records(count={len(self.records)})"

    @property
    def comments(self):
        return Comments(self.records, self.config).prepared()

    @property
    def response(self):
        return Response(self.records, self.config).prepared()
