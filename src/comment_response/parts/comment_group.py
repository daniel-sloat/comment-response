"""Provides access to grouped record data."""

from xlsx_rich_text.sheets.record import Record

from comment_response.parts.comment import Comment
from comment_response.parts.response import Response


class CommentGroup:
    """Group of comments. There may be one or more comments in a comment group, but
    only one response. If multiple responses are found, they will be grouped into one.
    """

    def __init__(self, records: list[Record], config: dict):
        self.records = records
        self.config = config
        self.comment_col = config["columns"]["commentresponse"]["comment"]

    @property
    def comments(self) -> list[Comment]:
        """List of comments. Empty comments are not included."""
        cmts = []
        for record in self.records:
            cmt = Comment(record, self.config)
            if cmt:
                cmts.append(cmt)
        return cmts

    @property
    def response(self) -> Response:
        """Singular response to comments."""
        return Response(self.records, self.config)
