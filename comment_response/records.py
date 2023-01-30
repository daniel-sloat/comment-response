"""Provides access to grouped record data."""


class Records:
    """Group of records."""

    def __init__(self, records, config):
        self.records = records
        self.comment_col = config["columns"]["commentresponse"]["comment"]
        self.response_col = config["columns"]["commentresponse"]["response"]

    def __repr__(self):
        return f"Records(count={len(self.records)})"

    @property
    def comments(self):
        new = []
        for record in self.records:
            com = record.col.get(self.comment_col)
            if com:  # Should be one comment per row, but just in case
                rich_text = com.value
                if rich_text:
                    new.append(rich_text.paragraphs)
        return new

    @property
    def response(self):
        new = []
        for record in self.records:
            resp = record.col.get(self.response_col)
            if resp:
                rich_text = resp.value
                if rich_text:
                    new.extend(rich_text.paragraphs)
        return new
