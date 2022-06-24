import re

from dataclasses import dataclass, field


@dataclass
class Run:
    props: str
    text: str


@dataclass
class Paragraph:
    runs: list[Run] = field(default_factory=list)


@dataclass
class xlRichText:
    runs: list[Run] = field(default_factory=list)


@dataclass
class RichText:
    paragraphs: list[Paragraph] = field(default_factory=list)


@dataclass
class Cell:
    col: int
    row: int
    col_name: str = None
    value: bool | int | float | str | list = None
    xl_rich: xlRichText = None
    xl_dtype: str = None
    format: str = None
    style: str = None

    @property
    def position(self) -> tuple[int, int]:
        return self.col, self.row

    @property
    def rich(self) -> RichText:
        try:
            split_runs = [
                Run(run.props, t)
                for run in self.xl_rich
                for t in re.split(r"(\n)", run.text)
            ]
        except TypeError:
            return RichText()
        para, rich_text = Paragraph(), RichText()
        last = split_runs[-1]
        for run in split_runs:
            if run.text:
                if run.text != "\n":
                    para.runs.append(run)
                    if run == last:
                        rich_text.paragraphs.append(para)
                else:
                    if not para.runs:
                        continue
                    rich_text.paragraphs.append(para)
                    para = Paragraph()
        return rich_text


@dataclass
class SheetData:
    data: list[Cell] = field(default_factory=list)

    def add(self, *d):
        self.data += d


@dataclass
class CommentResponseData:
    tags: str
    headings: list[str]
    sort: list[int | str]
    comment: RichText = field(default_factory=RichText())
    response: RichText = field(default_factory=RichText())
