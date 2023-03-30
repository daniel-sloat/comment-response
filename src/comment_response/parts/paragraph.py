"""Paragraph class"""

import re

from dataclasses import dataclass

from xlsx_rich_text.cell.run import Run


@dataclass
class Paragraph:
    """Paragraph: a list of runs. The first run is trimmed of leading whitespace and
    the last run is trimmed of trailing whitespace."""

    runs: list[Run]
    trim: bool = True
    clean: bool = True

    def __post_init__(self):
        for run in self.runs:
            if self.clean:
                run.text = re.sub(r"[^\S\n]+", " ", run.text)

            if self.trim:
                if run == self.runs[0]:
                    run.text = run.text.lstrip()
                if run == self.runs[-1]:
                    run.text = run.text.rstrip()
