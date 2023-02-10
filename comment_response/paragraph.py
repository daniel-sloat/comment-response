from dataclasses import dataclass

from xlsx.cell.run import Run


@dataclass
class Paragraph:
    runs: list[Run]

    def __post_init__(self):
        self.runs = list(self.runs)
        for run in self.runs:
            if run == self.runs[0]:
                run.text = run.text.lstrip()
            if run == self.runs[-1]:
                run.text = run.text.rstrip()

    def __repr__(self):
        return f"{self.__class__.__name__}({self.runs})"

    def __getitem__(self, key):
        return self.runs[key]

    def __iter__(self):
        return iter(self.runs)


@dataclass
class Paragraphs:
    paragraphs: list[Paragraph]

    def __repr__(self):
        return f"{self.__class__.__name__}({self.paragraphs})"

    def __iter__(self):
        return iter(self.paragraphs)
