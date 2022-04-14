#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import pandas as pd

class RichText(): #//TODO Convert to regular functions
    """Returns the rich-formatted shared strings associated with the 
    shared string index code value.
    """
    def __init__(self, ss: list, codes: list) -> None:
        self.ss = ss
        self.codes = codes
        pass
    
    @staticmethod
    def __initialize_df(sharedstrings_rich: list) -> pd.DataFrame:
        df = pd.DataFrame(sharedstrings_rich, 
                          columns=["Index","Style","String"])
        return df
    
    @staticmethod
    def __groupby_group_and_run(df: pd.DataFrame) -> pd.DataFrame:
        # Add run number column. A run is consolidated when it has the same
        # formatting and is of the same group.
        df["RunNo"] = (df["Index"].shift(1).ne(df["Index"]) | 
                       df["Style"].shift(1).ne(df["Style"])
                       ).cumsum()
        df = df.groupby(["Index","RunNo","Style"], 
                        as_index=False)["String"].apply("".join)
        return df
    
    def __rich_strings_df(self) -> pd.DataFrame:
        df = RichText.__initialize_df(self.ss)
        df = RichText.__groupby_group_and_run(df)
        df = df[df["Index"].astype(str).isin(self.codes)].copy()
        return df

    def __create_code_dict(self) -> dict:
        df = self.__rich_strings_df()
        df["Run"] = tuple(zip(df["Style"],df["String"]))  # Group by run
        df = df.groupby(["Index"])["Run"].agg(list)  # Group by para (not implemented) //TODO
        df = df.groupby(["Index"]).agg(tuple)  # Group by comment (i.e., cell)
        d = df.to_dict()
        return d
    
    def formats_used(self) -> list[str]:
        formats = self.__rich_strings_df()["Style"]
        return list(set(formats))

    def decode(self) -> list:
        replaced_column_data = []
        d = self.__create_code_dict()
        for code in self.codes:
            if isinstance(code, str):
                code_int = int(code)
                if code_int in d:
                    replaced_column_data.append(d[code_int])
            else:
                replaced_column_data.append(code)
        return replaced_column_data