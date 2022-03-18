#!/usr/bin/env python3.10
# -*- coding: utf-8 -*-

import pandas as pd

class RichText():
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
                          columns=["Group","Newline","Style","String"])
        return df
    
    @staticmethod
    def __groupby_group_and_run(df: pd.DataFrame) -> pd.DataFrame:
        # Add run number column. If next row indicates a newline character, 
        # or next group is incremented, then also increment run number. 
        # Consolidates sequential rows into one if the runs have the same
        # formatting, are of the same group, and there isn't a newline 
        # character. 
        df["RunNo"] = (df["Group"].shift(1).ne(df["Group"]) | 
                       df["Newline"].shift(1).gt(0) | 
                       df["Style"].shift(1).ne(df["Style"])
                       ).cumsum()
        df = df.drop(["Newline"], axis=1).reindex(
            columns=["Group","RunNo","Style","String"])
        df = df.groupby(["Group","RunNo","Style"], 
                        as_index=False)["String"].apply("".join)
        return df
    
    def __rich_strings_df(self) -> pd.DataFrame:
        df = RichText.__initialize_df(self.ss)
        df = RichText.__groupby_group_and_run(df)
        df = df[df["Group"].astype(str).isin(self.codes)].copy()
        return df

    def __create_code_dict(self) -> dict:
        df = self.__rich_strings_df()
        df["StyleString"] = df[["Style","String"]].values.tolist()
        #df_strings = df_strings.drop(["Style","String"], axis=1)
        df = df.groupby(["Group"], as_index=False)["StyleString"].agg(list)
        df = df.groupby(["Group"])["StyleString"].agg(list)
        d = df.to_dict()
        return d
    
    def formats_used(self) -> list[str]:
        formats = self.__rich_strings_df()["Style"]
        return list(set(formats))

    def decode(self) -> list:
        replaced_column_data = []
        d = self.__create_code_dict()
        for code in self.codes:
            if type(code) == str:
                if int(code) in d.keys():
                    replaced_column_data.append(d[int(code)])
            else:
                replaced_column_data.append(code)
        return replaced_column_data