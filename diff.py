import sys
import os
import json
import re
import pandas as pd
from pathlib import Path
from datetime import datetime


class ExcelDiff:
    def __init__(
        self,
        path_old,
        path_new,
        dangerous_drop_rules=None,
        important_change_rules=None,
        filename=None,
    ):
        self.path_old = path_old
        self.path_new = path_new
        self.filename = filename
        self.new_rows = None
        self.df_new = None
        self.df_old = None
        self.fields = None
        self.dropped_rows = None
        self.dangerous_dropped_rows = None
        self.changes = None
        self.writer = None
        self.workbook = None
        self.formats = {}
        self.dangerous_drop_rules = dangerous_drop_rules
        self.important_change_rules = important_change_rules

    def diff(self, verbose=False):
        dfs = []
        for path in [self.path_old, self.path_new]:
            ext = os.path.splitext(path)[1]
            if ext == ".csv":
                df = pd.read_csv(path).fillna("")
            elif ext == ".xlsx":
                df = pd.read_excel(path).fillna("")
            else:
                print(ext)
                raise Exception("File must be .csv or .xlsx")
            df = df.loc[:, ~df.columns.str.contains("^Unnamed")]
            dfs.append(df)

        [df_old, df_new] = dfs
        self.fields = df_new.loc[:, "Variable / Field Name"].values
        self.df_new = df_new
        self.df_old = df_old
        if not self.filename:
            self.filename = (
                f"DataDictionary_{datetime.now().strftime('%m-%d-%Y-%I%M%p')}"
            )
        # Save output and format
        fname = "{}.xlsx".format(self.filename)
        # TODO: Figure out how to handle date formatting/changes
        self.writer = pd.ExcelWriter(
            fname,
            engine="xlsxwriter",
            date_format="MM/DD/YY",
            datetime_format="MM/DD/YY",
        )
        # get xlsxwriter objects
        self.workbook = self.writer.book
        self.formats["new"] = self.workbook.add_format({"bg_color": "#90EE90"})
        self.formats["dropped"] = self.workbook.add_format({"bg_color": "#ff9999"})
        self.formats["changed"] = self.workbook.add_format({"bg_color": "#ffff66"})
        self.formats["important_changed"] = self.workbook.add_format(
            {"bg_color": "#ffB347"}
        )
        self.formats["header"] = self.workbook.add_format(
            {"border": 1, "bg_color": "#DCDCDC"}
        )
        self.formats["bold"] = self.workbook.add_format({"bold": True})
        self.formats["wrap"] = self.workbook.add_format({"text_wrap": True})

        if self.df_old.columns.tolist() == self.df_new.columns.tolist():
            return self.simple_diff(verbose=verbose)
        return self.complex_diff(verbose=verbose)

    def simple_diff(self, verbose=False):
        df_new_final = self.df_new.copy()
        df_diff = self.df_new.copy()
        self.dropped_rows = []
        self.new_rows = []
        self.changes = {}
        old_fields = self.df_old.loc[:, "Variable / Field Name"].values
        for ind in df_diff.index:
            field = df_diff.loc[ind, "Variable / Field Name"]
            if field in old_fields:
                # The field is not new
                old_row = self.df_old[self.df_old["Variable / Field Name"] == field]
                for col_ind, col in enumerate(self.df_new.columns):
                    if df_diff.loc[ind, col] != old_row[col].values[0]:
                        # The field was updated
                        if field not in self.changes:
                            self.changes[field] = {
                                "field": field,
                                "row_num": ind + 1,
                                "changed_cols": [],
                                "old_row_num": old_row.index[0] + 1,
                            }
                        col_change_dict = {
                            "col_name": col,
                            "col_num": col_ind,
                            "val": df_diff.loc[ind, col],
                        }
                        if col == "Choices, Calculations, OR Slider Labels":
                            col_change_dict["old_options"] = [
                                opt.strip().split(",", 1)[-1]
                                for opt in old_row[col].values[0].split("|")
                                if opt
                            ]

                            col_change_dict["new_options"] = [
                                opt.strip().split(",", 1)[-1]
                                for opt in df_diff.loc[ind, col].split("|")
                                if opt
                            ]
                        else:
                            col_change_dict["old_val"] = old_row[col].values[0]
                            col_change_dict["new_val"] = df_diff.loc[ind, col]
                        self.changes[field]["changed_cols"].append(col_change_dict)

            else:
                # The field is new
                self.new_rows.append({"field": field, "row_num": ind + 1})

        for ind in self.df_old.index:
            field = self.df_old.loc[ind, "Variable / Field Name"]
            if field not in self.fields:
                dropped_dict = {
                    "field": field,
                    "old_row_num": ind + 1,
                    "diff_row_num": df_diff.shape[0] + 1,
                }
                self.dropped_rows.append(dropped_dict)
                df_diff = df_diff.append(self.df_old.loc[ind, :], ignore_index=True)

        df_diff.fillna("").to_excel(self.writer, sheet_name="DIFF", index=False)
        df_new_final.fillna("").to_excel(self.writer, sheet_name="NEW", index=False)
        self.df_old.fillna("").to_excel(self.writer, sheet_name="OLD", index=False)
        pd.DataFrame().to_excel(self.writer, sheet_name="CHANGE_NOTES", index=False)

        worksheet1 = self.writer.sheets["DIFF"]
        worksheet2 = self.writer.sheets["NEW"]
        worksheet3 = self.writer.sheets["OLD"]
        worksheet4 = self.writer.sheets["CHANGE_NOTES"]

        new_row_inds = [row_data["row_num"] for row_data in self.new_rows]
        dropped_row_inds = [row_data["diff_row_num"] for row_data in self.dropped_rows]

        for row in range(df_diff.shape[0] + 1):
            if row in new_row_inds:
                worksheet1.set_row(row, 15, self.formats["new"])
            if row in dropped_row_inds:
                worksheet1.set_row(row, 15, self.formats["dropped"])

        for field, row_dict in self.changes.items():
            row_num = row_dict["row_num"]
            column_dict_list = row_dict["changed_cols"]
            for column_dict in column_dict_list:
                col_num = column_dict["col_num"]
                column_value = column_dict["val"]
                col_name = column_dict["col_name"]
                worksheet1.write(
                    row_num, col_num, column_value, self.formats["changed"]
                )

        worksheet1.set_column("A:Z", 30)
        worksheet2.set_column("A:Z", 30)
        worksheet3.set_column("A:Z", 30)

        self.create_changes_sheet(worksheet4)
        self.create_new_changes_sheet()

        self.writer.save()
        self.writer.close()

        if verbose:
            print(df_diff.shape)
            print("New Rows:")
            for row in self.new_rows:
                print(f"Field: {row['field']}, Row Number: {row['row_num']+1}")
            print("Dropped Rows:")
            for row in self.dropped_rows:
                print(
                    f"Field: {row['field']},\t\
                    Diff Row Number: {row['diff_row_num']+1},\t\
                        Old Row Number: {row['old_row_num']+1}"
                )
            print("Changed Rows:")
            for row, data in self.changes.items():
                print(
                    f"Name: {row}, Row Number: {data['row_num']+1}, Old Row Number: {data['old_row_num']+1}"
                )
                # print("Changed Columns:")
                for column in data["changed_cols"]:
                    print(f"Column: {column['col_name']}")
                print("*********")

    def complex_diff(self, required_cols_in_master=18, verbose=False):
        # self.df_old is Molly's spreadsheet (w/ 5 extra columns)
        # df_new is the latest master data dictionary

        if not self.dangerous_drop_rules:
            self.dangerous_drop_rules = {"Who requested this data?": ["CCDE"]}
        if not self.important_change_rules:
            self.important_change_rules = {
                "fields": ["Field Type", "Choices, Calculations, OR Slider Labels"]
            }

        if self.df_new.shape[1] != required_cols_in_master:
            err = f"The supplied new master data dictionary does not have the required number of columns. It has {self.df_new.shape[1]}, but needs {required_cols_in_master}."
            raise ValueError(err)

        df_new_final = self.df_new.copy()
        df_diff = self.df_new.copy()
        self.dropped_rows = []
        self.dangerous_dropped_rows = []
        self.new_rows = []
        self.changes = {}

        old_fields = self.df_old.loc[:, "Variable / Field Name"].values
        for ind in df_diff.index:
            field = df_diff.loc[ind, "Variable / Field Name"]
            if field in old_fields:
                # The field is not new
                old_row = self.df_old[self.df_old["Variable / Field Name"] == field]
                for additional_col in self.df_old.columns[required_cols_in_master:]:
                    # Add in Molly's columns at the appropriate index
                    df_diff.loc[ind, additional_col] = old_row[additional_col].values[0]
                    df_new_final.loc[ind, additional_col] = old_row[
                        additional_col
                    ].values[0]
                for col_ind, col in enumerate(self.df_new.columns):
                    if df_diff.loc[ind, col] != old_row[col].values[0]:
                        # The field was updated
                        if field not in self.changes:
                            self.changes[field] = {
                                "field": field,
                                "row_num": ind + 1,
                                "field_requester": old_row[
                                    "Who requested this data?"
                                ].values[0],
                                "changed_cols": [],
                                "old_row_num": old_row.index[0] + 1,
                            }
                        col_change_dict = {
                            "col_name": col,
                            "col_num": col_ind,
                            "val": df_diff.loc[ind, col],
                        }
                        if col == "Choices, Calculations, OR Slider Labels":
                            col_change_dict["old_options"] = [
                                opt.strip().split(",", 1)[-1]
                                for opt in old_row[col].values[0].split("|")
                                if opt
                            ]

                            col_change_dict["new_options"] = [
                                opt.strip().split(",", 1)[-1]
                                for opt in df_diff.loc[ind, col].split("|")
                                if opt
                            ]
                        else:
                            col_change_dict["old_val"] = old_row[col].values[0]
                            col_change_dict["new_val"] = df_diff.loc[ind, col]
                        self.changes[field]["changed_cols"].append(col_change_dict)

            else:
                # The field is new
                self.new_rows.append({"field": field, "row_num": ind + 1})

        for ind in self.df_old.index:
            field = self.df_old.loc[ind, "Variable / Field Name"]
            if field not in self.fields:
                dropped_dict = {
                    "field": field,
                    "old_row_num": ind + 1,
                    "diff_row_num": df_diff.shape[0] + 1,
                    "field_requester": self.df_old.loc[ind, "Who requested this data?"],
                }
                self.dropped_rows.append(dropped_dict)
                for field, value_arr in self.dangerous_drop_rules.items():
                    if self.df_old.loc[ind, field] in value_arr:
                        self.dangerous_dropped_rows.append(dropped_dict)
                df_diff = df_diff.append(self.df_old.loc[ind, :], ignore_index=True)

        df_diff.fillna("").to_excel(self.writer, sheet_name="DIFF", index=False)
        df_new_final.fillna("").to_excel(self.writer, sheet_name="NEW", index=False)
        self.df_old.fillna("").to_excel(self.writer, sheet_name="OLD", index=False)
        pd.DataFrame().to_excel(self.writer, sheet_name="CHANGE_NOTES", index=False)

        # get xlsxwriter objects
        self.workbook = self.writer.book
        worksheet1 = self.writer.sheets["DIFF"]
        worksheet2 = self.writer.sheets["NEW"]
        worksheet3 = self.writer.sheets["OLD"]
        worksheet4 = self.writer.sheets["CHANGE_NOTES"]

        new_row_inds = [row_data["row_num"] for row_data in self.new_rows]
        dropped_row_inds = [row_data["diff_row_num"] for row_data in self.dropped_rows]

        for row in range(df_diff.shape[0] + 1):
            if row in new_row_inds:
                worksheet1.set_row(row, 15, self.formats["new"])
            if row in dropped_row_inds:
                worksheet1.set_row(row, 15, self.formats["dropped"])

        for field, row_dict in self.changes.items():
            row_num = row_dict["row_num"]
            column_dict_list = row_dict["changed_cols"]
            for column_dict in column_dict_list:
                col_num = column_dict["col_num"]
                column_value = column_dict["val"]
                col_name = column_dict["col_name"]
                fmt = self.formats["changed"]
                if col_name in self.important_change_rules["fields"]:
                    fmt = self.formats["important_changed"]
                worksheet1.write(row_num, col_num, column_value, fmt)

        worksheet1.set_column("A:Z", 30)
        worksheet2.set_column("A:Z", 30)
        worksheet3.set_column("A:Z", 30)
        worksheet4.set_column("A:Z", 30)

        self.create_changes_sheet(worksheet4)

        # Add Molly's additional sheets in
        df_missing_different = pd.read_excel(
            self.path_old, sheet_name="Missing or changed from CCDE"
        ).fillna("")
        df_missing_different.to_excel(
            self.writer, sheet_name="Missing_different CCDE fields", index=False
        )
        df_key = pd.read_excel(self.path_old, sheet_name="Key").fillna("")
        df_key.to_excel(self.writer, sheet_name="Key", index=False)

        self.writer.save()
        self.writer.close()

        if verbose:
            print(df_diff.shape)
            print("New Rows:")
            for row in self.new_rows:
                print(f"Field: {row['field']}, Row Number: {row['row_num']+1}")
            print("Dropped Rows:")
            for row in self.dropped_rows:
                print(
                    f"Field: {row['field']},\t\
                    Diff Row Number: {row['diff_row_num']+1},\t\
                        Old Row Number: {row['old_row_num']+1},\t\
                            Dangerous? {row in self.dangerous_dropped_rows}"
                )
            print("Changed Rows:")
            for row, data in self.changes.items():
                print(
                    f"Name: {row}, Row Number: {data['row_num']+1}, Old Row Number: {data['old_row_num']+1}"
                )
                # print("Changed Columns:")
                for column in data["changed_cols"]:
                    print(
                        f"Column: {column['col_name']},\tImportant? {column['col_name'] in self.important_change_rules['fields']}"
                    )
                print("*********")

    def create_changes_sheet(self, worksheet):
        worksheet.set_column("A:A", 30)
        worksheet.set_column("B:Z", 15)

        worksheet.write(0, 0, "Change Notes")

        # New Rows
        start = 2
        worksheet.write(start, 0, "New Rows:", self.formats["bold"])
        worksheet.write(start + 1, 0, "Variable / Field Name", self.formats["header"])
        worksheet.write(start + 1, 1, "Row Number", self.formats["header"])
        start += 2
        for ind, row in enumerate(self.new_rows):
            worksheet.write(start + ind, 0, row["field"])
            worksheet.write(start + ind, 1, row["row_num"] + 1)
        start += len(self.new_rows) + 2

        if self.dangerous_drop_rules:
            # Dangerous Dropped Rows
            worksheet.write(start, 0, "Dangerous Dropped Rows:", self.formats["bold"])
            worksheet.write(
                start, 1, json.dumps(self.dangerous_drop_rules), self.formats["bold"]
            )
            worksheet.write(
                start + 1, 0, "Variable / Field Name", self.formats["header"]
            )
            worksheet.write(start + 1, 1, "Old Row Number", self.formats["header"])
            worksheet.write(start + 1, 2, "Diff Row Number", self.formats["header"])
            worksheet.write(start + 1, 3, "Field Requester", self.formats["header"])
            start += 2
            for ind, row in enumerate(self.dangerous_dropped_rows):
                worksheet.write(start + ind, 0, row["field"])
                worksheet.write(start + ind, 1, row["old_row_num"] + 1)
                worksheet.write(start + ind, 2, row["diff_row_num"] + 1)
                worksheet.write(start + ind, 3, row["field_requester"])
            start += len(self.dangerous_dropped_rows) + 2

        if self.important_change_rules:
            # Important Cell Changes
            worksheet.write(start, 0, "Important Changes:", self.formats["bold"])
            worksheet.write(
                start, 1, json.dumps(self.important_change_rules), self.formats["bold"]
            )
            for ind, (field, change_dict) in enumerate(self.changes.items()):
                important_changes = [
                    i
                    for i in change_dict["changed_cols"]
                    if i["col_name"] in self.important_change_rules["fields"]
                ]
                if important_changes and change_dict["field_requester"] == "CCDE":
                    worksheet.write(
                        start + 1, 0, "Variable / Field Name", self.formats["header"]
                    )
                    worksheet.write(
                        start + 1, 1, "Old Row Number", self.formats["header"]
                    )
                    worksheet.write(
                        start + 1, 2, "New Row Number", self.formats["header"]
                    )
                    worksheet.write(
                        start + 1, 3, "Field Requester", self.formats["header"]
                    )
                    start += 2
                    worksheet.write(start, 0, field)
                    worksheet.write(start, 1, change_dict["old_row_num"] + 1)
                    worksheet.write(start, 2, change_dict["row_num"] + 1)
                    worksheet.write(start, 3, change_dict["field_requester"])
                    start += 1

                    for ind2, col_change_dict in enumerate(important_changes):
                        worksheet.write(
                            start + ind2 * 3, 0, "Column Name", self.formats["header"]
                        )
                        worksheet.write(
                            start + ind2 * 3 + 1, 0, "Old Value", self.formats["header"]
                        )
                        worksheet.write(
                            start + ind2 * 3 + 2, 0, "New Value", self.formats["header"]
                        )

                        worksheet.write(
                            start + ind2 * 3, 1, col_change_dict["col_name"]
                        )

                        if col_change_dict.get("old_val") is not None:
                            worksheet.write(
                                start + ind2 * 3 + 1, 1, col_change_dict["old_val"]
                            )
                            worksheet.write(
                                start + ind2 * 3 + 2, 1, col_change_dict["new_val"]
                            )
                        else:
                            for opt_ind, (old_option, new_option) in enumerate(
                                zip(
                                    col_change_dict["old_options"],
                                    col_change_dict["new_options"],
                                )
                            ):
                                worksheet.write(
                                    start + ind2 * 3 + 1,
                                    1 + opt_ind,
                                    old_option,
                                    {}
                                    if old_option in col_change_dict["new_options"]
                                    else self.formats["dropped"],
                                )
                                worksheet.write(
                                    start + ind2 * 3 + 2,
                                    1 + opt_ind,
                                    new_option,
                                    {}
                                    if new_option in col_change_dict["old_options"]
                                    else self.formats["new"],
                                )
                        start += 1
                    start += (len(important_changes) - 1) * 3 + 2
        start += 2

        # All Dropped Rows
        worksheet.write(start, 0, "All Dropped Rows:", self.formats["bold"])
        worksheet.write(start + 1, 0, "Variable / Field Name", self.formats["header"])
        worksheet.write(start + 1, 1, "Old Row Number", self.formats["header"])
        worksheet.write(start + 1, 2, "Diff Row Number", self.formats["header"])
        worksheet.write(start + 1, 3, "Field Requester", self.formats["header"])
        start += 2
        for ind, row in enumerate(self.dropped_rows):
            worksheet.write(start + ind, 0, row["field"])
            worksheet.write(start + ind, 1, row["old_row_num"] + 1)
            worksheet.write(start + ind, 2, row["diff_row_num"] + 1)
            if row.get("field_requester"):
                worksheet.write(start + ind, 3, row["field_requester"])
        start += len(self.dropped_rows) + 2

        # All Cell Changes
        worksheet.write(start, 0, "All Changes:", self.formats["bold"])
        worksheet.write(
            start, 1, json.dumps(self.important_change_rules), self.formats["bold"]
        )

        for ind, (field, change_dict) in enumerate(self.changes.items()):
            worksheet.write(
                start + 1, 0, "Variable / Field Name", self.formats["header"]
            )
            worksheet.write(start + 1, 1, "Old Row Number", self.formats["header"])
            worksheet.write(start + 1, 2, "New Row Number", self.formats["header"])
            worksheet.write(start + 1, 3, "Field Requester", self.formats["header"])
            start += 2
            worksheet.write(start, 0, field)
            worksheet.write(start, 1, change_dict["old_row_num"] + 1)
            worksheet.write(start, 2, change_dict["row_num"] + 1)
            if row.get("field_requester"):
                worksheet.write(start, 3, change_dict["field_requester"])
            start += 1
            for ind2, col_change_dict in enumerate(change_dict["changed_cols"]):
                worksheet.write(
                    start + ind2 * 3, 0, "Column Name", self.formats["header"]
                )
                worksheet.write(
                    start + ind2 * 3 + 1, 0, "Old Value", self.formats["header"]
                )
                worksheet.write(
                    start + ind2 * 3 + 2, 0, "New Value", self.formats["header"]
                )

                worksheet.write(start + ind2 * 3, 1, col_change_dict["col_name"])

                if col_change_dict.get("old_val") is not None:
                    worksheet.write(start + ind2 * 3 + 1, 1, col_change_dict["old_val"])
                    worksheet.write(start + ind2 * 3 + 2, 1, col_change_dict["new_val"])
                else:
                    for opt_ind, old_option in enumerate(
                        col_change_dict["old_options"]
                    ):
                        worksheet.write(
                            start + ind2 * 3 + 1,
                            1 + opt_ind,
                            old_option,
                            {}
                            if old_option in col_change_dict["new_options"]
                            else self.formats["dropped"],
                        )

                    for opt_ind, new_option in enumerate(
                        col_change_dict["new_options"]
                    ):
                        worksheet.write(
                            start + ind2 * 3 + 2,
                            1 + opt_ind,
                            new_option,
                            {}
                            if new_option in col_change_dict["old_options"]
                            else self.formats["new"],
                        )
                start += 1

            start += (len(change_dict["changed_cols"]) - 1) * 3 + 2

    def create_new_changes_sheet(self):

        df_merged = self.df_new.merge(
            self.df_old,
            left_on="Variable / Field Name",
            right_on="Variable / Field Name",
            suffixes=("_new", "_old"),
            how="outer",
            indicator=True,
        )
        df_merged.rename(
            columns={"Variable / Field Name": "VARIABLE"},
            inplace=True,
        )
        df_merged.rename(
            columns={
                col: f"MODIFIED_NEW_VALUE: {col.replace('_new','')}"
                if col.endswith("_new")
                else f"MODIFIED_OLD_VALUE: {col.replace('_old','')}"
                for col in df_merged.columns.to_list()[1:-1]
            },
            inplace=True,
        )
        df_merged["merged_form"] = (
            df_merged["MODIFIED_NEW_VALUE: Form Name"]
            .fillna("")
            .combine(
                df_merged["MODIFIED_OLD_VALUE: Form Name"].fillna(""),
                lambda s1, s2: s1 if s1 else s2,
            )
        )
        df_merged["merged_form"] = df_merged["merged_form"].astype("category")
        sort_order = pd.Series(
            self.df_new["Form Name"].unique().tolist()
            + self.df_old["Form Name"].unique().tolist()
        ).unique()
        df_merged.merged_form.cat.set_categories(sort_order, inplace=True)
        df_merged.index.name = "index"
        df_merged.sort_values(["merged_form", "index"], inplace=True)

        def compare_values(row, col):
            old = row[f"MODIFIED_OLD_VALUE: {col}"]
            new = row[f"MODIFIED_NEW_VALUE: {col}"]
            old_with_pipe_stripping = re.sub(r"\s*\|\s*", "|", f"{old}")
            new_with_pipe_stripping = re.sub(r"\s*\|\s*", "|", f"{new}")
            if row._merge != "both":
                return 0
            if (
                old == new
                or old_with_pipe_stripping == new_with_pipe_stripping
                or (pd.isnull(old) and pd.isnull(new))
            ):
                return 0
            return 1

        def set_change_type(row):
            if row._merge == "both" and 1 in [row[col] for col in mod_cols]:
                return "Modified"
            if row._merge == "left_only":
                return "New"
            if row._merge == "right_only":
                return "Removed"
            return "Unchanged"

        mod_cols = []
        for col in self.df_new.columns.to_list()[1:]:
            col_name = f"MODIFIED: {col}"
            mod_cols.append(col_name)
            df_merged[col_name] = df_merged.apply(
                lambda row: compare_values(row, col), axis=1
            )

            df_merged[f"MODIFIED_NEW_VALUE: {col}"] = df_merged.apply(
                lambda row: row[f"MODIFIED_NEW_VALUE: {col}"]
                if row[col_name] == 1
                else "N/A",
                axis=1,
            )
            df_merged[f"MODIFIED_OLD_VALUE: {col}"] = df_merged.apply(
                lambda row: row[f"MODIFIED_OLD_VALUE: {col}"]
                if row[col_name] == 1
                else "N/A",
                axis=1,
            )

        df_merged["CHANGE_TYPE"] = df_merged.apply(
            lambda row: set_change_type(row), axis=1
        )
        df_merged = df_merged[df_merged.CHANGE_TYPE != "Unchanged"]

        df_merged.reset_index(drop=True, inplace=True)

        final_fields = ["VARIABLE", "merged_form", "CHANGE_TYPE"]
        [
            final_fields.extend(
                [
                    f"MODIFIED: {col}",
                    f"MODIFIED_OLD_VALUE: {col}",
                    f"MODIFIED_NEW_VALUE: {col}",
                ]
            )
            for col in self.df_new.columns.to_list()[1:]
        ]
        df_merged = df_merged[final_fields]
        df_merged.rename(columns={"merged_form": "FORM_NAME"}, inplace=True)
        df_merged.to_excel(self.writer, sheet_name="NEW_CHANGE_NOTES", index=False)
        worksheet = self.writer.sheets["NEW_CHANGE_NOTES"]
        worksheet.set_column("A:BD", 30, self.formats["wrap"])


def main(args):
    if len(args) < 2:
        print(
            "Please enter 'python3 diff.py' followed by the filepath for the old file and the filepath for the new file, separated by spaces. EX: 'python3 diff.py old_file.xlsx new_file.xlsx'"
        )
        return
    PATH_OLD = Path(args[1])
    PATH_NEW = Path(args[2])
    filename = None
    if len(args) > 3:
        filename = args[3]
    diff_class = ExcelDiff(PATH_OLD, PATH_NEW, filename=filename)
    diff_class.diff(verbose=True)


if __name__ == "__main__":
    main(sys.argv)
