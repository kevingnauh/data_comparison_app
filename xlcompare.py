import pandas as pd

class ExcelComparator:
    def __init__(self, file_path_1, file_path_2, file_sheet_1, file_sheet_2, unique_id_cols):
        self.df_day1 = pd.read_excel(file_path_1, sheet_name=file_sheet_1)
        self.df_day2 = pd.read_excel(file_path_2, sheet_name=file_sheet_2)
        self.file_sheet_1 = file_sheet_1
        self.file_sheet_2 = file_sheet_2
        self.unique_id_cols = unique_id_cols

        if ',' in self.unique_id_cols:
            self.unique_id_cols = [item.strip() for item in self.unique_id_cols.split(',')]

            self.df_day1 = self._create_unique_id(self.df_day1, self.unique_id_cols)

            self.df_day2 = self._create_unique_id(self.df_day2, self.unique_id_cols)

            self.unique_id_cols = 'unique_id'

    def _create_unique_id(self, df, column_names):
        # Combines columns to create 'unique_id' column
        df['unique_id'] = df[self.unique_id_cols].astype(str).agg('_'.join, axis=1)
        # df.drop(columns=self.unique_id_cols, inplace=True)
        # df.rename(columns={'unique_id': column_names}, inplace=True)

        return df

    def _compare_data(self, df1, df2, unique_id_cols, result_msg):
        """Compares two dataframe by the unique_id and creates a column indicating results"""
        df1['Result'] = 'no_change'

        for index, row in df1.iterrows():
            unique_id = row[unique_id_cols]
            if unique_id not in df2[unique_id_cols].values:
                df1.at[index, 'Result'] = result_msg
            else:
                changes = self._get_changes(row, df2[df2[unique_id_cols] == unique_id].iloc[0])
                if changes:
                    df1.at[index, 'Result'] = f'Values changed: {changes}'
        return df1

    def _get_changes(self, row1, row2):
        changes = []
        for col in self.df_day1.columns:
            if col != 'Result' and row1[col] != row2[col]:
                changes.append(col)
        return ', '.join(changes)

    def _combine_data(self):
        self.df_all = pd.concat([self.df_day1.assign(Sheet=self.file_sheet_1), self.df_day2.assign(Sheet=self.file_sheet_2)])

    def _save_to_excel(self):
        with pd.ExcelWriter('output.xlsx') as writer:
            self.df_day1.to_excel(writer, sheet_name=self.file_sheet_1, index=False)
            self.df_day2.to_excel(writer, sheet_name=self.file_sheet_2, index=False)
            self.df_all.to_excel(writer, sheet_name='All', index=False)

    def main(self):
        self.df_day1 = self._compare_data(self.df_day1, self.df_day2, self.unique_id_cols, 'removed')
        self.df_day2 = self._compare_data(self.df_day2, self.df_day1, self.unique_id_cols, 'added')

        self._combine_data()
        self._save_to_excel()

if __name__ == "__main__":
    file_path_1 = 'daily_record_changes.xlsx'
    file_path_2 = 'daily_record_changes.xlsx'
    file_sheet_1 = 'day_1'
    file_sheet_2 = 'day_2'
    unique_id = 'unique_id'
    # unique_id = 'last_name,first_name,county,district,school,primary_job, fte'

    comparator = ExcelComparator(file_path_1, file_path_2, file_sheet_1, file_sheet_2, unique_id)

    comparator.main()
