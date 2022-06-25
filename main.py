import pandas as pd


def df_maker(path: str):
    """ Makes a df out of .xlsx file
    :param path: full path of file
    :return: sorted DF, DF with alternatives' names
    """
    df = pd.read_excel(path, sheet_name=0)
    significance = dict(zip(list(df.iloc[-1])[1::], list(df.columns)[1::]))
    df.drop(labels=df.tail(1).index, axis=0, inplace=True)
    df_alternatives = df['Альтернативы']
    df.pop('Альтернативы')

    columns = [significance[i] for i in range(1, len(significance) + 1)][::-1]
    df = df.reindex(columns=columns)

    return df, df_alternatives


def xlsx_writer(df, path: str) -> None:
    """
    Writes df into a new Excel-sheet "Result"

    :param df: df
    :param path: path of the file
    :return: None
    """
    with pd.ExcelWriter(path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        df.to_excel(writer, sheet_name='Result')

    print('Done')

    return None


def __is_max(df, col_index: int):
    """Compares values of df's column with it is max value.

    :param df: df
    :param col_index: index of a column
    :return: 1 if is max value, 0 if not"""
    col = df.columns[col_index]
    return (df[col] == df[col].max()) * 1


def lexographic_order(df) -> list:
    """
    Uses lexographic_best() function to make an order of alternatives by its significance

    :param df: df
    :return: list of indexes in order of it's significance
    """
    new_df = df.copy()
    res = []
    while not new_df.empty:
        labels = lexographic_best(new_df)
        res.append(*labels)
        new_df.drop(labels=labels, axis=0, inplace=True)

    return res


def lexographic_best(df) -> list:
    """
    Uses lexographic method to find out the best value(s)

    :param df: df
    :return: list of indexes of best left alternatives
    """
    new_df = df.copy()
    for col in range(new_df.shape[1]):
        labels = new_df.index[__is_max(new_df, col) != 1].tolist()
        new_df.drop(labels=labels, axis=0, inplace=True)

        if new_df.shape[0] <= 1:
            break

    return new_df.index.to_list()


def pareto_set(df):
    res = []
    for i, s in zip(range(df.shape[1]), range(df.shape[1], 0, -1)):
        v = __is_max(df, i) * s / 100
        res.append(v)

    res = sum(res)
    new_df = res.loc[res <= res.mean()]
    return new_df.index.to_list()


def main():
    pd.set_option('display.max_rows', None)
    pd.set_option('display.max_columns', None)
    pd.set_option('display.max_colwidth', None)

    path = 'Videocards.xlsx'

    # Building DF
    df, names = df_maker(path)
    names.index += 1

    # DF after pareto-set-selection
    df = df.drop(pareto_set(df))

    # Lexographic method
    df_res = names.reindex(lexographic_order(df))
    xlsx_writer(df_res, path)


if __name__ == '__main__':
    main()
