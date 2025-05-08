import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt

# Mapping
DayMap = {
        "D1": "D1",
        "D3": "D3",
        "D5": "D5",
        "D7": "D3",
        "D10": "D5"
        }

FracDayMap = {
        "F1": "D1",
        "F5": "D5",
        "D1": "D5",
        "D3": "D3",
        "D5": "D5",
        "D7": "D3",
        "D10": "D5"
        }

def extract_sample_fields(df, SampleColumn):
    df["Day"] = df[SampleColumn].str.extract(r'([DF]\d+)')
    df["Replicate"] = df[SampleColumn].str.extract(r'\(([Pp]?n=\d+)\)')
    df["Rad"] = df[SampleColumn].str.extract(r'(.*?Gy)')

    # Assume missing Rad are 5x2Gy
    df["Rad"] = df["Rad"].fillna("5x2Gy")

    return df

def main():
    # Content is Gy and Day info
    # Target is gene
    SampleColumn = 'Content'
    Columns = ['Well', 'Target', 'Content', 'Cq']
    df = pd.read_excel("2025_04_25/result2.xlsx")
    df = df[Columns]

    # Pull out Day, Rad, and Replicate fields
    df = extract_sample_fields(df, SampleColumn)

    # Clean data
    df = df.dropna()
    df = df[df['Cq'] != 'Undetermined']

    # Calculate mean Cq
    df['Cq'] = df['Cq'].astype(float)
    df = df.groupby(['Target', 'Day', 'Rad'], as_index=False)['Cq'].mean()
    df = df.rename(columns={'Cq': 'Avg_Cq'})

    # Calculate dCq
    dCqs = []
    for index, row in df.iterrows():
        dCq = 0 
        if row['Rad'] == '5x2Gy':
            baseline = df[(df['Target'] == 'GUSB') & (df['Rad'] == row['Rad']) \
                    & (df['Day'] == row['Day'])]['Avg_Cq'].values

        else:
            baseline = df[(df['Target'] == 'GUSB') & (df['Rad'] == row['Rad']) \
                    & (df['Day'] == row['Day'])]['Avg_Cq'].values

        assert len(baseline) == 1
        dCq = row['Avg_Cq'] - baseline[0]

        dCqs.append(dCq)

    df['dCq'] = dCqs

    # Calculate ddCq
    ddCqs = []
    folds = []
    for index, row in df.iterrows():
        ddCq = 0
        fold = 0
        if row['Target'] != 'GUSB' and row['Rad'] != '0 Gy':
            if row['Rad'] == '5x2Gy':
                baseline = df[(df['Target'] == row['Target']) & (df['Rad'] == '0 Gy') \
                        & (df['Day'] == FracDayMap[row['Day']])]['dCq'].values

            else:
                baseline = df[(df['Target'] == row['Target']) & (df['Rad'] == '0 Gy') \
                        & (df['Day'] == DayMap[row['Day']])]['dCq'].values

            assert len(baseline) == 1
            ddCq = row['dCq'] - baseline[0]
            fold = 2**(-ddCq)

        ddCqs.append(ddCq)
        folds.append(fold)

    df['ddCq'] = ddCqs
    df['Fold'] = folds 

    with pd.option_context('display.max_rows', None, 'display.max_columns', None):  # more options can be specified also
        print(df)

    unique_days = df['Day'].unique()
    day_order = sorted(
            unique_days,
            key=lambda x: (
                0 if x[0] == 'F' else 1,
                int(x[1:])
                )
            )
    df['Day'] = pd.Categorical(df['Day'], categories=day_order, ordered=True)

    sns.set(font_scale=2)
    for target in df['Target'].unique():
        plt.title(target)
        sns.lineplot(data=df[(df['Target'] == target) & (df['Rad'] != '0 Gy')], x='Day', y='Fold', hue='Rad')
        plt.show()

if __name__ == "__main__":
    main()
