# Hayden Coffey
import pandas as pd
import seaborn as sns
import sys
import matplotlib.pyplot as plt

# Returns a df with the delta between both experiment means.
def get_run_delta(df1, df2):
    merged = pd.merge(
            df1,
            df2,
            on=["Target", "Sample"],
            how="outer",
            suffixes=("_run1", "_run2")
            )

    merged = merged.dropna()
    merged["Day"] = merged['Sample'].str.extract(r'([DF]\d+)')
    merged["Replicate"] = merged['Sample'].str.extract(r'\(([Pp]?n=\d+)\)')
    merged['Cq_Delta'] = abs(merged['Cq_mean_run1'] - merged['Cq_mean_run2'])

    return merged

# Create boxplot figure for given Gene/Radiation pairing.
def create_target_figure(inDF, target, rad):
    df = (inDF[inDF['Target'] == target])
    df = df[df['Sample'].str.contains(rad)]

    unique_days = df['Day'].unique()
    day_order = sorted(
            unique_days,
            key=lambda x: (
                0 if x[0] == 'F' else 1,
                int(x[1:])
                )
            )

    plt.figure(figsize=(24,12))
    sns.boxplot(data=df, x='Day', y='Cq_mean', fill=False, order=day_order)
    sns.scatterplot(data=df, x='Day', y='Cq_mean', hue='Replicate', style='Replicate', s=150)

    # If we want text annotations
    #df['Day'] = pd.Categorical(df['Day'], categories=day_order, ordered=True)
    #ax = plt.gca()
    #for i, row in df.iterrows():
    #    # get the x-coordinate of this category
    #    x = day_order.index(row['Day'])
    #    y = row['Cq_mean']
    #    ax.text(x, y + 0.1,      # tweak the +0.1 so it sits above the diamond
    #            f"{y:.2f}",     # formatted median value
    #            ha='center', va='bottom', fontsize=12)

    plt.title("{}: {}".format(target, rad))
    plt.savefig("./2025_04_27/{}_{}.png".format(target,rad), dpi=300, bbox_inches="tight")
    plt.close()

def preprocess_files(file1, file2):
    SelectColumns = ['Target', 'Sample', 'Cq']

    df1 = pd.read_excel(file1)
    df1 = df1[SelectColumns]

    df2 = pd.read_excel(file2)
    df2 = df2[SelectColumns]

    df1 = df1.dropna()
    df2 = df2.dropna()

    # Remove some entries with Undetermined cq
    df1 = df1[df1["Cq"] != "Undetermined"]
    df2 = df2[df2["Cq"] != "Undetermined"]

    # Typecast to float so we can average
    df1["Cq"] = df1["Cq"].astype(float)
    df2["Cq"] = df2["Cq"].astype(float)

    df1['Cq_mean'] = df1.groupby(['Sample', 'Target'])['Cq'].transform('mean')
    df2['Cq_mean'] = df2.groupby(['Sample', 'Target'])['Cq'].transform('mean')

    return df1, df2

def main():
    if len(sys.argv) < 3:
        print("Usage: python {} file1.xlsx file2.xlsx".format(sys.argv[0]))
        return

    sns.set(font_scale=2)
    file1 = sys.argv[1]
    file2 = sys.argv[2]

    df1, df2 = preprocess_files(file1, file2)

    print("Creating delta xlsx file...")
    delta_df = (get_run_delta(df1, df2))
    delta_df.to_excel("./2025_04_27/delta_output.xlsx")
    print("Done.")

    df = pd.concat([df1,df2], ignore_index=True).reset_index()

    # Capture both D# and F# entries
    df["Day"] = df['Sample'].str.extract(r'([DF]\d+)')
    df["Replicate"] = df['Sample'].str.extract(r'\(([Pp]?n=\d+)\)')
    df["Rad"] = df['Sample'].str.extract(r'(.*?Gy)')

    # Record median data to new xlsx file
    df['Cq_median_by_day'] = df.groupby(['Day', 'Target', 'Rad'])['Cq'].transform('median')
    df.to_excel("./2025_04_27/cq_median.xlsx")

    print("Generating figures...")
    for rad in (df['Rad'].unique()):
        for target in (df['Target'].unique()):
            print("Plotting [{}, {}]...".format(rad, target))
            create_target_figure(df, target, rad)
    print("Done.")

    return 0

if __name__ == "__main__":
    main()
