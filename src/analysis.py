from pathlib import Path
import argparse

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import pandas as pd
import seaborn as sns

ROOT = Path(__file__).resolve().parents[1]
DEFAULT_INPUT = ROOT / "data" / "raw" / "Pokemon.csv"
DEFAULT_REPORTS = ROOT / "reports"

RENAMES = {
    "#": "id",
    "Name": "name",
    "Type 1": "type_1",
    "Type 2": "type_2",
    "Total": "total",
    "HP": "hp",
    "Attack": "attack",
    "Defense": "defense",
    "Sp. Atk": "sp_atk",
    "Sp. Def": "sp_def",
    "Speed": "speed",
    "Generation": "generation",
    "Legendary": "legendary",
}
STAT_COLS = ["hp", "attack", "defense", "sp_atk", "sp_def", "speed"]


def load(path: Path) -> pd.DataFrame:
    df = pd.read_csv(path)
    df = df.rename(columns=RENAMES)
    df["type_2"] = df["type_2"].fillna("None")
    df["legendary"] = df["legendary"].astype(bool)
    return df


def save_tables(df: pd.DataFrame, out_dir: Path) -> dict[str, pd.DataFrame]:
    out_dir.mkdir(parents=True, exist_ok=True)

    summary = df[STAT_COLS + ["total"]].describe().T.round(2)
    top_total = df.nlargest(10, "total")[["name", "type_1", "type_2", "total"]]
    top_attack = df.nlargest(10, "attack")[["name", "attack"]]
    top_speed = df.nlargest(10, "speed")[["name", "speed"]]

    type_stats = (
        df.melt(id_vars=["total"], value_vars=["type_1", "type_2"], value_name="type")
        .query("type != 'None'")
        .groupby("type")["total"]
        .agg(["count", "mean", "median", "max"])
        .sort_values("mean", ascending=False)
        .round(1)
    )
    legendary = (
        df.groupby("legendary")[STAT_COLS + ["total"]]
        .mean()
        .round(1)
        .rename(index={False: "regular", True: "legendary"})
    )
    correlations = (
        df[STAT_COLS + ["total"]].corr()["total"].drop("total").sort_values(ascending=False).round(3)
    )

    tables = {
        "summary": summary,
        "top_total": top_total,
        "top_attack": top_attack,
        "top_speed": top_speed,
        "type_stats": type_stats,
        "legendary": legendary,
        "correlations": correlations,
    }
    for name, table in tables.items():
        table.to_csv(out_dir / f"{name}.csv")
    return tables


def make_figures(df: pd.DataFrame, out_dir: Path) -> None:
    out_dir.mkdir(parents=True, exist_ok=True)
    sns.set_theme(style="whitegrid")

    fig, ax = plt.subplots(figsize=(8, 4))
    sns.histplot(df["total"], bins=30, kde=True, ax=ax, color="steelblue")
    ax.set_title("Distribution of total stats")
    fig.tight_layout()
    fig.savefig(out_dir / "total_distribution.png", dpi=150)
    plt.close(fig)

    top_types = df["type_1"].value_counts().head(10).index
    fig, ax = plt.subplots(figsize=(10, 5))
    sns.boxplot(data=df[df["type_1"].isin(top_types)], x="type_1", y="total", order=top_types, ax=ax)
    ax.set_title("Total stats by primary type (top 10 most common)")
    ax.tick_params(axis="x", rotation=45)
    fig.tight_layout()
    fig.savefig(out_dir / "total_by_type.png", dpi=150)
    plt.close(fig)

    fig, ax = plt.subplots(figsize=(7, 4))
    means = df.groupby("legendary")[STAT_COLS].mean().T
    means.columns = ["Regular" if not c else "Legendary" for c in means.columns]
    means.plot(kind="bar", ax=ax, color=["#4c72b0", "#dd8452"])
    ax.set_title("Average stats: regular vs legendary")
    ax.set_ylabel("Average value")
    ax.tick_params(axis="x", rotation=0)
    fig.tight_layout()
    fig.savefig(out_dir / "legendary_comparison.png", dpi=150)
    plt.close(fig)

    fig, ax = plt.subplots(figsize=(8, 6))
    sns.heatmap(df[STAT_COLS + ["total"]].corr(), annot=True, fmt=".2f", cmap="coolwarm", ax=ax)
    ax.set_title("Correlation between stats")
    fig.tight_layout()
    fig.savefig(out_dir / "correlation_heatmap.png", dpi=150)
    plt.close(fig)

    fig, ax = plt.subplots(figsize=(8, 6))
    sns.scatterplot(
        data=df,
        x="attack",
        y="defense",
        hue="legendary",
        style="legendary",
        alpha=0.7,
        ax=ax,
    )
    ax.set_title("Attack vs defense (legendary highlighted)")
    fig.tight_layout()
    fig.savefig(out_dir / "attack_vs_defense.png", dpi=150)
    plt.close(fig)


def main() -> None:
    parser = argparse.ArgumentParser(description="Explore the Pokemon dataset.")
    parser.add_argument("--input", type=Path, default=DEFAULT_INPUT)
    parser.add_argument("--reports", type=Path, default=DEFAULT_REPORTS)
    args = parser.parse_args()

    df = load(args.input)
    tables = save_tables(df, args.reports / "tables")
    make_figures(df, args.reports / "figures")

    print(f"Rows: {len(df)}, columns: {df.shape[1]}")
    print(f"Missing values: {int(df.isna().sum().sum())}")
    print(f"Legendary Pokemon: {int(df['legendary'].sum())} of {len(df)}")
    print("\nTop 5 by total stats:")
    print(tables["top_total"].head(5).to_string(index=False))
    print("\nStrongest types by mean total:")
    print(tables["type_stats"].head(5).to_string())
    print("\nAverage stats by legendary status:")
    print(tables["legendary"].to_string())
    print(f"\nTables written to {args.reports / 'tables'}")
    print(f"Figures written to {args.reports / 'figures'}")


if __name__ == "__main__":
    main()
