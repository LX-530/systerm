#!/usr/bin/env python3
"""
价格与毛利结构分析脚本

功能:
1. 读取 `程宇昕.xlsx` (或指定Excel) 中的 `销售数据` 工作表。
2. 清洗数据(填充一级分类、剔除分类汇总行、转数值)。
3. 生成品类/单品层级的销售额与毛利汇总表。
4. 输出关键洞察CSV，并绘制用于价格走向与毛利对比的图形。

用法:
    python3 price_analysis.py --input 程宇昕.xlsx --sheet 销售数据 --out outputs
"""

from __future__ import annotations

import argparse
from pathlib import Path
from typing import Tuple

import matplotlib.pyplot as plt
import pandas as pd
import seaborn as sns


NUMERIC_COLS = ["求和项:实际金额", "求和项:销售毛利", "求和项:标品毛利率", "求和项:贡献值"]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="商品价格与毛利分析")
    parser.add_argument("--input", default="程宇昕.xlsx", help="Excel文件路径")
    parser.add_argument("--sheet", default="销售数据", help="读取的工作表名称")
    parser.add_argument("--out", default="outputs", help="输出目录")
    parser.add_argument("--top-n", type=int, default=20, help="价格走势图中展示的Top N商品")
    return parser.parse_args()


def load_data(path: Path, sheet_name: str) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name=sheet_name)
    df["一级分类"] = df["一级分类"].ffill()
    df = df[df["商品名称"].notna()].copy()
    df = df[~df["一级分类"].astype(str).str.contains("汇总", na=False)].copy()
    for col in NUMERIC_COLS:
        df[col] = pd.to_numeric(df[col], errors="coerce")
    return df


def summarize(df: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame]:
    cat_summary = (
        df.groupby("一级分类")
        .agg(
            sales_amount=("求和项:实际金额", "sum"),
            gross_profit=("求和项:销售毛利", "sum"),
            avg_margin_rate=("求和项:标品毛利率", "mean"),
            median_margin_rate=("求和项:标品毛利率", "median"),
            sku_count=("商品名称", "nunique"),
        )
        .reset_index()
    )
    cat_summary["realized_margin_rate"] = cat_summary["gross_profit"] / cat_summary[
        "sales_amount"
    ]

    product_summary = (
        df.groupby(["一级分类", "商品名称"])
        .agg(
            sales_amount=("求和项:实际金额", "sum"),
            gross_profit=("求和项:销售毛利", "sum"),
            avg_margin_rate=("求和项:标品毛利率", "mean"),
        )
        .reset_index()
    )
    product_summary["realized_margin_rate"] = product_summary["gross_profit"] / product_summary[
        "sales_amount"
    ]
    product_summary.replace([pd.NA, pd.NaT, float("inf"), float("-inf")], 0, inplace=True)

    return cat_summary, product_summary


def plot_category_margin(cat_summary: pd.DataFrame, out_dir: Path) -> None:
    sorted_cat = cat_summary.sort_values("sales_amount", ascending=False)
    plt.figure(figsize=(10, 6))
    sns.barplot(
        data=sorted_cat,
        x="sales_amount",
        y="一级分类",
        color="#4c72b0",
        label="销售额",
    )
    ax = plt.gca()
    ax.set_xlabel("销售额")
    ax.set_ylabel("一级分类")

    ax2 = ax.twiny()
    ax2.plot(
        sorted_cat["realized_margin_rate"] * 100,
        sorted_cat["一级分类"],
        color="#dd8452",
        marker="o",
        label="毛利率(%)",
    )
    ax2.set_xlabel("毛利率(%)")
    ax.set_title("一级分类销售额 vs. 实际毛利率")
    ax.legend(loc="lower right")
    plt.tight_layout()
    plt.savefig(out_dir / "category_margin.png", dpi=200)
    plt.close()


def plot_price_trend(product_summary: pd.DataFrame, out_dir: Path, top_n: int) -> None:
    top_products = (
        product_summary.sort_values("sales_amount", ascending=False)
        .head(top_n)
        .sort_values("sales_amount")
    )
    plt.figure(figsize=(12, 6))
    sns.lineplot(
        data=top_products,
        x="商品名称",
        y="sales_amount",
        marker="o",
        label="销售额(价格指标)",
        color="#1b9e77",
    )
    ax = plt.gca()
    ax.set_xticklabels(top_products["商品名称"], rotation=60, ha="right")
    ax.set_xlabel("Top商品 (按销售额排序)")
    ax.set_ylabel("销售额")

    ax2 = ax.twinx()
    sns.lineplot(
        data=top_products,
        x="商品名称",
        y=top_products["realized_margin_rate"] * 100,
        marker="s",
        color="#d95f02",
        label="毛利率(%)",
        ax=ax2,
    )
    ax2.set_ylabel("毛利率(%)")
    ax.set_title(f"Top {top_n} 商品价格(销售额)走势与毛利率对比")
    ax.grid(axis="y", linestyle="--", alpha=0.3)
    plt.tight_layout()

    # 解决双轴图例
    handles, labels = [], []
    for axis in [ax, ax2]:
        h, l = axis.get_legend_handles_labels()
        handles.extend(h)
        labels.extend(l)
    ax.legend(handles, labels, loc="upper left")

    plt.savefig(out_dir / "price_trend.png", dpi=200)
    plt.close()


def plot_margin_distribution(df: pd.DataFrame, out_dir: Path) -> None:
    plt.figure(figsize=(8, 5))
    sns.histplot(
        df["求和项:标品毛利率"],
        bins=40,
        color="#2c7fb8",
        kde=True,
        stat="density",
    )
    plt.title("单品毛利率分布")
    plt.xlabel("毛利率")
    plt.ylabel("密度")
    plt.tight_layout()
    plt.savefig(out_dir / "margin_distribution.png", dpi=200)
    plt.close()


def export_tables(cat_summary: pd.DataFrame, product_summary: pd.DataFrame, out_dir: Path) -> None:
    cat_summary.to_csv(out_dir / "category_summary.csv", index=False)
    product_summary.to_csv(out_dir / "product_summary.csv", index=False)

    high_margin = product_summary[
        product_summary["realized_margin_rate"] >= 0.18
    ].sort_values("realized_margin_rate", ascending=False)
    low_margin = product_summary[
        product_summary["realized_margin_rate"] <= 0.05
    ].sort_values("sales_amount", ascending=False)
    zero_margin = product_summary[
        product_summary["realized_margin_rate"] <= 0
    ].sort_values("sales_amount", ascending=False)

    high_margin.to_csv(out_dir / "high_margin_products.csv", index=False)
    low_margin.to_csv(out_dir / "low_margin_products.csv", index=False)
    zero_margin.to_csv(out_dir / "zero_margin_products.csv", index=False)


def main() -> None:
    args = parse_args()
    input_path = Path(args.input)
    out_dir = Path(args.out)
    out_dir.mkdir(parents=True, exist_ok=True)

    df = load_data(input_path, args.sheet)
    cat_summary, product_summary = summarize(df)

    export_tables(cat_summary, product_summary, out_dir)
    plot_category_margin(cat_summary, out_dir)
    plot_price_trend(product_summary, out_dir, args.top_n)
    plot_margin_distribution(df, out_dir)

    print("分析完成。输出目录:", out_dir.resolve())


if __name__ == "__main__":
    sns.set_theme(style="whitegrid", palette="Set2")
    main()
