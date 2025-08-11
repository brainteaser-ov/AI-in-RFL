import os
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt


def collect_tfidf(folder: str, sheet: str,
                  col_names=("term", "value")) -> pd.DataFrame:
    combined = None
    for fname in sorted(os.listdir(folder)):
        if not (fname.endswith(".xlsx") and fname[:4].isdigit()):
            continue
        year = fname[:4]
        path = os.path.join(folder, fname)

        try:
            df = pd.read_excel(path, sheet_name=sheet, names=col_names)
        except ValueError:                       # если лист назван иначе
            idx = 0 if sheet == "keywords_tfidf" else 1
            df = pd.read_excel(path, sheet_name=idx, names=col_names)

        df_year = df.rename(columns={col_names[1]: year})
        combined = df_year if combined is None else \
                   combined.merge(df_year, on=col_names[0], how="outer")

    if combined is None:
        raise FileNotFoundError("В папке нет подходящих XLSX-файлов.")

    return combined.set_index(col_names[0])


def plot_heatmap(df: pd.DataFrame) -> None:
    df_sorted = (df.assign(max_val=df.max(axis=1, skipna=True))
                   .sort_values("max_val", ascending=False)
                   .drop(columns="max_val"))

    data = df_sorted.fillna(0)
    fig, ax = plt.subplots(figsize=(10, max(4, len(data) * 0.4)))
    im = ax.imshow(data.values, aspect="auto", cmap="Pastel1")

    # Цветовая шкала с увеличенным шрифтом
    cbar = fig.colorbar(im, ax=ax)
    cbar.set_label("TF-IDF", fontsize=12)  # Увеличиваем шрифт подписи
    cbar.ax.tick_params(labelsize=10)      # Увеличиваем шрифт значений на шкале

    # Настройка осей с увеличенным шрифтом
    ax.set_xticks(np.arange(len(data.columns)))
    ax.set_xticklabels(data.columns, fontsize=12)  # Увеличили шрифт годов
    ax.set_yticks(np.arange(len(data.index)))
    ax.set_yticklabels(data.index, fontsize=12)    # Увеличили шрифт терминов

    # Добавление значений в ячейки с увеличенным шрифтом
    for i in range(len(data.index)):
        for j in range(len(data.columns)):
            val = df_sorted.iloc[i, j]
            txt = "–" if pd.isna(val) else f"{val:.4f}"
            ax.text(j, i, txt,
                    ha="center",
                    va="center",
                    fontsize=10)  # Увеличили размер текста в ячейках

    fig.tight_layout()
    plt.show()



if __name__ == "__main__":
    folder = '/укажите путь к своей папке с файлами'

    # KEYWORDS
    kw_df = collect_tfidf(folder, "keywords_tfidf")
    plot_heatmap(kw_df)

    # ANNOTATIONS
    ann_df = collect_tfidf(folder, "annotation_tfidf")
    plot_heatmap(ann_df)
