import pandas as pd
from datetime import date
import matplotlib.pyplot as plt
import matplotlib.image as mpimg

# https://stackoverflow.com/questions/50172905/center-a-label-inside-a-circle-with-matplotlib

DATA_PATH = "spreadsheet.xlsx"
IMG_PATH = "img.jpg"

IMG_WIDTH = 1280
IMG_HEIGHT = 720

def parse_date(d: str) -> date:
    splt = map(int, d.split("-"))
    return date(splt[0], splt[1], splt[2])

def main() -> None:
    df = pd.read_excel(DATA_PATH)
    df = df.dropna(how="all", axis=1)

    df.columns = df.iloc[0].tolist()
    df = df[1:]
    df = df.reset_index()
    df = df.drop("index", axis=1)

    plt.imshow(mpimg.imread(IMG_PATH))
    plt.show()

if __name__ == "__main__":
    main()
