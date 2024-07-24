from operator import ge
import openpyxl
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.image as mpimg
import string
import numpy as np

from openpyxl.styles import colors
from openpyxl.styles import Font, Color

# https://stackoverflow.com/questions/50172905/center-a-label-inside-a-circle-with-matplotlib

DATA_PATH = "spreadsheet.xlsx"
IMG_PATH = "img.jpg"

IMG_WIDTH = 700
IMG_HEIGHT = 720

def create_radar_ids() -> list[str]:
    ids = []
    for i in range(100):
        if i >= 10:
            ids.append(str(i))
        else:
            ids.append("0" + str(i))

    for c in string.ascii_uppercase:
        for i in range(10):
            ids.append(c + str(i))
    
    return ids

def sort_project_status(df: pd.DataFrame) -> list[list[str]]:
    green = []
    amber = []
    hold = []
    red = []
    for index, row in df.iterrows():
        match row["Overall Health"]:
            case "GREEN":
                green.append(row["Project Name"])
            case "Amber":
                amber.append(row["Project Name"])
            case "ON HOLD":
                hold.append(row["Project Name"])
            case "RED": # check
                red.append(row["Project Name"])
    return [green, amber, hold, red]

def generate_excel(radar_ids: list[str], names: list[list[str]]) -> None:
    gkey = "Green Status: " + str(len(names[0])) + " projects"
    akey = "Amber Status: " + str(len(names[1])) + " projects"
    okey = "On Hold: " + str(len(names[2])) + " projects"
    rkey = "Red Status: " + str(len(names[3])) + " projects"

    data = dict()
    data[gkey] = []
    data[akey] = []
    data[okey] = []
    data[rkey] = []
    
    idx = 0
    for i in range(len(names)):
        for j in range(len(names[i])):
            match i:
                case 0:
                    data[gkey].append(names[i][j] + ": " + radar_ids[idx])
                case 1:
                    data[akey].append(names[i][j] + ": " + radar_ids[idx])
                case 2:
                    data[okey].append(names[i][j] + ": " + radar_ids[idx])
                case 3:
                    data[rkey].append(names[i][j] + ": " + radar_ids[idx])
            idx += 1

    max_len = max(map(len, data.values()))
    for v in data.values():
        for i in range(max_len-len(v)):
            v.append(np.nan)

    df = pd.DataFrame.from_dict(data)
    # print(df.head)

    def color(val: str) -> str:
        if val in df[gkey]:
            return "color: green"
        if val in df[akey]:
            return "color: orange"
        if val in df[okey]:
            return "color: gray"
        if val in df[rkey]:
            return "color: red"

    df.insert(1,"","")
    df.insert(3," ","")
    df.insert(5,"  ","")

    df.to_excel("test.xlsx", index=False)
    wb = openpyxl.load_workbook("test.xlsx")
    ws = wb.active
    for c in ["A", "C", "E", "G"]:
        for i in range(1, max_len+2):
            match c:
                case "A":
                    ws[c + str(i)].font = Font(color=colors.Color("00ff00"))
                case "C":
                    ws[c + str(i)].font = Font(color=colors.Color("e66419"))
                case "E":
                    ws[c + str(i)].font = Font(color=colors.Color("000000"))
                case "G":
                    ws[c + str(i)].font = Font(color=colors.Color("ff0000"))

    wb.save("test.xlsx")

def calculate_radius(percent: float) -> float:
    return (310 * (1 - percent))

def main() -> None:
    df = pd.read_excel(DATA_PATH)
    df = df.dropna(how="all", axis=1)

    df.columns = df.iloc[0].tolist()
    df = df[1:]
    df = df.reset_index()
    df = df.drop("index", axis=1)
    names = sort_project_status(df)
    ids = create_radar_ids()
    # generate_excel(create_radar_ids(), names)
    idx = 0
    for i in range(len(names)):
        for name in names[i]:
            row_num = 0
            for j in range(len(df["Project Name"])):
                if df["Project Name"][j] == name:
                    row_num = j
                    break
            r = calculate_radius(df["%Project Duration Completed2"][row_num])
            match df["Service Category"][row_num]:
                case "InfoSec Protection Services":
                    # 0 - pi/4
                    pass
                case "IT Risk Management":
                    # pi/4 - pi/2
                    pass
                case "Identity and Access":
                    # pi/2 - 3pi/4
                    pass
                case "Threat Management":
                    # 3pi/4 - pi
                    pass
                case "InfoSec Program Management":
                    # pi - 5pi/4
                    pass
                case "InfoSec Program Support":
                    # 5pi/4 - 3pi/2
                    pass
                case "Security Design Services":
                    # 3pi/2 - 7pi/4
                    pass
                case "Compliance & Assurance":
                    # 7pi/4 - 2pi
                    pass
                case _:
                    print("This should not run")
            idx += 1

    fig, ax = plt.subplots(frameon=False)

    ax.imshow(mpimg.imread(IMG_PATH), aspect="auto")

    circle = plt.Circle((660, IMG_HEIGHT//2), 10, color="green", fill=True)
    ax.add_artist(circle)
    ax.text(660, IMG_HEIGHT//2, "te", ha="center", va="center", fontsize=7, color="white")
    ax.set_aspect("equal", adjustable="box")
    ax.set_axis_off()

    # Save the figure without borders
    fig.savefig("radar.png", dpi=250, bbox_inches="tight")

    plt.show()
    
    

if __name__ == "__main__":
    main()
