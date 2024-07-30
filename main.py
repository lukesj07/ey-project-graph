from operator import ge
from numpy.lib.function_base import angle
import openpyxl
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.image as mpimg
import string
import numpy as np
import math

from openpyxl.styles import colors
from openpyxl.styles import Font, Color
from pandas.core.common import iterable_not_string

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
            case "RED":
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


def calculate_position(percent: float, positions: list[list[list]], idx: int, angle_bounds: list[float], name: str) -> list[float]:
    i = 0
    while True:
        overlapping = []
        if 0.75 <= percent <= 1.0:
            out = 0 + ((1 - 0) / (1 - 0.75)) * (percent - 0.75)
            ideal_radius = round(150 * (1 - out) + 15)

        elif percent < 0.75:
            out = 0 + ((1 - 0) / (0.75 - 0)) * (percent - 0)
            ideal_radius = round(160 * (1 - out) + 150)
    
    
        max_points = (((angle_bounds[1] - angle_bounds[0]) * ideal_radius) // 12) - 1

        for p in positions[idx]:
            if abs(p[1] - ideal_radius) < 16 and (angle_bounds[0] - 0.1) <= p[0] <= (angle_bounds[1] + 0.1):
                overlapping.append(p)
    
        if len(overlapping) == 0:
            print(f"No overlapping - name: {name} - r: {ideal_radius} - angle: {angle_bounds[0] + (10/(ideal_radius * 2))} - iters: {i}")
            return [ideal_radius, angle_bounds[0] + (10/(ideal_radius * 2))]

        if len(overlapping) >= max_points:
            percent -= 0.01
        else:
            max_a = max([p[0] for p in overlapping])
            if max_a + (20/(ideal_radius*2)) > angle_bounds[1]:
                percent -= 0.01
            else:
                print(f"Take max current angle - name: {name} - r: {ideal_radius} - angle: {max_a + (30/(ideal_radius * 2))} - iters: {i}")
                return [ideal_radius, max_a + (30/(ideal_radius*2))]

        i += 1


def main() -> None:
    df = pd.read_excel(DATA_PATH)
    df = df.dropna(how="all", axis=1)

    df.columns = df.iloc[0].tolist()
    df = df[1:]
    df = df.reset_index()
    df = df.drop("index", axis=1)
    names = sort_project_status(df)
    ids = create_radar_ids()
    generate_excel(create_radar_ids(), names)
    positions = [[], [], [], [], [], [], [], [], []] # list[list[list[float, float, str, str]]] theta, r, id, health
    
    id_idx = 0
    for i in range(len(names)):
        for name in names[i]:
            row_num = 0
            for idx, row in df.iterrows():
                if any(name in str(value) for value in row):
                    row_num = idx
                    break

            percent = df["%Project Duration Completed2"][row_num]

            match df["Service Category"][row_num]:
                case "InfoSec Protection Services":
                    # 0 - pi/4
                    r, a = calculate_position(percent, positions, 0, [0.01, math.pi/4], name)
                    positions[0].append([a, r, ids[id_idx], df["Overall Health"][row_num]])
                case "IT Risk Management":
                    # pi/4 - pi/2
                    r, a = calculate_position(percent, positions, 1, [math.pi/4+0.05, math.pi/2], name)
                    positions[1].append([a, r, ids[id_idx], df["Overall Health"][row_num]])
                case "Identity and Access":
                    # pi/2 - 3pi/4
                    r, a = calculate_position(percent, positions, 2, [math.pi/2, 3*math.pi/4], name)
                    positions[2].append([a, r, ids[id_idx], df["Overall Health"][row_num]])
                case "Threat Management":
                    # 3pi/4 - pi
                    r, a = calculate_position(percent, positions, 3, [3*math.pi/4+0.05, math.pi], name)
                    positions[3].append([a, r, ids[id_idx], df["Overall Health"][row_num]])
                case "InfoSec Program Management":
                    # pi - 5pi/4
                    r, a = calculate_position(percent, positions, 4, [math.pi + 0.01, 5*math.pi/4], name)
                    positions[4].append([a, r, ids[id_idx], df["Overall Health"][row_num]])
                case "InfoSec Program Support":
                    # 5pi/4 - 3pi/2
                    r, a = calculate_position(percent, positions, 5, [5*math.pi/4+0.05, 3*math.pi/2], name)
                    positions[5].append([a, r, ids[id_idx], df["Overall Health"][row_num]])
                case "Security Design Services":
                    # 3pi/2 - 7pi/4
                    r, a = calculate_position(percent, positions, 6, [3*math.pi/2, 7*math.pi/4], name)
                    positions[6].append([a, r, ids[id_idx], df["Overall Health"][row_num]])
                case "Compliance & Assurance":
                    # 7pi/4 - 2pi
                    r, a = calculate_position(percent, positions, 7, [7*math.pi/4+0.07, 2*math.pi], name)
                    positions[7].append([a, r, ids[id_idx], df["Overall Health"][row_num]])
                case _:
                    print("This should not run")

            id_idx += 1
    

    fig, ax = plt.subplots(frameon=False)

    ax.set_aspect("equal", adjustable="box")
    ax.set_axis_off()
    ax.imshow(mpimg.imread(IMG_PATH), aspect="auto")
    # theta, r, id, health
    for c in positions:
        for p in c:
            curr_color = ""
            match p[3]:
                case "GREEN":
                    curr_color = "#00ff00"
                case "Amber":
                    curr_color = "#e66419"
                case "ON HOLD":
                    curr_color = "#808080"
                case "RED": 
                    curr_color = "#ff0000"
            circle = plt.Circle((p[1]*math.cos(p[0]) + IMG_WIDTH//2, IMG_HEIGHT//2 - p[1]*math.sin(p[0])), 8, color=curr_color, fill=True)
            ax.add_artist(circle)
            ax.text(p[1]*math.cos(p[0]) + IMG_WIDTH//2,  IMG_HEIGHT//2 - p[1]*math.sin(p[0]), p[2], ha="center", va="center", fontsize=6, color="white")

    # Save the figure without borders
    fig.savefig("radar.png", dpi=250, bbox_inches="tight")

    plt.show()
    

if __name__ == "__main__":
    main()
