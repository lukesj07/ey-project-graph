import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.image as mpimg
import numpy as np
import math

from openpyxl import load_workbook

DATA_PATH = "spreadsheet.xlsx"
IMG_PATH = "img.jpg"
IMG_WIDTH = 700
IMG_HEIGHT = 720

def sort_project_status(df: pd.DataFrame) -> dict[str, list[str]]:
    statuses = {"GREEN": [], "AMBER": [], "ON HOLD": [], "RED": []}
    
    for index, row in df.iterrows():
        health = str(row["Overall Health"]).upper()
        
        if health in statuses:
            statuses[health].append(row["Project Name"])
        elif health not in ["CANCELED", "CLOSED", "-"]:
            statuses["GREEN"].append(row["Project Name"])  # default to GREEN for any unspecified health
            
    return statuses

def calculate_position(percent: float, positions: list[list[list]], idx: int, angle_bounds: list[float], name: str) -> list[float]:
    while True:
        overlapping = []
        if 0.75 <= percent <= 1.0:
            out = (1 / (0.25)) * (percent - 0.75)
            ideal_radius = round(150 * (1 - out) + 25)
        else:
            out = (1 / 0.75) * percent
            ideal_radius = round(160 * (1 - out) + 150)

        max_points = (((angle_bounds[1] - angle_bounds[0]) * ideal_radius) // 12) - 1

        for p in positions[idx]:
            if abs(p[1] - ideal_radius) < 16 and (angle_bounds[0] - 0.1) <= p[0] <= (angle_bounds[1] + 0.1):
                overlapping.append(p)
    
        if len(overlapping) == 0:
            return [ideal_radius, angle_bounds[0] + (15 / (ideal_radius * 2))]

        if len(overlapping) >= max_points:
            percent -= 0.01
        else:
            max_a = max([p[0] for p in overlapping])
            if max_a + (20 / (ideal_radius * 2)) > angle_bounds[1]:
                percent -= 0.01
            else:
                return [ideal_radius, max_a + (30 / (ideal_radius * 2))]

def plot_radar_chart(df: pd.DataFrame, positions: list[list[list[float, float, str, str]]]) -> None:
    fig, ax = plt.subplots(frameon=False)
    ax.set_aspect("equal", adjustable="box")
    ax.set_axis_off()
    ax.imshow(mpimg.imread(IMG_PATH), aspect="auto")

    for category_positions in positions:
        for theta, r, radar_id, health in category_positions:
            color_map = {
                "GREEN": "#70ad46",
                "AMBER": "#ffc000",
                "ON HOLD": "#7f7f7f",
                "RED": "#c00000",
            }
            curr_color = color_map.get(health, "#70ad46")  # default to green if health isn't found

            circle = plt.Circle(
                (r * math.cos(theta) + IMG_WIDTH // 2, IMG_HEIGHT // 2 - r * math.sin(theta)),
                8, color=curr_color, fill=True
            )
            ax.add_artist(circle)
            ax.text(r * math.cos(theta) + IMG_WIDTH // 2, IMG_HEIGHT // 2 - r * math.sin(theta),
                    radar_id, ha="center", va="center", fontsize=5, color="white")

    fig.savefig("radar.png", dpi=250, bbox_inches="tight")

def main() -> None:
    df = pd.read_excel(DATA_PATH)
    df = df.dropna(how="all", axis=1)
    df.columns = df.iloc[0].tolist()
    df = df[1:].reset_index(drop=True)
    
    statuses = sort_project_status(df)
    positions = [[] for _ in range(8)]  # 8 sectors in the radar chart
    
    for status_idx, (status_name, projects) in enumerate(statuses.items()):
        for project_name in projects:
            row_num = df[df["Project Name"] == project_name].index[0]
            percent = df["%Project Duration Completed"][row_num]

            # sector bounds can be improved by defining them in a separate data structure
            sectors = {
                "InfoSec Protection Services": [0 + 0.03, math.pi / 4],
                "IT Risk Management": [math.pi / 4 + 0.05, math.pi / 2],
                "Identity and Access": [math.pi / 2, 3 * math.pi / 4],
                "Threat Management": [3 * math.pi / 4 + 0.05, math.pi],
                "InfoSec Program Management": [math.pi + 0.01, 5 * math.pi / 4],
                "InfoSec Program Support": [5 * math.pi / 4 + 0.05, 3 * math.pi / 2],
                "Security Design Services": [3 * math.pi / 2 + 0.05, 7 * math.pi / 4],
                "Compliance and Assurance": [7 * math.pi / 4 + 0.1, 2 * math.pi],
            }

            service_category = df["Service Category"][row_num]
            if service_category in sectors:
                angle_bounds = sectors[service_category]
                r, a = calculate_position(percent, positions, status_idx, angle_bounds, project_name)
                positions[status_idx].append([a, r, df["Radar ID"][row_num], df["Overall Health"][row_num]])
    
    plot_radar_chart(df, positions)

if __name__ == "__main__":
    main()
