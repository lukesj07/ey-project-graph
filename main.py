import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.image as mpimg
import math

DATA_PATH = "Project Radar Spreadsheet.xlsx"
IMG_PATH = "FY25 radar, active projects.jpg"
IMG_WIDTH = 872
IMG_HEIGHT = 856

def sort_project_status(df: pd.DataFrame) -> dict[str, list[str]]:
    statuses = {"GREEN": [], "AMBER": [], "ON HOLD": [], "RED": []}
    for _, row in df.iterrows():
        health = row["Overall Health"]

        if pd.isna(health) or health == "":
            statuses["GREEN"].append(row["Project Name"])
        else:
            health = health.upper()
            if health in statuses:
                statuses[health].append(row["Project Name"])
    
    return statuses

def calculate_position(percent: float, positions: list[list[float]], angle_bounds: list[float]) -> list[float]:
    while True:
        overlapping = []
        if 0.75 <= percent <= 1.0:
            out = (percent - 0.75) / 0.25
            ideal_radius = round(150 * (1 - out) + 25)
        else:
            out = percent / 0.75
            ideal_radius = round(160 * (1 - out) + 150)

        max_points = ((angle_bounds[1] - angle_bounds[0]) * ideal_radius) // 12 - 1

        for p in positions:
            if abs(p[1] - ideal_radius) < 16 and (angle_bounds[0] - 0.1) <= p[0] <= (angle_bounds[1] + 0.1):
                overlapping.append(p)
    
        if not overlapping:
            return [ideal_radius, angle_bounds[0] + (15 / (ideal_radius * 2))]

        if len(overlapping) >= max_points:
            percent -= 0.01
        else:
            max_a = max(p[0] for p in overlapping)
            if max_a + (20 / (ideal_radius * 2)) > angle_bounds[1]:
                percent -= 0.01
            else:
                return [ideal_radius, max_a + (30 / (ideal_radius * 2))]

def plot_radar_chart(df: pd.DataFrame, positions: dict[str, list[list[float, float, str, str]]]) -> None:
    fig, ax = plt.subplots(figsize=(IMG_WIDTH / 100, IMG_HEIGHT / 100), frameon=False)
    ax.set_aspect('equal', adjustable='datalim')
    ax.set_axis_off()
    img = mpimg.imread(IMG_PATH)
    ax.imshow(img, extent=[0, IMG_WIDTH, 0, IMG_HEIGHT])

    color_map = {
        "GREEN": "#70ad46",
        "AMBER": "#ffc000",
        "ON HOLD": "#7f7f7f",
        "RED": "#c00000",
    }

    for sector, category_positions in positions.items():
        for theta, r, radar_id, health in category_positions:
            color = color_map.get(health, "#70ad46")  # default to green if health isn't recognized
            
            radar_id_str = str(radar_id)
            radar_id_display = radar_id_str[-2:] if len(radar_id_str) > 2 else radar_id_str

            x = r * math.cos(theta) + IMG_WIDTH / 2
            y = r * math.sin(theta) + IMG_HEIGHT / 2

            circle = plt.Circle((x, y), 8, color=color, fill=True)
            ax.add_artist(circle)
            ax.text(x, y, radar_id_display, ha="center", va="center", fontsize=5, color="white")

    fig.savefig("radar.png", dpi=250, bbox_inches='tight', pad_inches=0)
    plt.close(fig)

def main() -> None:
    df = pd.read_excel(DATA_PATH)
    df.dropna(how="all", axis=1, inplace=True)
    df.columns = df.iloc[0].tolist()
    df = df[1:].reset_index(drop=True)
    
    duplicate_projects = df[df.duplicated(subset=["Project Name", "Strategic Priority: Primary"], keep=False)]
    if not duplicate_projects.empty:
        print("Duplicate Projects Found:\n", duplicate_projects)

    statuses = sort_project_status(df)
    
    sectors = {
        "1.1 Governance accountability": [math.pi / 2, 13 * math.pi / 18],
        "1.2 Strategic traceability": [5 * math.pi / 18, math.pi / 2],
        "1.3 Strategic alignment": [math.pi / 18 + 0.1, 5 * math.pi / 18],
        "2.1 Scalable simplicity": [11 * math.pi / 6, 2 * math.pi + math.pi / 18],
        "2.2 Automation & self-service": [29 * math.pi / 18, 11 * math.pi / 6],
        "2.3 Collaborative empowerment": [25 * math.pi / 18, 29 * math.pi / 18],
        "3.1 Stakeholder-enabling integration": [21 * math.pi / 18, 25 * math.pi / 18],
        "3.2 Stakeholder centricity": [17 * math.pi / 18, 21 * math.pi / 18],
        "3.3 Empowered security culture": [13 * math.pi / 18, 17 * math.pi / 18]
    }

    positions = {sector: [] for sector in sectors}

    for status_name, projects in statuses.items():
        for project_name in projects:
            row = df[df["Project Name"] == project_name].iloc[0]
            percent = row["%Project Duration Completed"]
            service_category = row["Strategic Priority: Primary"]
            
            if service_category in sectors:
                angle_bounds = sectors[service_category]
                r, a = calculate_position(percent, positions[service_category], angle_bounds)
                
                positions[service_category].append([a, r, row["2 digit Radar ID"], row["Overall Health"]])

    plot_radar_chart(df, positions)

if __name__ == "__main__":
    main()
