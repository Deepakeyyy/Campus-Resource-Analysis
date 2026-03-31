from __future__ import annotations

import argparse
import random
from datetime import date, timedelta

import numpy as np
import pandas as pd


def _pick_course_id(dept: str, i: int) -> str:
    prefix = {
        "Operations": "OPS",
        "Mechanical": "MEC",
        "Computer Science": "CSE",
        "Electronics": "ECE",
    }[dept]
    return f"{prefix}-{100 + (i % 400):03d}"


def make_resources(rng: random.Random) -> pd.DataFrame:
    buildings = [
        "Main Block",
        "Engineering Block",
        "Innovation Hub",
        "Science Wing",
        "Tech Center",
        "Workshop Complex",
    ]
    room_types = [
        "Smart Classroom",
        "CNC Lab",
        "Surface Finishing Lab",
        "Seminar Hall",
        "Electronics Lab",
        "Computer Lab",
        "Project Studio",
    ]

    n_rooms = rng.randint(30, 50)
    rooms = []
    for i in range(1, n_rooms + 1):
        room_id = f"R{str(i).zfill(3)}"
        building = rng.choice(buildings)
        rtype = rng.choice(room_types)

        if rtype == "Seminar Hall":
            cap = rng.randint(120, 200)
        elif "Lab" in rtype:
            cap = rng.randint(20, 60)
        else:
            cap = rng.randint(30, 120)

        rooms.append(
            {
                "Room_ID": room_id,
                "Building": building,
                "Type": rtype,
                "Capacity": int(cap),
            }
        )
    return pd.DataFrame(rooms)


def make_schedule(rng: random.Random, resources: pd.DataFrame) -> pd.DataFrame:
    days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]
    # 8:00 AM to 5:00 PM, hourly slots (10 values)
    times = [
        "08:00",
        "09:00",
        "10:00",
        "11:00",
        "12:00",
        "13:00",
        "14:00",
        "15:00",
        "16:00",
        "17:00",
    ]
    departments = ["Operations", "Mechanical", "Computer Science", "Electronics"]

    room_ids = resources["Room_ID"].astype(str).tolist()

    n_rows = 500
    rows = []
    for i in range(1, n_rows + 1):
        slot_id = f"S{str(i).zfill(4)}"
        day = rng.choice(days)
        t = rng.choice(times)
        room_id = rng.choice(room_ids)
        dept = rng.choice(departments)
        course_id = _pick_course_id(dept, i)
        rows.append(
            {
                "Slot_ID": slot_id,
                "Day": day,
                "Time": t,
                "Room_ID": room_id,
                "Course_ID": course_id,
                "Department": dept,
            }
        )
    return pd.DataFrame(rows)


def make_utilization(
    rng: random.Random,
    resources: pd.DataFrame,
    schedule: pd.DataFrame,
    *,
    n_rows: int = 1000,
    n_days: int = 30,
) -> pd.DataFrame:
    # Many rows referencing the schedule Slot_IDs (many-to-one)
    # Spread across days to feel real.
    start = date(2026, 1, 5)  # Monday-ish
    dates = [start + timedelta(days=i) for i in range(0, int(n_days))]

    cap_by_room = resources.set_index("Room_ID")["Capacity"].to_dict()
    room_by_slot = schedule.set_index("Slot_ID")["Room_ID"].to_dict()
    slot_ids = schedule["Slot_ID"].astype(str).tolist()

    rows = []
    for i in range(int(n_rows)):
        slot_id = rng.choice(slot_ids)
        room_id = str(room_by_slot[slot_id])
        capacity = int(cap_by_room.get(room_id, 60))

        d = rng.choice(dates)

        # Base: most sessions are moderately filled.
        base_ratio = rng.triangular(0.15, 0.75, 0.55)
        attendance = int(round(capacity * base_ratio))

        # Inject Ghost Rooms: very low attendance in large rooms.
        if capacity >= 120 and rng.random() < 0.12:
            attendance = rng.randint(0, 12)  # e.g. 5 students in 150-seat hall

        # Inject Overcrowding: occasionally exceed capacity.
        if rng.random() < 0.06:
            attendance = capacity + rng.randint(1, max(3, int(capacity * 0.15)))

        # Small noise
        attendance = max(0, attendance + rng.randint(-2, 3))

        rows.append(
            {
                "Slot_ID": slot_id,
                "Date": pd.Timestamp(d),
                "Actual_Attendance": int(attendance),
            }
        )

    util = pd.DataFrame(rows)

    # Ensure we have at least a few explicit ghost + over-capacity cases.
    # (Deterministic given seed; but we harden with a gentle patch-up.)
    if util["Actual_Attendance"].max() <= 0:
        util.loc[0, "Actual_Attendance"] = 10

    return util


def main() -> None:
    parser = argparse.ArgumentParser(description="Generate synthetic CampusFlow Excel backend.")
    parser.add_argument("--seed", type=int, default=42)
    parser.add_argument("--out", type=str, default="Campus_Data.xlsx")
    parser.add_argument("--utilization-rows", type=int, default=1000)
    parser.add_argument("--utilization-days", type=int, default=30)
    args = parser.parse_args()

    rng = random.Random(args.seed)
    np.random.seed(args.seed)

    resources = make_resources(rng)
    schedule = make_schedule(rng, resources)
    utilization = make_utilization(
        rng,
        resources,
        schedule,
        n_rows=int(args.utilization_rows),
        n_days=int(args.utilization_days),
    )

    out_path = str(args.out)
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        resources.to_excel(writer, index=False, sheet_name="Resources")
        schedule.to_excel(writer, index=False, sheet_name="Schedule")
        utilization.to_excel(writer, index=False, sheet_name="Utilization")

    # Quick integrity checks
    assert utilization["Slot_ID"].isin(schedule["Slot_ID"]).all()
    assert schedule["Room_ID"].isin(resources["Room_ID"]).all()

    print(f"Wrote {out_path}")
    print(f"Resources: {len(resources)} rows")
    print(f"Schedule: {len(schedule)} rows")
    print(f"Utilization: {len(utilization)} rows")


if __name__ == "__main__":
    main()

