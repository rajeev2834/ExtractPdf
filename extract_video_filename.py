import os
import re
import json
import subprocess
import pandas as pd
from datetime import datetime

def to_dt(s):
    if not s:
        return None
    return datetime.strptime(s, "%Y-%m-%d %H:%M:%S")

def split_duration(td):
    if td is None:
        return None, None, None
    total_seconds = int(td.total_seconds())
    hours = total_seconds // 3600
    minutes = (total_seconds % 3600) // 60
    seconds = total_seconds % 60
    return hours, minutes, seconds


def format_timedelta(td):
    if td is None:
        return None
    total_seconds = int(td.total_seconds())
    hours = total_seconds // 3600
    minutes = (total_seconds % 3600) // 60
    seconds = total_seconds % 60
    return f"{hours:02d}:{minutes:02d}:{seconds:02d}"

# Regex to extract timestamp from filename
pattern = re.compile(r".*-(\d{4})-(\d{2})-(\d{2})-(\d{2})-(\d{2})-(\d{2})")

def extract_filename_datetime(filename):
    match = pattern.match(filename)
    if match:
        y, mo, d, h, mi, s = match.groups()
        return f"{y}-{mo}-{d} {h}:{mi}:{s}"
    return None

def get_creation_time(file_path):
    ts = os.path.getctime(file_path)
    return datetime.fromtimestamp(ts).strftime("%Y-%m-%d %H:%M:%S")

def get_duration_ffprobe(file_path):
    try:
        cmd = [
            "ffprobe",
            "-v", "quiet",
            "-print_format", "json",
            "-show_format",
            file_path
        ]

        output = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
        data = json.loads(output.stdout)
        return float(data["format"]["duration"])
    except:
        return None

def scan_nawada(root_path):
    rows = []

    for booth in os.listdir(root_path):
        booth_path = os.path.join(root_path, booth)
        if not os.path.isdir(booth_path):
            continue

        for direction in ["IN", "OUT"]:
            dir_path = os.path.join(booth_path, direction)
            if not os.path.isdir(dir_path):
                continue

            for camera in os.listdir(dir_path):
                cam_path = os.path.join(dir_path, camera)
                if not os.path.isdir(cam_path):
                    continue

                files = [
                    f for f in os.listdir(cam_path)
                    if f.lower().endswith((".mp4", ".avi", ".mkv"))
                ]
                files.sort()

                if not files:
                    continue

                first_file = files[0]
                last_file = files[-1]

                first_path = os.path.join(cam_path, first_file)
                last_path = os.path.join(cam_path, last_file)

                first_time = extract_filename_datetime(first_file)
                last_time = extract_filename_datetime(last_file)

                first_creation = get_creation_time(first_path)
                last_creation = get_creation_time(last_path)

                first_duration = get_duration_ffprobe(first_path)
                last_duration = get_duration_ffprobe(last_path)

                first_dt = to_dt(first_time)
                last_dt = to_dt(last_time)

                time_gap = last_dt - first_dt if first_dt and last_dt else None
                hours, minutes, seconds = split_duration(time_gap)

                rows.append({
                    "Booth": booth,
                    "Direction": direction,
                    "Camera": camera,
                    "Total Files": len(files),

                    "First File": first_file,
                    "First File (Filename Time)": first_time,
                    "First File (Creation Time)": first_creation,
                    "First File Duration (sec)": first_duration,

                    "Last File": last_file,
                    "Last File (Filename Time)": last_time,
                    "Last File (Creation Time)": last_creation,
                    "Last File Duration (sec)": last_duration,

                    "Total Duration (HH:MM:SS)": format_timedelta(time_gap),
                    "Total Duration (Seconds)": int(time_gap.total_seconds()) if time_gap else None,
                    "Duration Hours": hours,
                    "Duration Minutes": minutes,
                    "Duration Seconds": seconds,      

                    "Folder Path": cam_path
                })

    return pd.DataFrame(rows)


if __name__ == "__main__":
    root = r"F:\GOBINDPUR"
    df = scan_nawada(root)
    df.to_excel("gobindpur_video_report.xlsx", index=False)
    print("Report generated: nawada_video_report.xlsx")