"""Verify the recorded PowerPoint acceptance run; does not control Office."""
import argparse
import copy
import datetime as dt
import json
import math
from pathlib import Path
import re
import sys

EPSILON = 0.01
SNAPSHOTS = ("owned-drag-before", "owned-drag-after", "undo-b", "undo-a",
             "redo-a", "redo-b", "before-restart", "after-restart")


def require(condition, message):
    if not condition:
        raise ValueError(message)


def timestamp(value):
    value = re.sub(r"(\.\d{6})\d+", r"\1", value)
    return dt.datetime.fromisoformat(value.replace("Z", "+00:00"))


def close(a, b):
    a, b = float(a), float(b)
    return math.isfinite(a) and math.isfinite(b) and abs(a - b) <= EPSILON


def geometry(snapshot):
    shapes = snapshot["Shapes"]
    require(len(shapes) == 2, "Snapshot must contain exactly two shapes")
    result = {shape["Tag"].upper(): shape for shape in shapes}
    require(len(result) == 2 and all(result), "Shape GUID tags must be distinct and nonempty")
    return result


def same_geometry(actual, expected, label):
    require(set(actual) == set(expected), label + ": shape tags changed")
    for tag, shape in expected.items():
        require(actual[tag]["Name"] == shape["Name"], label + ": shape name changed")
        for field in ("Left", "Top", "Width", "Height"):
            require(close(actual[tag][field], shape[field]),
                    "{}: {}.{} expected {}, got {}".format(
                        label, shape["Name"], field, shape[field], actual[tag][field]))


def verify(directory):
    required = [directory / (name + ".json") for name in SNAPSHOTS]
    required.append(directory / "owned-drag.log")
    require(all(path.is_file() for path in required),
            "Missing evidence: " + ", ".join(path.name for path in required if not path.is_file()))
    snapshots = {name: json.loads((directory / (name + ".json")).read_text(encoding="utf-8-sig"))
                 for name in SNAPSHOTS}
    times = [timestamp(snapshots[name]["Utc"]) for name in SNAPSHOTS]
    require(all(left < right for left, right in zip(times, times[1:])), "Evidence snapshot sequence is stale or out of order")
    before, after = snapshots[SNAPSHOTS[0]], snapshots[SNAPSHOTS[1]]
    start, stop = timestamp(before["Utc"]), timestamp(after["Utc"])
    require(start < stop, "Main-run timestamps are reversed or equal")
    require(before["ProcessId"] == after["ProcessId"], "Main run changed PowerPoint process")
    initial = geometry(before)
    by_name = {shape["Name"]: tag for tag, shape in initial.items()}
    require(set(by_name) == {"GCS_Rectangle_A", "GCS_Rectangle_B"}, "Unexpected test shape names")
    a, b = by_name["GCS_Rectangle_A"], by_name["GCS_Rectangle_B"]
    require(close(initial[a]["Top"], initial[b]["Top"]), "Initial rectangles are not aligned")
    require(len(before["ConstraintXml"]) == 1, "Exactly one constraint XML part is required")
    state, states, drags, frames = copy.deepcopy(initial), [], [], []
    for line in (directory / "owned-drag.log").read_text(encoding="utf-8-sig").splitlines():
        parts = line.split(" ", 2)
        if len(parts) != 3 or not start <= timestamp(parts[0]) <= stop:
            continue
        event, data = parts[1], dict(re.findall(r"(\w+)=([^\s]+)", parts[2]))
        require(not event.endswith("error"), "Runtime error during main run: " + line)
        if event == "drag-start":
            frames = []
        elif event == "drag-frame":
            frames.append(data)
        elif event == "drag-end":
            if int(data["frames"]) == 0 or data["cancelled"] != "False":
                frames = []
                continue
            driver = data["driver"].upper()
            require(driver in state, "Drag driver does not match a recorded shape GUID")
            follower = b if driver == a else a
            held = [frame for frame in frames if frame["mouseDown"] == "True"]
            require(len(held) > 1, "Need multiple applied frames while mouse button remains down")
            require(len(frames) == int(data["frames"]), "Drag frame count does not match end record")
            for frame in frames:
                require(frame["driver"].upper() == driver, "Frame belongs to another driver")
                require(close(frame["a"], frame["b"]), "Rectangles diverged during a drag")
                require(close(frame["followerLeft"], state[follower]["Left"]), "Follower moved horizontally")
            require(close(data["leftBefore"], state[driver]["Left"]), "Drag start Left disagrees with prior state")
            require(close(data["aBefore"], state[a]["Top"]) and close(data["bBefore"], state[b]["Top"]),
                    "Drag start Top disagrees with prior state")
            require(close(data["aAfter"], data["bAfter"]), "Rectangles diverged at drag end")
            state[driver]["Left"] = float(data["leftAfter"])
            state[a]["Top"], state[b]["Top"] = float(data["aAfter"]), float(data["bAfter"])
            states.append(copy.deepcopy(state))
            drags.append({"driver": driver, "frames": len(frames), "mouseDownFrames": len(held)})
            frames = []
    require(len(drags) == 2, "Expected exactly two meaningful completed drags, got " + str(len(drags)))
    require([drag["driver"] for drag in drags] == [a, b], "Main run must drag A then B")
    expected = {"owned-drag-after": states[1], "undo-b": states[0], "undo-a": initial,
                "redo-a": states[0], "redo-b": states[1]}
    for name, value in expected.items():
        same_geometry(geometry(snapshots[name]), value, name)
        require(snapshots[name]["ConstraintXml"] == before["ConstraintXml"], name + ": constraint XML changed")
        require(snapshots[name]["ProcessId"] == before["ProcessId"], name + ": PowerPoint process changed")
    saved, reopened = snapshots["before-restart"], snapshots["after-restart"]
    same_geometry(geometry(saved), states[1], "before-restart")
    same_geometry(geometry(reopened), geometry(saved), "after-restart")
    require(saved["ConstraintXml"] == reopened["ConstraintXml"] == before["ConstraintXml"],
            "Constraint XML changed across save/reopen")
    require(saved["File"] == reopened["File"] and saved["SlideId"] == reopened["SlideId"],
            "Restart snapshots refer to different documents or slides")
    require(len(saved["ProcessId"]) == len(reopened["ProcessId"]) == 1 and
            saved["ProcessId"][0] != reopened["ProcessId"][0], "PowerPoint process was not restarted")
    return {"passed": True, "tolerancePt": EPSILON, "drags": drags,
            "checks": ["bidirectional live frames", "follower horizontal position", "two undo steps",
                       "two redo steps", "geometry and constraints after process restart"]}


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--artifacts", type=Path, default=Path(__file__).resolve().parents[1] / "artifacts")
    parser.add_argument("--output", type=Path, default=Path("acceptance.json"))
    args = parser.parse_args()
    try:
        result = verify(args.artifacts)
    except (ValueError, KeyError, TypeError, OSError) as error:
        result = {"passed": False, "error": str(error)}
    result["artifacts"] = str(args.artifacts.resolve())
    result["checkedUtc"] = dt.datetime.now(dt.timezone.utc).isoformat()
    output = args.output if args.output.is_absolute() else args.artifacts / args.output
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(result, ensure_ascii=False, indent=2))
    sys.exit(0 if result["passed"] else 1)
