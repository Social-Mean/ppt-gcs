"""Check recorded native PowerPoint layout-v2 evidence; never generates gestures."""
import argparse
import copy
import datetime as dt
import json
import math
from pathlib import Path
import re
import sys
import xml.etree.ElementTree as ET

EPSILON = 0.01
FIELDS = ("Left", "Top", "Width", "Height")
KINDS = ("SameY", "SameX", "SameWidth", "SameHeight", "EqualHorizontalGap", "EqualVerticalGap")
AXIS = dict(zip(KINDS, ("Top", "Left", "Width", "Height", "Left", "Top")))


def require(condition, message):
    if not condition:
        raise ValueError(message)


def timestamp(value):
    return dt.datetime.fromisoformat(re.sub(r"(\.\d{6})\d+", r"\1", value).replace("Z", "+00:00"))


def close(a, b):
    return math.isfinite(float(a)) and math.isfinite(float(b)) and abs(float(a) - float(b)) <= EPSILON


def geometries(items):
    result = {item["Id"].upper(): item for item in items}
    require(len(result) == len(items) and all(result), "Duplicate or empty shape identity")
    for shape in result.values():
        require(all(math.isfinite(float(shape[field])) for field in FIELDS), "Non-finite geometry")
        require(shape["Width"] >= 1 - EPSILON and shape["Height"] >= 1 - EPSILON, "Invalid shape size")
    return result


def same_geometry(actual, expected, label):
    require(set(actual) == set(expected), label + ": shape identities changed")
    for identity, shape in expected.items():
        for field in FIELDS:
            require(close(actual[identity][field], shape[field]),
                    "{}: {}.{} expected {}, got {}".format(label, identity, field, shape[field], actual[identity][field]))


def requested_geometry(begin, point):
    """Independent pointer/handle arithmetic shared by native and replay checks."""
    driver = begin["driverId"].upper()
    result = copy.deepcopy(geometries(begin["before"])[driver])
    scale = begin["scale"]
    require(all(math.isfinite(float(scale[axis])) and scale[axis] > 0 for axis in ("x", "y")), "Invalid screen scale")
    require(all(math.isfinite(float(cursor[axis])) for cursor in (point, begin["cursor"]) for axis in ("x", "y")), "Invalid pointer coordinates")
    dx = (point["x"] - begin["cursor"]["x"]) / scale["x"]
    dy = (point["y"] - begin["cursor"]["y"]) / scale["y"]
    mode = begin["mode"]
    if mode == "Move":
        result["Left"] += dx
        result["Top"] += dy
        return result
    require(mode in ("ResizeNW", "ResizeN", "ResizeNE", "ResizeE", "ResizeSE", "ResizeS", "ResizeSW", "ResizeW"), "Unknown resize mode")
    west, east = mode in ("ResizeNW", "ResizeW", "ResizeSW"), mode in ("ResizeNE", "ResizeE", "ResizeSE")
    north, south = mode in ("ResizeNW", "ResizeN", "ResizeNE"), mode in ("ResizeSW", "ResizeS", "ResizeSE")
    old_width, old_height = result["Width"], result["Height"]
    if west or east:
        result["Width"] = max(1.0, old_width + (-dx if west else dx))
    if north or south:
        result["Height"] = max(1.0, old_height + (-dy if north else dy))
    if west:
        result["Left"] += old_width - result["Width"]
    if north:
        result["Top"] += old_height - result["Height"]
    return result


def same_state(actual, expected, label):
    require(set(actual) == set(expected), label + ": slides changed")
    for slide_id in expected:
        same_geometry(actual[slide_id], expected[slide_id], label + " slide " + str(slide_id))


def check_constraints(geometry, constraints, label):
    for constraint in constraints:
        kind, ids = constraint["Kind"], [value.upper() for value in constraint["ShapeIds"]]
        require(kind in KINDS and all(value in geometry for value in ids), label + ": unknown kind or shape")
        require(len(ids) >= (3 if "Gap" in kind else 2), label + ": too few constraint members")
        if "Gap" in kind:
            position, size = ("Left", "Width") if kind == "EqualHorizontalGap" else ("Top", "Height")
            values = [geometry[right][position] - geometry[left][position] - geometry[left][size]
                      for left, right in zip(ids, ids[1:])]
        else:
            values = [geometry[value][AXIS[kind]] for value in ids]
        require(all(close(value, values[0]) for value in values), label + ": violated " + kind)


def load_snapshot(directory, name):
    path = directory / (name + ".json")
    require(path.is_file(), "Missing evidence: " + path.name)
    snapshot = json.loads(path.read_text(encoding="utf-8-sig"))
    require(len(snapshot["slides"]) == 6, name + ": expected six fixture slides")
    require([slide["kind"] for slide in snapshot["slides"]] == list(KINDS), name + ": wrong fixture cases")
    require(len(snapshot["ProcessId"]) == 1, name + ": expected one PowerPoint process")
    require(len(snapshot["ConstraintXml"]) == 1, name + ": expected one XML part")
    require(ET.fromstring(snapshot["ConstraintXml"][0]).tag == "{urn:ppt-gcs:v2}constraints", name + ": wrong XML schema")
    return snapshot


def snapshot_state(snapshot):
    result = {slide["slideId"]: geometries(slide["geometry"]) for slide in snapshot["slides"]}
    require(len(result) == 6, "Duplicate slide IDs")
    return result


def metadata(snapshot):
    return [(slide["slideId"], slide["name"], slide["kind"], slide["constraints"],
             [(shape["Id"], shape["Name"]) for shape in slide["geometry"]]) for slide in snapshot["slides"]]


def verify(directory):
    before, after = load_snapshot(directory, "before"), load_snapshot(directory, "after")
    start, stop = timestamp(before["Utc"]), timestamp(after["Utc"])
    require(start < stop, "Main run snapshot times are reversed or equal")
    initial, state = snapshot_state(before), snapshot_state(before)
    slides = {slide["slideId"]: slide for slide in before["slides"]}
    for slide_id, slide in slides.items():
        check_constraints(initial[slide_id], slide["constraints"], "Initial slide")
    log_path = directory / "gestures.log"
    require(log_path.is_file(), "Missing evidence: gestures.log")
    active, states, gestures = None, [copy.deepcopy(state)], []
    coverage = {kind: set() for kind in KINDS}
    for line in log_path.read_text(encoding="utf-8-sig").splitlines():
        parts = line.split(" ", 2)
        if len(parts) != 3 or not re.match(r"^\d{4}-\d{2}-\d{2}T", parts[0]):
            continue
        if not start <= timestamp(parts[0]) <= stop:
            continue
        event = parts[1]
        require(not event.endswith("error"), "Runtime error in main run: " + line)
        if event not in ("gesture-start", "gesture-frame", "gesture-release", "gesture-end"):
            continue
        data = json.loads(parts[2])
        if event == "gesture-start":
            require(active is None, "Overlapping gesture starts")
            require(data["slideId"] in slides, "Gesture refers to another slide")
            require(data.get("test") is False, "Programmatic replay cannot count as real pointer evidence")
            active = {"start": data, "frames": [], "release": None}
        else:
            require(active is not None and active["start"]["gestureId"] == data["gestureId"], "Unmatched gesture event")
            if event == "gesture-frame":
                require(active["release"] is None, "Frame appears after the release record")
                active["frames"].append(data)
                continue
            if event == "gesture-release":
                require(active["release"] is None, "Duplicate release record")
                active["release"] = data
                continue
            begin, frames, release = active["start"], active["frames"], active["release"]
            active = None
            sid, driver = begin["slideId"], begin["driverId"].upper()
            require(data["driverId"].upper() == driver and data["mode"] == begin["mode"], "Gesture identity changed")
            require(data.get("test") is False, "Programmatic gesture end mixed into real pointer evidence")
            prior, end = geometries(begin["before"]), geometries(data["after"])
            same_geometry(prior, state[sid], "Gesture baseline")
            same_geometry(geometries(data["before"]), prior, "Gesture end baseline")
            require(len(frames) == data["frames"], "Frame count disagrees with gesture end")
            if data["cancelled"]:
                same_geometry(end, prior, "Cancelled gesture")
                require(not frames, "Cancelled mutation creates ambiguous native undo evidence; rerun without cancellation")
                continue
            require(isinstance(data["dragThresholdPassed"], bool), "Missing or invalid drag-threshold state")
            if data["dragThresholdPassed"] is False:
                require(begin["mode"] == "Move" and not frames, "Below-threshold click unexpectedly mutated geometry")
                same_geometry(end, prior, "Below-threshold click")
                continue
            if not frames:
                same_geometry(end, prior, "Click-only gesture")
                require(release is not None, "Click-only gesture has no release record")
                same_geometry({driver: end[driver]}, {driver: requested_geometry(begin, release["releasePoint"])}, "Click-only release target")
                continue
            require(release is not None and release["geometryMatched"] is True and release.get("error") is None,
                    "Release position was not verified against the final motion")
            require(release["driverId"].upper() == driver and release["mode"] == begin["mode"] and release["frames"] == len(frames),
                    "Release identity or frame count changed")
            require(driver in prior, "Driver is absent from baseline")
            held = [frame for frame in frames if frame["mouseDown"] is True]
            constraints = slides[sid]["constraints"]
            for index, frame in enumerate(frames, 1):
                require(frame["frame"] == index and frame["driverId"].upper() == driver and
                        frame["mode"] == begin["mode"], "Frame sequence/driver/mode mismatch")
                require(frame.get("input") == "owned-gesture-loop", "Geometry was written outside the owned gesture loop")
                require(isinstance(frame.get("mouseDown"), bool) and frame.get("pointerMessage") in ("move", "up"), "Invalid pointer sample")
                require(frame["pointerMessage"] != "up" or frame["mouseDown"] is False, "Released pointer frame is marked held")
                values = geometries(frame["geometry"])
                require(set(values) == set(prior), "Frame shape identities changed")
                check_constraints(values, constraints, "Frame " + str(index))
                same_geometry({driver: values[driver]}, {driver: requested_geometry(begin, frame["cursor"])}, "Independent frame endpoint " + str(index))
                if begin["mode"] == "Move":
                    for identity in values:
                        for dimension in ("Width", "Height"):
                            require(close(values[identity][dimension], prior[identity][dimension]), "Moving resized a shape")
                relevant = [constraint for constraint in constraints if driver in [v.upper() for v in constraint["ShapeIds"]]]
                if len(relevant) == 1:
                    kind = relevant[0]["Kind"]
                    unaffected = {"SameX": ("Top",), "SameY": ("Left",), "SameWidth": ("Left", "Top", "Height"),
                                  "SameHeight": ("Left", "Top", "Width"), "EqualHorizontalGap": ("Top",), "EqualVerticalGap": ("Left",)}[kind]
                    for identity in values:
                        if identity != driver:
                            require(all(close(values[identity][field], prior[identity][field]) for field in unaffected), "Unrelated follower geometry changed")
            same_geometry(end, geometries(frames[-1]["geometry"]), "Final frame")
            same_geometry({driver: end[driver]}, {driver: requested_geometry(begin, release["releasePoint"])}, "Independent release endpoint")
            check_constraints(end, constraints, "Gesture end")
            kind = slides[sid]["kind"]
            mode_matches = begin["mode"].startswith("Resize") if kind in ("SameWidth", "SameHeight") else begin["mode"] == "Move"
            qualifies = (len(held) >= 2 and mode_matches and
                         not close(prior[driver][AXIS[kind]], end[driver][AXIS[kind]]))
            if qualifies:
                coverage[kind].add(driver)
            state[sid] = end
            states.append(copy.deepcopy(state))
            gestures.append({"gestureId": begin["gestureId"], "kind": kind, "driverId": driver,
                             "mode": begin["mode"], "frames": len(frames), "mouseDownFrames": len(held),
                             "qualifiesForCoverage": qualifies})
    require(active is None, "Main run ends during a gesture")
    require(all(len(drivers) >= 2 for drivers in coverage.values()), "Each constraint kind needs two distinct driven shapes")
    sequence = [("after", after, states[-1])]
    count = len(gestures)
    sequence += [("undo-{:02d}".format(i), load_snapshot(directory, "undo-{:02d}".format(i)), states[count-i]) for i in range(1, count+1)]
    sequence += [("redo-{:02d}".format(i), load_snapshot(directory, "redo-{:02d}".format(i)), states[i]) for i in range(1, count+1)]
    saved, reopened = load_snapshot(directory, "before-restart"), load_snapshot(directory, "after-restart")
    sequence += [("before-restart", saved, states[-1]), ("after-restart", reopened, states[-1])]
    prior_time = start
    for label, snapshot, expected in sequence:
        time = timestamp(snapshot["Utc"])
        require(time > prior_time, "Stale/out-of-order snapshot: " + label)
        prior_time = time
        same_state(snapshot_state(snapshot), expected, label)
        require(metadata(snapshot) == metadata(before) and snapshot["ConstraintXml"] == before["ConstraintXml"], label + ": metadata changed")
        require(snapshot["presentation"] == before["presentation"], label + ": presentation changed")
        if label != "after-restart":
            require(snapshot["ProcessId"] == before["ProcessId"], label + ": process changed before restart")
    require(saved["ProcessId"] != reopened["ProcessId"], "PowerPoint process was not restarted")
    return {"passed": True, "tolerancePt": EPSILON, "gestures": gestures, "undoSteps": count, "redoSteps": count,
            "checks": ["six kinds with bidirectional mouse-held updates", "constraints on every frame", "independent driver frame and release endpoints", "unrelated geometry",
                       "complete native undo and redo", "geometry and XML after process restart"]}


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--artifacts", type=Path, default=Path(__file__).resolve().parents[1] / "artifacts" / "layout-v2-validation")
    parser.add_argument("--output", type=Path, default=Path("acceptance.json"))
    args = parser.parse_args()
    try:
        result = verify(args.artifacts)
    except (ValueError, KeyError, TypeError, OSError, ET.ParseError) as error:
        result = {"passed": False, "error": str(error)}
    result.update(artifacts=str(args.artifacts.resolve()), checkedUtc=dt.datetime.now(dt.timezone.utc).isoformat())
    output = args.output if args.output.is_absolute() else args.artifacts / args.output
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(result, ensure_ascii=False, indent=2))
    sys.exit(0 if result["passed"] else 1)
