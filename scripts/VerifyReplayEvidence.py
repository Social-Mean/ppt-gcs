"""Verify programmatic layout replay/native Undo evidence, never physical mouse input."""
import argparse
import copy
import datetime as dt
import json
import importlib.util
import math
from pathlib import Path
import re
import sys
import xml.etree.ElementTree as ET

_spec = importlib.util.spec_from_file_location("layout_evidence", Path(__file__).with_name("VerifyLayoutEvidence.py"))
_helpers = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_helpers)
EPSILON, AXIS = _helpers.EPSILON, _helpers.AXIS
require, timestamp, close = _helpers.require, _helpers.timestamp, _helpers.close
geometries, same_geometry, same_state = _helpers.geometries, _helpers.same_geometry, _helpers.same_state
check_constraints, load_snapshot = _helpers.check_constraints, _helpers.load_snapshot
snapshot_state, metadata = _helpers.snapshot_state, _helpers.metadata


requested_geometry = _helpers.requested_geometry


def normalized_constraints(items):
    return [{"Id": item["Id"], "SlideId": item["SlideId"],
             "Kind": item.get("Kind", item.get("kind")), "ShapeIds": item["ShapeIds"]} for item in items]


def verify_gesture(begin, frames, release, end, slide, previous):
    driver, kind = begin["driverId"].upper(), slide["kind"]
    label = kind + " " + driver
    prior, final = geometries(begin["before"]), geometries(end["after"])
    same_geometry(prior, previous, label + " baseline")
    same_geometry(geometries(end["before"]), prior, label + " end baseline")
    require(begin.get("test") is True and end.get("test") is True, label + ": replay was not labelled test")
    require(end["cancelled"] is False and end["reason"] == "mouse-up", label + ": incomplete/cancelled replay")
    require(end["driverId"].upper() == driver and end["mode"] == begin["mode"], label + ": driver/mode changed")
    require(normalized_constraints(begin["constraints"]) == slide["constraints"], label + ": constraints changed")
    require(len(frames) == end["frames"] and len(frames) >= 2, label + ": frame count mismatch or insufficient samples")
    require(release is not None and release["geometryMatched"] is True and release.get("error") is None, label + ": release verification failed")
    require(release["driverId"].upper() == driver and release["mode"] == begin["mode"], label + ": release identity mismatch")
    require(release["frames"] == len(frames), label + ": release frame count mismatch")
    unaffected = {"SameX": ("Top",), "SameY": ("Left",), "SameWidth": ("Left", "Top", "Height"),
                  "SameHeight": ("Left", "Top", "Width"), "EqualHorizontalGap": ("Top",), "EqualVerticalGap": ("Left",)}[kind]
    for index, frame in enumerate(frames, 1):
        require(frame["frame"] == index and frame["driverId"].upper() == driver and frame["mode"] == begin["mode"], label + ": frame sequence mismatch")
        require(frame.get("input") == "owned-gesture-loop" and frame.get("mouseDown") is False, label + ": replay masquerades as physical mouse evidence")
        require(frame.get("pointerMessage") in ("move", "up"), label + ": unsupported pointer sample")
        require(math.isfinite(float(frame["residualPt"])) and 0 <= frame["residualPt"] <= EPSILON, label + ": invalid recorded residual")
        values = geometries(frame["geometry"])
        require(set(values) == set(prior), label + ": shape identities changed")
        check_constraints(values, slide["constraints"], label + " frame " + str(index))
        expected_driver = requested_geometry(begin, frame["cursor"])
        same_geometry({driver: values[driver]}, {driver: expected_driver}, label + " requested frame position")
        for identity, shape in values.items():
            if begin["mode"] == "Move":
                require(all(close(shape[f], prior[identity][f]) for f in ("Width", "Height")), label + ": move resized shape")
            if identity != driver:
                require(all(close(shape[f], prior[identity][f]) for f in unaffected), label + ": unrelated follower geometry changed")
    moves = [frame for frame in frames if frame["pointerMessage"] == "move"]
    require(moves and frames[-1]["pointerMessage"] == "up" and sum(f["pointerMessage"] == "up" for f in frames) == 1, label + ": exact release was not separately applied")
    require(frames[-1]["cursor"] == release["releasePoint"], label + ": final frame is not the release point")
    axis = "x" if kind in ("SameX", "SameWidth", "EqualHorizontalGap") else "y"
    correction = abs(release["releasePoint"][axis] - moves[-1]["cursor"][axis])
    require(9 <= correction <= 20, label + ": release did not test a 9–20 px correction beyond last move")
    require((release["releasePoint"][axis] - moves[-1]["cursor"][axis]) *
            (moves[-1]["cursor"][axis] - begin["cursor"][axis]) > 0, label + ": release did not continue beyond the motion sample")
    same_geometry(final, geometries(frames[-1]["geometry"]), label + " final frame")
    same_geometry({driver: final[driver]}, {driver: requested_geometry(begin, release["releasePoint"])}, label + " exact release endpoint")
    require(not close(prior[driver][AXIS[kind]], final[driver][AXIS[kind]]), label + ": constrained attribute did not change")
    check_constraints(final, slide["constraints"], label + " end")
    return final, {"gestureId": begin["gestureId"], "kind": kind, "driverId": driver,
                   "mode": begin["mode"], "frames": len(frames), "releaseCorrectionPixels": correction}


def verify(directory):
    before = load_snapshot(directory, "before")
    sequence = [("before", before)]
    for number in range(1, 13):
        for suffix in ("after", "undo", "redo"):
            name = "case-{:02d}-{}".format(number, suffix)
            sequence.append((name, load_snapshot(directory, name)))
    for action in ("undo", "redo"):
        for number in (1, 2):
            name = "chain-{}-{:02d}".format(action, number)
            sequence.append((name, load_snapshot(directory, name)))
    snapshots = dict(sequence)
    previous_time = None
    for name, snapshot in sequence:
        time = timestamp(snapshot["Utc"])
        require(previous_time is None or time > previous_time, "Stale/out-of-order snapshot: " + name)
        previous_time = time
        require(snapshot["ProcessId"] == before["ProcessId"] and snapshot["presentation"] == before["presentation"], name + ": process/document changed")
        require(metadata(snapshot) == metadata(before) and snapshot["ConstraintXml"] == before["ConstraintXml"], name + ": XML or shape metadata changed")
    start, stop = timestamp(before["Utc"]), timestamp(sequence[-1][1]["Utc"])
    state = snapshot_state(before)
    for slide in before["slides"]:
        check_constraints(state[slide["slideId"]], slide["constraints"], "Initial " + slide["kind"])
    states, summaries, active = [copy.deepcopy(state)], [], None
    log_path = directory / "gestures.log"
    require(log_path.is_file(), "Missing evidence: gestures.log")
    for line in log_path.read_text(encoding="utf-8-sig").splitlines():
        parts = line.split(" ", 2)
        if len(parts) != 3 or not re.match(r"^\d{4}-\d{2}-\d{2}T", parts[0]):
            continue
        time, event = timestamp(parts[0]), parts[1]
        if not start <= time <= stop:
            continue
        require(not event.endswith("error") and "constraint-rejected" not in event, "Runtime failure during replay: " + line)
        require(not event.startswith("gesture-"), "Physical gesture events mixed into programmatic replay evidence")
        if event not in ("test-gesture-start", "test-gesture-frame", "test-gesture-release", "test-gesture-end"):
            continue
        data = json.loads(parts[2])
        if event == "test-gesture-start":
            require(active is None and len(summaries) < 12, "Overlapping or extra replay gestures")
            number = len(summaries) + 1
            lower = before if number == 1 else snapshots["case-{:02d}-redo".format(number-1)]
            require(time > timestamp(lower["Utc"]), "Replay preceded its baseline snapshot")
            active = {"begin": data, "frames": [], "release": None, "time": time}
            continue
        require(active is not None and active["begin"]["gestureId"] == data["gestureId"], "Unmatched replay event")
        require(time >= active["time"], "Replay event time moved backwards")
        active["time"] = time
        if event == "test-gesture-frame":
            require(active["release"] is None, "Frame appears after release record")
            active["frames"].append(data)
        elif event == "test-gesture-release":
            require(active["release"] is None, "Duplicate release record")
            active["release"] = data
        else:
            number = len(summaries) + 1
            slide = before["slides"][(number-1)//2]
            sid = slide["slideId"]
            expected_mode = "ResizeSE" if slide["kind"] == "SameWidth" else "ResizeS" if slide["kind"] == "SameHeight" else "Move"
            members = sorted(slide["geometry"], key=lambda item: item["Name"])
            expected_driver = members[0 if number % 2 else -1]["Id"].upper()
            require(active["begin"]["slideId"] == sid and active["begin"]["mode"] == expected_mode and active["begin"]["driverId"].upper() == expected_driver, "Replay case order/mode/driver differs from plan")
            final, summary = verify_gesture(active["begin"], active["frames"], active["release"], data, slide, state[sid])
            require(time < timestamp(snapshots["case-{:02d}-after".format(number)]["Utc"]), "Replay ended after its after snapshot")
            state[sid] = final
            states.append(copy.deepcopy(state)); summaries.append(summary)
            for suffix, expected in (("after", states[-1]), ("undo", states[-2]), ("redo", states[-1])):
                name = "case-{:02d}-{}".format(number, suffix)
                same_state(snapshot_state(snapshots[name]), expected, name)
            active = None
    require(active is None and len(summaries) == 12, "Expected twelve complete replay gestures")
    for name, expected in (("chain-undo-01", states[11]), ("chain-undo-02", states[10]),
                           ("chain-redo-01", states[11]), ("chain-redo-02", states[12])):
        same_state(snapshot_state(snapshots[name]), expected, name)
    return {"passed": True, "evidenceType": "programmatic-replay-native-undo", "provesPhysicalMouseInput": False,
            "tolerancePt": EPSILON, "cases": summaries, "nativeUndoSteps": 14, "nativeRedoSteps": 14,
            "checks": ["six kinds and two drivers", "constraints on every programmatic frame", "exact release target independent of solver",
                       "release correction beyond last move", "single native undo and redo per gesture", "two-step native history chain", "all-slide geometry and XML preservation"]}


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--artifacts", type=Path, default=Path(__file__).resolve().parents[1] / "artifacts" / "layout-v2-replay")
    parser.add_argument("--output", type=Path, default=Path("report.json"))
    args = parser.parse_args()
    try:
        result = verify(args.artifacts)
    except (ValueError, KeyError, TypeError, OSError, IndexError, ET.ParseError) as error:
        result = {"passed": False, "evidenceType": "programmatic-replay-native-undo", "provesPhysicalMouseInput": False, "error": str(error)}
    result.update(artifacts=str(args.artifacts.resolve()), checkedUtc=dt.datetime.now(dt.timezone.utc).isoformat())
    output = args.output if args.output.is_absolute() else args.artifacts / args.output
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(result, ensure_ascii=False, indent=2))
    sys.exit(0 if result["passed"] else 1)
