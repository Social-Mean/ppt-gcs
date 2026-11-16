"""Read-only audit that each recorded physical gesture is one native undo step."""
import datetime as dt
import json
from pathlib import Path

directory = Path(__file__).resolve().parents[1] / "artifacts" / "v3-interaction"


def read(name):
    return json.loads((directory / (name + ".json")).read_text(encoding="utf-8-sig"))


def geometry(state):
    return {s["Id"]: s for page in state["Snapshot"]["slides"] for s in page["geometry"]}


def matches(actual, expected):
    return all(s["Id"] in actual and all(abs(actual[s["Id"]][k] - s[k]) <= .01
               for k in ("Left", "Top", "Width", "Height")) for s in expected)


before, after, acceptance = read("before"), read("after"), read("acceptance")
assert acceptance["Passed"]
start, stop = dt.datetime.fromisoformat(before["Utc"]), dt.datetime.fromisoformat(after["Utc"])
gestures, commands = [], []
for line in (directory / "events.log").read_text(encoding="utf-8-sig").splitlines():
    parts = line.split(" ", 2)
    if len(parts) != 3 or not start <= dt.datetime.fromisoformat(parts[0]) <= stop:
        continue
    data = json.loads(parts[2])
    if parts[1] == "gesture-end" and data["frames"] and not data["cancelled"]:
        assert data["test"] is False
        gestures.append(data)
    elif parts[1] == "constraint-command":
        commands.append(data)

states = [after] + [read("undo-{:02}".format(i)) for i in range(1, acceptance["NativeUndoSteps"] + 1)]
claimed = set()
gesture_steps = []
for gesture in gestures:
    found = [i for i in range(1, len(states))
             if states[i-1]["ConstraintXml"] == states[i]["ConstraintXml"]
             and matches(geometry(states[i-1]), gesture["after"])
             and matches(geometry(states[i]), gesture["before"])]
    assert len(found) == 1, ("Physical gesture split or missing in native history", gesture["gestureId"], found)
    assert found[0] not in claimed
    claimed.add(found[0])
    gesture_steps.append({"gestureId": gesture["gestureId"], "mode": gesture["mode"], "nativeUndoStep": found[0]})

metadata_steps = [i for i in range(1, len(states)) if states[i-1]["ConstraintXml"] != states[i]["ConstraintXml"]]
assert len(metadata_steps) == len(commands), "Constraint commands were split or lost"
claimed.update(metadata_steps)
unconstrained_steps = []
for i in range(1, len(states)):
    if i in claimed:
        continue
    old, new = geometry(states[i-1]), geometry(states[i])
    assert set(old) == set(new), "Unexpected shape identity change outside a constraint transaction"
    changed = [key for key in old if any(abs(old[key][k] - new[key][k]) > .01
               for k in ("Left", "Top", "Width", "Height"))]
    active = {key for page in states[i-1]["Snapshot"]["slides"] for relation in page["constraints"]
              if relation["Enabled"] for key in relation["ShapeIds"]}
    assert changed and not active.intersection(changed), "An unexplained native edit changed constrained geometry"
    unconstrained_steps.append({"nativeUndoStep": i, "shapeIds": changed})

result = {"Passed": True, "WholePhysicalGestures": gesture_steps,
          "MetadataTransactionSteps": metadata_steps, "AdditionalUnconstrainedNativeEdits": unconstrained_steps,
          "Checks": ["one complete native undo step per physical gesture", "one native step per metadata command",
                     "extra native edits only affect unconstrained shapes"]}
(directory / "history-audit.json").write_text(json.dumps(result, indent=2) + "\n", encoding="utf-8")
print("PASS: {} complete physical gestures, {} metadata commands, {} unconstrained native edit(s)".format(
    len(gestures), len(commands), len(unconstrained_steps)))
