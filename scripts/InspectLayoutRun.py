"""Build expected native undo/redo states from actual recorded gesture endpoints."""
import copy
import json
from pathlib import Path
import sys
sys.path.insert(0, str(Path(__file__).resolve().parent))
import VerifyLayoutEvidence as evidence

directory = Path(__file__).resolve().parents[1] / 'artifacts' / 'layout-v2-validation'
before = evidence.load_snapshot(directory, 'before')
after = evidence.load_snapshot(directory, 'after')
start, stop = evidence.timestamp(before['Utc']), evidence.timestamp(after['Utc'])
state = evidence.snapshot_state(before)
slides = {slide['slideId']: slide for slide in before['slides']}
states, gestures, errors = [copy.deepcopy(state)], [], []
active = None
for line in (directory / 'gestures.log').read_text(encoding='utf-8-sig').splitlines():
    time, event, raw = line.split(' ', 2)
    if not start <= evidence.timestamp(time) <= stop:
        continue
    data = json.loads(raw)
    if event.endswith('error'):
        errors.append({'event': event, 'data': data})
    if event == 'gesture-start':
        active = {'start': data, 'held': 0}
    elif event == 'gesture-frame':
        evidence.require(active is not None, 'Missing gesture start')
        sid = active['start']['slideId']
        evidence.check_constraints(evidence.geometries(data['geometry']), slides[sid]['constraints'], event)
        active['held'] += int(data['mouseDown'])
    elif event == 'gesture-end':
        evidence.require(active is not None, 'Missing gesture start')
        sid = active['start']['slideId']
        evidence.same_geometry(evidence.geometries(data['before']), state[sid], 'Gesture baseline')
        if data['frames'] and not data['cancelled']:
            state[sid] = evidence.geometries(data['after'])
            states.append(copy.deepcopy(state))
            gestures.append({'id': data['gestureId'], 'slideId': sid, 'kind': slides[sid]['kind'],
                             'driverId': data['driverId'], 'mode': data['mode'],
                             'frames': data['frames'], 'heldFrames': active['held']})
        active = None
evidence.same_state(evidence.snapshot_state(after), state, 'Main run after')
result = {'gestureCount': len(gestures), 'gestures': gestures, 'states': states,
          'inputErrors': errors, 'inputChecksPassed': not errors}
(directory / 'history-plan.json').write_text(json.dumps(result, indent=2) + '\n', encoding='utf-8')
print(json.dumps({'gestureCount': len(gestures), 'gestures': gestures, 'errors': errors}, ensure_ascii=False, indent=2))
evidence.require(not errors, 'Input checks failed; diagnostic history plan is not an acceptance result')
