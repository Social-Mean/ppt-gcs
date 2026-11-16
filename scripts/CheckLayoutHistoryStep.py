"""Compare one real native history snapshot with the recorded gesture sequence."""
import json
from pathlib import Path
import sys
sys.path.insert(0, str(Path(__file__).resolve().parent))
import VerifyLayoutEvidence as evidence

direction, step = sys.argv[1], int(sys.argv[2])
if direction not in ('undo', 'redo'):
    raise ValueError('Expected undo or redo')
directory = Path(__file__).resolve().parents[1] / 'artifacts' / 'layout-v2-validation'
plan = json.loads((directory / 'history-plan.json').read_text(encoding='utf-8'))
index = plan['gestureCount'] - step if direction == 'undo' else step
snapshot = evidence.load_snapshot(directory, '{}-{:02d}'.format(direction, step))
expected = {int(key): value for key, value in plan['states'][index].items()}
evidence.same_state(evidence.snapshot_state(snapshot), expected, direction + ' ' + str(step))
before = evidence.load_snapshot(directory, 'before')
evidence.require(evidence.metadata(snapshot) == evidence.metadata(before), 'Metadata changed')
evidence.require(snapshot['ConstraintXml'] == before['ConstraintXml'], 'XML changed')
evidence.require(snapshot['ProcessId'] == before['ProcessId'], 'PowerPoint process changed')
print('PASS {} {}/{}: all six slides and constraint XML match'.format(direction, step, plan['gestureCount']))
