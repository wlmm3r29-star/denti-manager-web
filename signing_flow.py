"""Session-only signature stages. No patient data is written to disk."""
from datetime import datetime
from zoneinfo import ZoneInfo

ROLES = ('paciente', 'prestador')


def new_flow(identity):
    return dict(identity=identity, role='paciente', generation=0, accepted={},
                target=None, last_event=None, phase='DOCUMENTO_CARGADO')


def repeat_stage(flow):
    role = flow['role']
    flow['accepted'].pop(role, None)
    flow['repeat_role'] = role
    flow['generation'] += 1
    flow['target'] = None
    flow['phase'] = 'ESPERANDO_FIRMA_' + role.upper()


def consume_event(flow, event, validate_png):
    """Accept only an event bound to the current document, role and rectangle."""
    target = flow.get('target')
    if not isinstance(event, dict) or not target:
        return
    if event.get('context') != target['context'] or event.get('role') != flow['role']:
        return
    if not event.get('id') or event['id'] == flow['last_event']:
        return
    kind = event.get('kind')
    role = flow['role']
    if kind == 'accepted':
        if role in flow['accepted']:
            return
        png = validate_png(event.get('png'))
        flow['accepted'][role] = dict(target, png=png,
            signed_at=datetime.now(ZoneInfo('America/Bogota')))
        flow['phase'] = 'FIRMA_' + role.upper() + '_ACEPTADA'
    elif kind == 'repeat':
        repeat_stage(flow)
    elif kind == 'captured' and role not in flow['accepted']:
        flow['phase'] = 'FIRMA_' + role.upper() + '_CAPTURADA'
    elif kind == 'ready' and role not in flow['accepted']:
        flow['phase'] = 'ESPERANDO_FIRMA_' + role.upper()
    else:
        return
    flow['last_event'] = event['id']


def ready_to_save(flow):
    return bool(flow['accepted'].get('paciente') and
                flow['role'] in flow['accepted'] and
                not flow['phase'].endswith('_CAPTURADA'))
