from typing import Optional


def normalize_wbs(val) -> Optional[str]:
    if val is None:
        return None
    if isinstance(val, int):
        return str(val)
    if isinstance(val, float):
        return str(int(val)) if val == int(val) else str(val)
    return str(val).strip() or None


def get_wbs_level(wbs: str) -> int:
    return wbs.count('.') + 1


def find_parent(wbs: str, wbs_name_map: dict) -> Optional[str]:
    """Longest dot-split prefix that exists in the map."""
    parts = wbs.split('.')
    for i in range(len(parts) - 1, 0, -1):
        candidate = '.'.join(parts[:i])
        if candidate in wbs_name_map:
            return candidate
    return None


def get_ancestor_names(wbs: str, wbs_name_map: dict) -> set:
    names = set()
    current = wbs
    while True:
        parent = find_parent(current, wbs_name_map)
        if parent is None:
            break
        names.add(wbs_name_map[parent])
        current = parent
    return names


def is_leaf(wbs: str, all_wbs: set) -> bool:
    prefix = wbs + '.'
    return not any(w.startswith(prefix) for w in all_wbs)


_MGMT_PARENTS = {'项目监控活动', '需求管理活动', '策划活动'}


def classify_tasks(tasks: list[dict]) -> dict:
    """
    tasks: list of {id, wbs, name}
    Returns: {task_id: task_type}
    """
    wbs_name_map = {t['wbs']: t['name'] for t in tasks if t.get('wbs')}
    all_wbs = set(wbs_name_map.keys())

    result = {}
    for task in tasks:
        wbs = task.get('wbs')
        task_id = task['id']

        if not wbs:
            result[task_id] = 'other'
            continue

        if get_wbs_level(wbs) == 1:
            result[task_id] = 'phase'
            continue

        if not is_leaf(wbs, all_wbs):
            result[task_id] = 'other'
            continue

        ancestor_names = get_ancestor_names(wbs, wbs_name_map)
        parent_wbs = find_parent(wbs, wbs_name_map)
        parent_name = wbs_name_map.get(parent_wbs, '') if parent_wbs else ''

        if '配置管理活动' in ancestor_names:
            result[task_id] = 'cm'
        elif '质量保证活动' in ancestor_names:
            result[task_id] = 'qa'
        elif '评审活动' in ancestor_names:
            result[task_id] = 'review'
        elif parent_name in _MGMT_PARENTS:
            result[task_id] = 'mgmt'
        elif '软件工程活动' in ancestor_names:
            result[task_id] = 'engineering'
        else:
            result[task_id] = 'other'

    return result
