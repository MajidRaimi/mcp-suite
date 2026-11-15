from __future__ import annotations

from collections import deque
from collections.abc import Callable

from office_mcp.shared.ax import AxApp, as_sequence


def find(
    app: AxApp,
    root: object,
    matches: Callable[[object], bool],
    max_depth: int = 30,
    prune: Callable[[object], bool] | None = None,
) -> object | None:
    found = find_all(app, root, matches, max_depth, prune, limit=1)
    return found[0] if found else None


def find_all(
    app: AxApp,
    root: object,
    matches: Callable[[object], bool],
    max_depth: int = 30,
    prune: Callable[[object], bool] | None = None,
    limit: int = 500,
) -> list[object]:
    found: list[object] = []
    queue: deque[tuple[object, int]] = deque([(root, 0)])
    while queue and len(found) < limit:
        element, depth = queue.popleft()
        if matches(element):
            found.append(element)
        if depth >= max_depth or (prune is not None and prune(element)):
            continue
        for child in as_sequence(app.attribute(element, "AXChildren")):
            if not app.same(child, element):
                queue.append((child, depth + 1))
    return found


def by_role_and_label(app: AxApp, role: str, label: str) -> Callable[[object], bool]:
    def matches(element: object) -> bool:
        if app.text(element, "AXRole") != role:
            return False
        return element_label(app, element) == label

    return matches


def element_label(app: AxApp, element: object) -> str:
    for name in ("AXTitle", "AXDescription", "AXHelp"):
        text = " ".join(app.text(element, name).split())
        if text:
            return text
    return ""
