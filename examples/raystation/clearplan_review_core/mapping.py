"""Structure mapping: TG-263 template name + explicit aliases, never a fuzzy substitute.

Normalisation follows ClearPlan STRUCTURE_ALIASES.md: case, spaces and underscores are ignored; TG-263
semantic markers (~ - + ^ ! = /) are kept. Exactly one contoured ROI -> resolved; more than one ->
ambiguous; none -> unresolved. Ambiguous and unresolved rows are never evaluated.
"""

RESOLVED, AMBIGUOUS, UNRESOLVED = "resolved", "ambiguous", "unresolved"


def normalize(name):
    return "".join(ch for ch in (name or "").strip().lower() if ch not in (" ", "_", "\t"))


class MappingResult(object):
    def __init__(self, template_structure, status, roi=None, candidates=None, note=""):
        self.template_structure = template_structure
        self.status = status
        self.roi = roi
        self.candidates = list(candidates or [])
        self.note = note


def names_for(template_structure, table_aliases=None, alias_document=None):
    names = [template_structure] + list(table_aliases or [])
    key = normalize(template_structure)
    for canonical, aliases in (alias_document or {}).items():
        group = [canonical] + list(aliases)
        if key in [normalize(n) for n in group]:
            names.extend(group)
    seen, out = set(), []
    for name in names:
        n = normalize(name)
        if n and n not in seen:
            seen.add(n)
            out.append(name)
    return out


def resolve(template_structure, rois, table_aliases=None, alias_document=None):
    wanted = set(normalize(n) for n in names_for(template_structure, table_aliases, alias_document))
    matches = [roi for roi in rois if normalize(roi.name) in wanted]
    contoured = [roi for roi in matches if roi.has_contours is not False]
    if len(contoured) == 1:
        return MappingResult(template_structure, RESOLVED, contoured[0], contoured)
    if len(contoured) > 1:
        return MappingResult(template_structure, AMBIGUOUS, None, contoured,
                             "%d plan structures match the template name or its aliases." % len(contoured))
    if matches:
        return MappingResult(template_structure, UNRESOLVED, None, [], "Matching ROI has no contours.")
    return MappingResult(template_structure, UNRESOLVED, None, [], "No plan structure matches the template name or aliases.")
