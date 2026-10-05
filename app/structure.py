from pathlib import PurePath

def structure_parts(relative_pdf):
    parts = PurePath(relative_pdf).parts
    if len(parts) >= 3:
        return parts[0], parts[1]
    if len(parts) == 2:
        return "", parts[0]
    return "", ""
