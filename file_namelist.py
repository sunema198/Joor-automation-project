import re
from collections import defaultdict
from pathlib import Path
from natsort import natsorted


IMG_EXTS = [".jpg", ".jpeg", ".png"]

MAX_FILE_SIZE = 4 * 1024 * 1024  # max file size

num_pattern = re.compile(r"^\d+[\(*\d+\)*]*")

"""
Returns linesheets in following format
{
    'code1': {
        'colors': {'black'},
        'color_photos': {
            'black': [....]
        }
        'photos': [....]
    },
    'code1': {
        'colors': {'blue'},
        'color_photos': {
            'blue': [....]
        }
        'photos': [....]
    },
}
"""

linesheets = defaultdict(list)
large_photos = []

pattern = re.compile(r'(-|_| |:)')


def match_pattern(name_parts, pattern_parts):
    if not name_parts or not pattern_parts:
        return False
    return len(name_parts) >= len(pattern_parts)


def get_code_or_color(slice, name_parts, pattern_parts):
    code_ = ''
    for code, p in zip(name_parts[slice], pattern_parts[slice]):
        if not code.isalnum():
            code_ += p
        else:
            code_ += code
    return code_


def get_linesheets(image_path=None):
    image_path = image_path or Path("560347")
    id_pattern = input(
        'id -> Style Id\n'
        'Ex: If name is 1419A_Black_02 then `id`, '
        '1033-011_Front -> id-id, 2458D SKIRT MARION DUSTY WHITE -> id\n'
        'Input id pattern: '
    ).strip().lower()
    parts = []
    if id_pattern:
        parts = pattern.split(id_pattern)
        code_slice = slice(0, len(parts))

    for item in natsorted(image_path.iterdir(), key=lambda x: x.stem.lower()):
        if item.is_file() and item.suffix.lower() in IMG_EXTS:
            if item.stat().st_size > MAX_FILE_SIZE:
                large_photos.append(item)
            else:
                name = item.stem
                name_parts = pattern.split(name)
                while not match_pattern(name_parts, parts):
                    name_pattern = input(
                        f'Pattern did not match for image {name} input new pattern or'
                        f'press s to skip: '
                    ).lower()
                    if name_pattern == 's':
                        code_slice = None
                        break
                    parts = pattern.split(name_pattern)
                    code_slice = slice(0, len(parts))

                if code_slice:
                    code = "".join(name_parts[code_slice])
                    linesheets[code].append(item)

    import pprint
    pprint.pprint(linesheets)
    return linesheets, large_photos


if __name__ == "__main__":
    get_linesheets()
