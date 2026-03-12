import base64
import datetime as _dt
import hashlib
import io
import random
import time
from typing import Any, Tuple, Optional

def strip_yaml_comment(line: str) -> str:
    in_single = False
    in_double = False
    for i, ch in enumerate(line):
        if ch == "'" and not in_double:
            in_single = not in_single
        elif ch == '"' and not in_single:
            in_double = not in_double
        elif ch == "#" and not in_single and not in_double:
            return line[:i]
    return line

def parse_yaml_scalar(value: str) -> Any:
    value = value.strip()
    if value == "":
        return ""
    if value.lower() in {"true", "false"}:
        return value.lower() == "true"
    if value.lower() in {"null", "none", "~"}:
        return None
    if (value.startswith('"') and value.endswith('"')) or (value.startswith("'") and value.endswith("'")):
        return value[1:-1]
    import re
    if re.fullmatch(r"-?\d+", value):
        try:
            return int(value)
        except Exception:
            return value
    if re.fullmatch(r"-?\d+\.\d+", value):
        try:
            return float(value)
        except Exception:
            return value
    return value

def now_iso() -> str:
    return _dt.datetime.now().isoformat(timespec="seconds")

def today_key() -> str:
    return _dt.date.today().isoformat()

def sleep_with_jitter(base_seconds: float, jitter_seconds: float) -> None:
    if base_seconds < 0:
        base_seconds = 0
    if jitter_seconds < 0:
        jitter_seconds = 0
    time.sleep(base_seconds + random.random() * jitter_seconds)

def rect_to_bbox(rect: Any) -> Optional[Tuple[int, int, int, int]]:
    if rect is None:
        return None
    try:
        left = int(rect.left)
        top = int(rect.top)
        right = int(rect.right)
        bottom = int(rect.bottom)
    except Exception:
        try:
            if isinstance(rect, (tuple, list)) and len(rect) == 4:
                left, top, right, bottom = [int(v) for v in rect]
            else:
                return None
        except Exception:
            return None
    if right <= left or bottom <= top:
        return None
    return (left, top, right, bottom)

def image_to_base64_png(img: Any) -> str:
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return base64.b64encode(buf.getvalue()).decode("ascii")

def sha1_text(text: str) -> str:
    return hashlib.sha1(text.encode("utf-8", errors="ignore")).hexdigest()
