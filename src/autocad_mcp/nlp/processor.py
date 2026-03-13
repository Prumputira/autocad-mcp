"""
NLP Processor for AutoCAD MCP.

Parses natural language commands and extracts parameters for CAD operations.
Supports shape drawing, layer management, and entity manipulation.
"""

import re
import logging
from typing import Any
from dataclasses import dataclass

logger = logging.getLogger(__name__)


@dataclass
class ParsedCommand:
    """Result of parsing a natural language command."""
    operation: str          # e.g. "create_line", "create_circle", "layer_create"
    parameters: dict[str, Any]
    confidence: float       # 0.0-1.0
    raw_text: str


class NLPProcessor:
    """Process natural language commands for CAD operations."""

    COMMAND_KEYWORDS = {
        "draw": ["draw", "create", "make", "add", "rita", "skapa"],
        "erase": ["erase", "delete", "remove", "clear", "radera", "ta bort"],
        "move": ["move", "translate", "shift", "flytta"],
        "rotate": ["rotate", "turn", "spin", "rotera"],
        "scale": ["scale", "resize", "enlarge", "shrink", "skala"],
        "dimension": ["dimension", "annotate", "measure", "dimensionera", "mäta"],
        "layer": ["layer", "lager"],
        "rename": ["rename", "change name", "byt namn"],
        "turn_on": ["turn on", "show", "unhide", "enable", "visa"],
        "turn_off": ["turn off", "hide", "disable", "dölj", "göm"],
    }

    SHAPE_KEYWORDS = {
        "line": ["line", "segment", "linje"],
        "circle": ["circle", "round", "cirkel"],
        "arc": ["arc", "curve", "båge"],
        "rectangle": ["rectangle", "rect", "square", "rektangel", "fyrkant"],
        "polyline": ["polyline", "polygon", "path", "polylinje"],
        "ellipse": ["ellipse", "oval", "ellips"],
        "text": ["text", "label", "note", "etikett"],
        "hatch": ["hatch", "fill", "pattern", "skraffering"],
    }

    COLOR_NAMES = {
        "black": 0, "red": 1, "yellow": 2, "green": 3, "cyan": 4,
        "blue": 5, "magenta": 6, "white": 7, "gray": 8, "grey": 8,
        "orange": 30,
        # Swedish
        "svart": 0, "röd": 1, "gul": 2, "grön": 3, "blå": 5, "vit": 7, "grå": 8,
    }

    NUMBER_PATTERN = r"[+-]?(?:\d+\.?\d*|\d*\.\d+)(?:[eE][+-]?\d+)?"
    COORD_PATTERN = rf"\(?[\s]*(?:x\s*=\s*)?({NUMBER_PATTERN})\s*[,;]\s*(?:y\s*=\s*)?({NUMBER_PATTERN})(?:\s*[,;]\s*(?:z\s*=\s*)?({NUMBER_PATTERN}))?\s*\)?"

    def __init__(self, strict_mode: bool = False):
        self.strict_mode = strict_mode
        self.coord_regex = re.compile(self.COORD_PATTERN)
        self.number_regex = re.compile(self.NUMBER_PATTERN)

    def _keyword_in_text(self, text: str, keyword: str) -> bool:
        return re.search(rf"\b{re.escape(keyword)}\b", text) is not None

    def parse_command(self, text: str) -> ParsedCommand:
        """Parse natural language command into structured operation."""
        text_lower = text.lower()
        command_type = self._identify_command_type(text_lower)

        if not command_type:
            shape_type = self._identify_shape_type(text_lower)
            if shape_type:
                command_type = "draw"
            else:
                return ParsedCommand(
                    operation="unknown",
                    parameters={"raw": text},
                    confidence=0.0,
                    raw_text=text,
                )

        if command_type in ("layer", "rename", "turn_on", "turn_off"):
            return self._parse_layer_command(command_type, text_lower, text)

        if command_type == "draw":
            shape_type = self._identify_shape_type(text_lower)
            if not shape_type:
                return ParsedCommand(
                    operation="unknown", parameters={"raw": text},
                    confidence=0.1, raw_text=text,
                )
            params = self._extract_shape_params(shape_type, text_lower, text)
            confidence = self._calculate_confidence(text_lower, shape_type, params)
            # Map to backend method names
            op_map = {
                "line": "create_line", "circle": "create_circle",
                "arc": "create_arc", "rectangle": "create_rectangle",
                "polyline": "create_polyline", "ellipse": "create_ellipse",
                "text": "create_text", "hatch": "create_hatch",
            }
            return ParsedCommand(
                operation=op_map.get(shape_type, f"create_{shape_type}"),
                parameters=params,
                confidence=confidence,
                raw_text=text,
            )

        if command_type == "erase":
            return ParsedCommand(
                operation="entity_erase",
                parameters={"raw": text},
                confidence=0.5,
                raw_text=text,
            )

        if command_type == "move":
            coords = list(self.coord_regex.finditer(text_lower))
            params = {}
            if len(coords) >= 2:
                params["dx"] = float(coords[1].group(1)) - float(coords[0].group(1))
                params["dy"] = float(coords[1].group(2)) - float(coords[0].group(2))
            return ParsedCommand(
                operation="entity_move", parameters=params,
                confidence=0.5 if params else 0.2, raw_text=text,
            )

        return ParsedCommand(
            operation="unknown", parameters={"raw": text},
            confidence=0.1, raw_text=text,
        )

    def _identify_command_type(self, text_lower: str) -> str | None:
        if re.search(r"\b(create|make|add|skapa)\s+.*?\b(layer|lager)\b", text_lower):
            return "layer"
        if "layer" in text_lower or "lager" in text_lower:
            if any(w in text_lower for w in ("rename", "byt namn")):
                return "rename"
            if re.search(r"\b(turn\s+on|visa|show|unhide)\b", text_lower):
                return "turn_on"
            if re.search(r"\b(turn\s+off|dölj|göm|hide)\b", text_lower):
                return "turn_off"
            if re.search(r"\b(delete|remove|erase|radera)\s+(layer|lager)\b", text_lower):
                return "layer"
            if re.search(r"\b(list|show)\s+(all\s+)?(layers?|lager)\b", text_lower):
                return "layer"
        for cmd_type in ("draw",):
            keywords = self.COMMAND_KEYWORDS.get(cmd_type, [])
            if any(kw in text_lower for kw in keywords):
                return cmd_type
        for cmd_type, keywords in self.COMMAND_KEYWORDS.items():
            if cmd_type not in ("layer", "rename", "turn_on", "turn_off", "draw"):
                if any(kw in text_lower for kw in keywords):
                    return cmd_type
        return None

    def _identify_shape_type(self, text_lower: str) -> str | None:
        sorted_shapes = sorted(
            self.SHAPE_KEYWORDS.items(),
            key=lambda x: max(len(kw) for kw in x[1]),
            reverse=True,
        )
        for shape_type, keywords in sorted_shapes:
            if any(self._keyword_in_text(text_lower, kw) for kw in keywords):
                return shape_type
        return None

    def _extract_shape_params(self, shape: str, text_lower: str, original: str) -> dict:
        params: dict[str, Any] = {}
        params["color"] = self._extract_color(text_lower)
        params["layer"] = self._extract_layer(original)

        if shape == "line":
            matches = list(self.coord_regex.finditer(text_lower))
            if len(matches) >= 2:
                params["x1"] = float(matches[0].group(1))
                params["y1"] = float(matches[0].group(2))
                params["x2"] = float(matches[1].group(1))
                params["y2"] = float(matches[1].group(2))
            elif not self.strict_mode:
                params.update({"x1": 0, "y1": 0, "x2": 100, "y2": 100})

        elif shape == "circle":
            match = self.coord_regex.search(text_lower)
            if match:
                params["cx"] = float(match.group(1))
                params["cy"] = float(match.group(2))
            else:
                params["cx"], params["cy"] = 0.0, 0.0
            r_match = re.search(r"(?:radius|r)\s*(?:=|of|:)?\s*(\d+\.?\d*)", text_lower)
            params["radius"] = float(r_match.group(1)) if r_match else 50.0

        elif shape == "rectangle":
            matches = list(self.coord_regex.finditer(text_lower))
            if len(matches) >= 2:
                params["x1"] = float(matches[0].group(1))
                params["y1"] = float(matches[0].group(2))
                params["x2"] = float(matches[1].group(1))
                params["y2"] = float(matches[1].group(2))
            elif not self.strict_mode:
                params.update({"x1": 0, "y1": 0, "x2": 100, "y2": 100})

        elif shape == "polyline":
            matches = list(self.coord_regex.finditer(text_lower))
            params["points"] = [[float(m.group(1)), float(m.group(2))] for m in matches]
            params["closed"] = "closed" in text_lower or "sluten" in text_lower

        elif shape == "arc":
            match = self.coord_regex.search(text_lower)
            if match:
                params["cx"] = float(match.group(1))
                params["cy"] = float(match.group(2))
            else:
                params["cx"], params["cy"] = 0.0, 0.0
            r_match = re.search(r"(?:radius|r)\s*(?:=|of|:)?\s*(\d+\.?\d*)", text_lower)
            params["radius"] = float(r_match.group(1)) if r_match else 50.0
            sa = re.search(r"start[_\s]?angle\s*(?:=|:)?\s*(\d+\.?\d*)", text_lower)
            ea = re.search(r"end[_\s]?angle\s*(?:=|:)?\s*(\d+\.?\d*)", text_lower)
            params["start_angle"] = float(sa.group(1)) if sa else 0.0
            params["end_angle"] = float(ea.group(1)) if ea else 90.0

        elif shape == "text":
            match = self.coord_regex.search(text_lower)
            if match:
                params["x"] = float(match.group(1))
                params["y"] = float(match.group(2))
            else:
                params["x"], params["y"] = 0.0, 0.0
            t_match = re.search(r'(?:text|label|content)\s*["\']([^"\']+)["\']', original)
            params["text"] = t_match.group(1) if t_match else "Text"
            h_match = re.search(r"(?:height|size)\s*(?:=|of|:)?\s*(\d+\.?\d*)", text_lower)
            params["height"] = float(h_match.group(1)) if h_match else 2.5

        elif shape == "ellipse":
            match = self.coord_regex.search(text_lower)
            if match:
                params["cx"] = float(match.group(1))
                params["cy"] = float(match.group(2))
            else:
                params["cx"], params["cy"] = 0.0, 0.0
            params["major_x"], params["major_y"] = 50.0, 0.0
            ratio_match = re.search(r"(?:ratio|minor)\s*(?:=|:)?\s*(\d+\.?\d*)", text_lower)
            params["ratio"] = float(ratio_match.group(1)) if ratio_match else 0.5

        return params

    def _extract_color(self, text: str) -> str | None:
        for name in self.COLOR_NAMES:
            if name in text:
                return name
        return None

    def _extract_layer(self, text: str) -> str | None:
        match = re.search(r'(?:layer|lager|on)\s+["\']([^"\']+)["\']', text, re.IGNORECASE)
        if match:
            return match.group(1)
        match = re.search(r"(?:layer|lager)\s+([a-zA-Z0-9_\-]+)", text, re.IGNORECASE)
        if match:
            return match.group(1)
        return None

    def _calculate_confidence(self, text: str, shape: str, params: dict) -> float:
        score = 0.0
        if any(self._keyword_in_text(text, kw) for kw in ("draw", "create", "rita", "skapa")):
            score += 0.15
        if self._keyword_in_text(text, shape):
            score += 0.30
        has_coords = any(k in params for k in ("x1", "cx", "x", "points"))
        if has_coords:
            score += 0.35
        if params.get("radius") or params.get("height"):
            score += 0.10
        if params.get("color"):
            score += 0.05
        if params.get("layer"):
            score += 0.05
        return min(1.0, score)

    def _parse_layer_command(self, cmd_type: str, text_lower: str, original: str) -> ParsedCommand:
        params: dict[str, Any] = {}

        if cmd_type == "rename":
            quoted = re.findall(r'["\']([^"\']+)["\']', text_lower)
            if len(quoted) >= 2:
                params["old_name"], params["new_name"] = quoted[0], quoted[1]
            else:
                m = re.search(r"(?:layer|lager)\s+(\S+)\s+(?:to|till)\s+(\S+)", text_lower)
                if m:
                    params["old_name"], params["new_name"] = m.group(1), m.group(2)
            return ParsedCommand(
                operation="layer_rename", parameters=params,
                confidence=0.9 if params else 0.3, raw_text=original,
            )

        layer_name = self._extract_layer_name(text_lower, original)
        if cmd_type == "turn_on":
            return ParsedCommand(
                operation="layer_thaw", parameters={"name": layer_name or "0"},
                confidence=0.9 if layer_name else 0.3, raw_text=original,
            )
        if cmd_type == "turn_off":
            return ParsedCommand(
                operation="layer_freeze", parameters={"name": layer_name or "0"},
                confidence=0.9 if layer_name else 0.3, raw_text=original,
            )
        # Generic layer command
        if any(kw in text_lower for kw in ("create", "make", "add", "skapa")):
            params["name"] = layer_name or "NewLayer"
            color = self._extract_color(text_lower)
            if color:
                params["color"] = color
            return ParsedCommand(
                operation="layer_create", parameters=params,
                confidence=0.9, raw_text=original,
            )
        if any(kw in text_lower for kw in ("delete", "remove", "radera")):
            return ParsedCommand(
                operation="layer_delete", parameters={"name": layer_name or ""},
                confidence=0.8 if layer_name else 0.3, raw_text=original,
            )
        if any(kw in text_lower for kw in ("list", "show", "alla")):
            return ParsedCommand(
                operation="layer_list", parameters={},
                confidence=0.95, raw_text=original,
            )
        return ParsedCommand(
            operation="layer_list", parameters={},
            confidence=0.5, raw_text=original,
        )

    def _extract_layer_name(self, text_lower: str, original: str) -> str | None:
        quoted = re.search(r'["\']([^"\']+)["\']', original)
        if quoted:
            return quoted.group(1)
        m = re.search(r"(?:layer|lager)\s+([a-zA-Z0-9_\-]+)", original, re.IGNORECASE)
        if m:
            return m.group(1)
        return None
