import re


class ResistorNormalizer:
    MULTIPLIERS = {
        "R": 1,
        "K": 1_000,
        "M": 1_000_000,
        "G": 1_000_000_000,
    }

    def normalize(self, value):
        if value is None:
            return None

        source = str(value).strip().replace(",", "")
        if not source:
            return ""

        milliohm = re.fullmatch(
            r"\s*(\d+(?:\.\d+)?)\s*m\s*(?:OHMS?|Ω)\s*",
            source,
        )
        if milliohm:
            return float(milliohm.group(1)) / 1_000

        normalized = source.upper().replace("Ω", "Ω")
        normalized = re.sub(r"OHMS?|Ω", "", normalized)
        normalized = re.sub(r"\s+", "", normalized)

        if normalized in {"0", "0R", "R0"}:
            return 0.0

        embedded_marker = re.fullmatch(r"(\d*)([RKMG])(\d+)", normalized)
        if embedded_marker:
            whole, marker, fraction = embedded_marker.groups()
            magnitude = float(f"{whole or '0'}.{fraction}")
            return magnitude * self.MULTIPLIERS[marker]

        suffix = re.fullmatch(r"(\d+(?:\.\d+)?)([RKMG]?)", normalized)
        if suffix:
            magnitude, marker = suffix.groups()
            return float(magnitude) * self.MULTIPLIERS.get(marker or "R", 1)

        return source