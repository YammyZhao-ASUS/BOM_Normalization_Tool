import re


class CapacitorNormalizer:

    UNIT_TO_PF = {
        "F": 1_000_000_000_000,
        "MF": 1_000_000_000,
        "UF": 1_000_000,
        "U": 1_000_000,
        "NF": 1_000,
        "N": 1_000,
        "PF": 1,
        "P": 1,
    }

    def normalize(self, value):
        if value is None:
            return ""

        value = str(value).upper().strip()
        value = value.replace(" ", "").replace("μ", "U").replace("µ", "U")

        if len(value) == 3 and value.isdigit():
            picofarads = int(value[:2]) * (10 ** int(value[2]))
            return self.format_pf(picofarads)

        match = re.fullmatch(r"((?:\d+(?:\.\d*)?)|(?:\.\d+))(MF|UF|NF|PF|U|N|P|F)", value)
        if not match:
            return value

        magnitude, unit = match.groups()
        return self.format_pf(float(magnitude) * self.UNIT_TO_PF[unit])

    @staticmethod
    def format_pf(picofarads):
        if picofarads >= 1_000_000_000_000:
            return f"{picofarads / 1_000_000_000_000:g}F"

        if picofarads >= 1_000_000_000:
            return f"{picofarads / 1_000_000_000:g}mF"

        if picofarads >= 1_000_000:
            return f"{picofarads / 1_000_000:g}uF"

        if picofarads >= 1_000:
            return f"{picofarads / 1_000:g}nF"

        return f"{picofarads:g}pF"

    def format_nf(self, nanofarads):
        return self.format_pf(nanofarads * 1_000)