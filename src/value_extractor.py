import re


class ValueExtractor:

    def extract(self, component_name, component_type="UNKNOWN"):
        """Return the raw value token appropriate for the component type."""
        if component_type == "C":
            return self.extract_value(component_name)

        if component_type == "R":
            return self.extract_resistor_value(component_name)

        return ""


    @staticmethod
    def normalize_cap_value(value):

        if value is None:
            return ""

        value = str(value).upper().strip()

        value = value.replace(" ", "")
        value = value.replace("μ", "U")

        # --------------------------------
        # uF
        # --------------------------------

        m = re.search(r'([\d\.]+)UF', value)
        if m:
            return f"{float(m.group(1)):g}uF"

        # --------------------------------
        # nF
        # --------------------------------

        m = re.search(r'([\d\.]+)NF', value)
        if m:
            return f"{float(m.group(1)):g}nF"

        # --------------------------------
        # pF
        # --------------------------------

        m = re.search(r'([\d\.]+)PF', value)
        if m:
            return f"{float(m.group(1)):g}pF"

        # --------------------------------
        # F
        # --------------------------------

        m = re.search(r'([\d\.]+)F', value)
        if m:
            return f"{float(m.group(1)):g}F"

        return ""


    @staticmethod
    def extract_value(component_name):
        if component_name is None:
            return ""

        s = str(component_name).upper()

        s = s.replace("μ", "U").replace("µ", "U")

        patterns = [
            r"(?<![\w.])((?:\d+(?:\.\d*)?|\.\d+)(?:MF|UF|NF|PF|U|N|P|F))(?![A-Z])",
            r"\b(\d{3})\b",
        ]

        for pattern in patterns:
            match = re.search(pattern, s)
            if match:
                return match.group(1)

        return ""


    @staticmethod
    def extract_voltage(component_name):
        if component_name is None:
            return ""

        s = str(component_name).upper()

        embedded_decimal = re.search(r"(?<!\d)(\d+)V(\d+)(?:DC)?(?!\d)", s)
        if embedded_decimal:
            return f"{float(f'{embedded_decimal.group(1)}.{embedded_decimal.group(2)}'):g}V"

        m = re.search(r"(?<![\w.])((?:\d+(?:\.\d*)?|\.\d+)\s*V(?:DC)?)(?![A-Z])", s)

        if m:
            magnitude = re.search(r"\d+(?:\.\d*)?|\.\d+", m.group(1)).group(0)
            return f"{float(magnitude):g}V"

        return ""


    @staticmethod
    def extract_dielectric(component_name):
        if component_name is None:
            return ""

        s = str(component_name).upper()

        for d in [
            "X5R",
            "X5S",
            "X6S",
            "X7S",
            "X7R",
            "X8R",
            "Y5V",
            "Z5U",
            "NP0",
            "NPO",
            "COG",
            "C0G",
        ]:
            if d in s:
                if d in {"COG", "C0G", "NPO"}:
                    return "NP0"
                return d

        return ""


    @staticmethod
    def extract_size(component_name):
        if component_name is None:
            return ""

        s = str(component_name).upper()

        m = re.search(
            r"(?<!\d)(01005|0201|0402|0603|0805|1206|1210|1808|1812|2010|2512|2220)(?!\d)",
            s,
        )

        if m:

            return m.group(1)

        return ""


    @staticmethod
    def extract_resistor_value(component_name):
        if component_name is None:
            return ""

        source = str(component_name).replace("Ω", "Ω")
        milliohm = re.search(
            r"(?<![\w.])((?:\d+(?:\.\d*)?|\.\d+)\s*m\s*(?:OHMS?|Ω))(?!\w)",
            source,
        )
        if milliohm:
            return milliohm.group(1)

        s = source.upper()

        patterns = [
            r"(?<![\w.])(?:\d+(?:\.\d+)?\s*[RKMG]?|[RKM]\d+)\s*(?:OHMS?|Ω)(?!\w)",
            r"(?<!\w)(?:\d*[RKM]\d+|\d+(?:\.\d+)?[RKMG])(?!\w)",
        ]

        for pattern in patterns:
            match = re.search(pattern, s)
            if match:
                return match.group(0).strip()

        return ""

    @staticmethod
    def extract_tolerance(component_name):
        if component_name is None:
            return ""

        match = re.search(r"(?:±|\+/-)?\s*(\d+(?:\.\d+)?)\s*%", str(component_name))
        return f"{float(match.group(1)):g}%" if match else ""

    @staticmethod
    def extract_power_rating(component_name):
        if component_name is None:
            return ""

        source = str(component_name).upper()
        fraction = re.search(r"(?<!\d)(\d+)\s*/\s*(\d+)\s*W(?![A-Z])", source)
        if fraction and float(fraction.group(2)):
            watts = float(fraction.group(1)) / float(fraction.group(2))
            return f"{watts:g}W"

        milliwatts = re.search(
            r"(?<![\w.])(\d+(?:\.\d+)?)\s*MW(?![A-Z])",
            source,
        )
        if milliwatts:
            return f"{float(milliwatts.group(1)) / 1_000:g}W"

        watts = re.search(r"(?<![\w.])(\d+(?:\.\d+)?)\s*W(?![A-Z])", source)
        return f"{float(watts.group(1)):g}W" if watts else ""

    def extract_material(self, component_name, component_type):
        if component_type == "C":
            return self.extract_dielectric(component_name)

        if component_type != "R" or component_name is None:
            return ""

        source = str(component_name).upper()
        aliases = (
            ("THIN FILM", ("THIN FILM", "THINFILM")),
            ("THICK FILM", ("THICK FILM", "THICKFILM")),
            ("METAL FILM", ("METAL FILM", "METALFILM")),
            ("METAL FOIL", ("METAL FOIL", "FOIL")),
            ("WIREWOUND", ("WIREWOUND", "WIRE WOUND")),
            ("CARBON FILM", ("CARBON FILM", "CARBON")),
        )
        for normalized, values in aliases:
            if any(value in source for value in values):
                return normalized
        return ""