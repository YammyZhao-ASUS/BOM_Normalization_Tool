class ComponentDetector:


    def detect(self, text):

        if text is None:
            return "UNKNOWN"


        text=str(text).upper()


        if "MLCC" in text:
            return "C"


        if "RES" in text:
            return "R"


        if "IND" in text:
            return "L"


        return "UNKNOWN"