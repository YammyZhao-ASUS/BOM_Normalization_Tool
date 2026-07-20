class BOMColumnMapper:


    def __init__(self, columns):

        self.columns=list(columns)



    def find_reference(self):

        candidates=[
            "Part Reference",
            "Reference",
            "RefDes"
        ]

        for col in candidates:

            if col in self.columns:
                return col

        return None



    def find_description(self):

        candidates=[
            "Component_Name",
            "Description",
            "Part Number"
        ]


        for col in candidates:

            if col in self.columns:
                return col


        return None



    def find_package(self):

        candidates=[
            "Package_Type",
            "PCB Footprint",
            "Package"
        ]


        for col in candidates:

            if col in self.columns:
                return col


        return None