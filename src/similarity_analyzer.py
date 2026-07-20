import pandas as pd



class SimilarityAnalyzer:


    def analyze(self, df):


        result = []


        cap_df = df[
            df["Component_Type"]=="C"
        ]


        groups = (
            cap_df
            .groupby(
                "Normalized_Value"
            )
        )


        group_id = 1


        for value, group in groups:


            if len(group) < 2:
                continue



            result.append({

                "Group_ID":
                    f"CAP_GROUP_{group_id:03d}",


                "Normalized_Value":
                    value,


                "Count":
                    len(group),


                "References":
                    ",".join(
                        group["Part Reference"]
                        .astype(str)
                    ),


                "Package":
                    ",".join(
                        group["Package_Type"]
                        .unique()
                    ),


                "Description":
                    "Same capacitance value, check voltage/dielectric"


            })


            group_id +=1



        return pd.DataFrame(result)