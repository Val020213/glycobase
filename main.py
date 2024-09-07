import pandas as pd


class glycobase:
    def __init__(self, name: str, range: str) -> None:
        self.name: str = name
        range = range.replace(",", ".")
        split = range.split("-")
        self.range: tuple[str, str] = (
            (split[0], split[0]) if len(split) < 2 else (split[0], split[1])
        )
        self.range: tuple[float, float] = (float(self.range[0]), float(self.range[1]))

    def is_glycobase(self, gu: float):
        return self.range[0] <= gu and gu <= self.range[1]


class glycobase2GB:
    def __init__(self) -> None:
        gb = pd.read_excel("./db/2 GB.xlsx")
        gb.columns = [
            "",
            "name_gb_color",
            "range_gb_color",
            "",
            "",
            "name_gb_bw",
            "range_gb_bw",
        ]

        self.glycobases_color: list[glycobase] = []
        self.glycobases_bw: list[glycobase] = []

        for i in range(len(gb)):
            if (
                str(gb["name_gb_color"][i]) != "nan"
                and str(gb["range_gb_color"][i]) != "nan"
            ):
                self.glycobases_color.append(
                    glycobase(str(gb["name_gb_color"][i]), str(gb["range_gb_color"][i]))
                )

            if str(gb["name_gb_bw"][i]) != "nan" and str(gb["range_gb_bw"][i]) != "nan":
                self.glycobases_bw.append(
                    glycobase(str(gb["name_gb_bw"][i]), str(gb["range_gb_bw"][i]))
                )
        return

    def get_glycobase_color(self):
        return self.glycobases_color

    def get_black_and_white_gb(self):
        return self.glycobases_bw


if __name__ == "__main__":
    gb = glycobase2GB()  # read glycobase 2 GB
    all_experiments = pd.read_excel(
        "./db/experiments.xlsx", sheet_name=None
    )  # read experiments
    all_experiment_color_classifier = all_experiments
    all_experiment_bw_classifier = all_experiments

    for page in all_experiments:

        experiments = all_experiments[page]
        experiment_color_classifier = experiments
        experiment_bw_classifier = experiments

        # mark as GU the column of experiments
        columns = []
        for i in experiments.values[1]:
            columns.append(i)
        experiments.columns = columns

        # classify the experiments

        for index, row in experiments[2:].iterrows():
            row_gu = row["GU"].dropna()

            glycobase_color_classifier = []
            glycobase_bw_classifier = []

            for gbc in gb.get_glycobase_color():
                is_gb = True
                for gu in row_gu:
                    if not gbc.is_glycobase(gu):
                        is_gb = False
                        break
                if is_gb:
                    glycobase_color_classifier.append(gbc)

            for gb_bw in gb.get_black_and_white_gb():
                is_gb = True
                for gu in row_gu:
                    if not gb_bw.is_glycobase(gu):
                        is_gb = False
                        break
                if is_gb:
                    glycobase_bw_classifier.append(gb_bw)

            # add the classification to the row
            cont = 0
            for gbc in glycobase_color_classifier:
                experiment_color_classifier.at[index, "Name" + str(cont)] = gbc.name
                experiment_color_classifier.at[index, "Range" + str(cont)] = str(
                    gbc.range
                )
                cont += 1

            cont = 0
            for gbc in glycobase_bw_classifier:
                experiment_bw_classifier.at[index, "Name" + str(cont)] = gbc.name
                experiment_bw_classifier.at[index, "Range" + str(cont)] = str(gbc.range)
                cont += 1

            all_experiment_color_classifier[page] = experiment_color_classifier
            all_experiment_bw_classifier[page] = experiment_bw_classifier

    # save the classification
    with pd.ExcelWriter(
        "./db/experiments_color_classifier.xlsx", engine="xlsxwriter"
    ) as writer:
        for sheet_name, df in all_experiment_color_classifier.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)

    with pd.ExcelWriter(
        "./db/experiments_bw_classifier.xlsx", engine="xlsxwriter"
    ) as writer:
        for sheet_name, df in all_experiment_bw_classifier.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)
