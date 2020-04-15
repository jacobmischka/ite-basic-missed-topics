import csv


def dump_csv(labels, body, outpath):
    rows = [labels, *body]

    with open(outpath, "w") as outfile:
        writer = csv.writer(outfile)
        writer.writerows(rows)


def get_csv_rows(sections):
    rows = []
    items = []

    for section in sections:
        rows += section.get_csv_rows()
        rows.append([])
        items += section.items

    rows.append([])

    rows.append(
        [
            "Overall averages",
            "",
            sum([item.cby_total for item in items]) / len(items),
            sum([item.cby for item in items]) / len(items),
            sum([item.ca1_total for item in items]) / len(items),
            sum([item.ca1 for item in items]) / len(items),
            sum([item.ca2_total for item in items]) / len(items),
            sum([item.ca2 for item in items]) / len(items),
            sum([item.ca3_total for item in items]) / len(items),
            sum([item.ca3 for item in items]) / len(items),
            "",
            sum([item.cby_diff for item in items]) / len(items),
            sum([item.ca1_diff for item in items]) / len(items),
            sum([item.ca2_diff for item in items]) / len(items),
            sum([item.ca3_diff for item in items]) / len(items),
            "",
            sum([item.cby_missed for item in items]) / len(items),
            sum([item.ca1_missed for item in items]) / len(items),
            sum([item.ca2_missed for item in items]) / len(items),
            sum([item.ca3_missed for item in items]) / len(items),
        ]
    )

    advanced_items = [item for item in items if item.item_type == "A"]
    basic_items = [item for item in items if item.item_type == "B"]

    rows.append(
        [
            "Advanced averages",
            "",
            sum([item.cby_total for item in advanced_items]) / len(advanced_items),
            sum([item.cby for item in advanced_items]) / len(advanced_items),
            sum([item.ca1_total for item in advanced_items]) / len(advanced_items),
            sum([item.ca1 for item in advanced_items]) / len(advanced_items),
            sum([item.ca2_total for item in advanced_items]) / len(advanced_items),
            sum([item.ca2 for item in advanced_items]) / len(advanced_items),
            sum([item.ca3_total for item in advanced_items]) / len(advanced_items),
            sum([item.ca3 for item in advanced_items]) / len(advanced_items),
            "",
            sum([item.cby_diff for item in advanced_items]) / len(advanced_items),
            sum([item.ca1_diff for item in advanced_items]) / len(advanced_items),
            sum([item.ca2_diff for item in advanced_items]) / len(advanced_items),
            sum([item.ca3_diff for item in advanced_items]) / len(advanced_items),
            "",
            sum([item.cby_missed for item in advanced_items]) / len(advanced_items),
            sum([item.ca1_missed for item in advanced_items]) / len(advanced_items),
            sum([item.ca2_missed for item in advanced_items]) / len(advanced_items),
            sum([item.ca3_missed for item in advanced_items]) / len(advanced_items),
        ]
    )
    rows.append(
        [
            "Basic averages",
            "",
            sum([item.cby_total for item in basic_items]) / len(basic_items),
            sum([item.cby for item in basic_items]) / len(basic_items),
            sum([item.ca1_total for item in basic_items]) / len(basic_items),
            sum([item.ca1 for item in basic_items]) / len(basic_items),
            sum([item.ca2_total for item in basic_items]) / len(basic_items),
            sum([item.ca2 for item in basic_items]) / len(basic_items),
            sum([item.ca3_total for item in basic_items]) / len(basic_items),
            sum([item.ca3 for item in basic_items]) / len(basic_items),
            "",
            sum([item.cby_diff for item in basic_items]) / len(basic_items),
            sum([item.ca1_diff for item in basic_items]) / len(basic_items),
            sum([item.ca2_diff for item in basic_items]) / len(basic_items),
            sum([item.ca3_diff for item in basic_items]) / len(basic_items),
            "",
            sum([item.cby_missed for item in basic_items]) / len(basic_items),
            sum([item.ca1_missed for item in basic_items]) / len(basic_items),
            sum([item.ca2_missed for item in basic_items]) / len(basic_items),
            sum([item.ca3_missed for item in basic_items]) / len(basic_items),
        ]
    )

    return rows


def dump_section_csv(sections, outpath):
    with open(outpath, "w") as outfile:
        writer = csv.writer(outfile)
        writer.writerows(get_csv_rows(sections))
