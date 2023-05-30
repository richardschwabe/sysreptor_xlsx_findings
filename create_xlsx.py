import json
import xlsxwriter


def load_file(file_name):
    report_data = None
    with open(file_name, "r") as json_file:
        report_data = json.load(json_file)
    return report_data


def run():
    file_name = "report_export.json"
    report_data = load_file(file_name)

    workbook = xlsxwriter.Workbook("findings.xlsx")
    worksheet = workbook.add_worksheet("Findings")
    worksheet.hide_gridlines(2)

    # Default Row and Column Size
    worksheet.set_column("A:A", 10)
    worksheet.set_column("B:B", 30)
    worksheet.set_column("C:C", 80)
    worksheet.set_column("D:D", 50)

    # Add Default Header Format
    format_default_header = workbook.add_format()
    format_default_header.set_align("left")
    format_default_header.set_align("vcenter")
    format_default_header.set_text_wrap()
    format_default_header.set_bold()
    format_default_header.set_border(1)
    format_default_header.set_bottom(6)
    format_default_header.set_bg_color("#F2F2F2")
    format_default_header.set_font_size(16)

    # Add Default Row Format
    format_default_row = workbook.add_format()
    format_default_row.set_align("left")
    format_default_row.set_align("vcenter")
    format_default_row.set_text_wrap()
    format_default_row.set_border(1)

    # Add Solved Format
    format_solved = workbook.add_format(
        {"bold": True, "font_color": "#107A27", "bg_color": "#D1FFDB"}
    )
    worksheet.conditional_format(
        "A2:A99",
        {
            "type": "text",
            "criteria": "containing",
            "value": "Resolved",
            "format": format_solved,
        },
    )

    # Add Open Format
    format_open = workbook.add_format(
        {"bold": True, "font_color": "#787612", "bg_color": "#F2F1A2"}
    )
    worksheet.conditional_format(
        "A2:A99",
        {
            "type": "text",
            "criteria": "containing",
            "value": "Open",
            "format": format_open,
        },
    )

    # Add Ignored Format
    format_ignored = workbook.add_format(
        {"bold": True, "font_color": "#52101E", "bg_color": "#F2A2B3"}
    )
    worksheet.conditional_format(
        "A2:A99",
        {
            "type": "text",
            "criteria": "containing",
            "value": "Ignored",
            "format": format_ignored,
        },
    )

    # Write Header
    worksheet.write(0, 0, "Status", format_default_header)
    worksheet.write(0, 1, "Finding", format_default_header)
    worksheet.write(0, 2, "Recommendation", format_default_header)
    worksheet.write(0, 3, "CVSS", format_default_header)
    worksheet.set_row(0, 50)

    row = 1
    col = 0

    for finding in report_data["findings"]:
        # Set Finding Data
        finding_title = finding["data"]["title"]
        finding_short_recommendation = finding["data"]["short_recommendation"]
        finding_cvss = finding["data"]["cvss"]

        # Set Default List Item
        worksheet.write(row, col, "Open", format_default_row)

        # Format Row
        worksheet.set_row(row, 40)

        # Add Dropdowns
        worksheet.data_validation(
            row,
            col,
            row,
            col,
            {
                "validate": "list",
                "source": [
                    "Ignored",
                    "Resolved",
                    "Open",
                ],
                "value": "Open",
                "input_title": "Status of Finding",
                "input_message": "Choose from dropdown",
            },
        )

        # Write Finding Data
        worksheet.write(row, col + 1, finding_title, format_default_row)
        worksheet.write(row, col + 2, finding_short_recommendation, format_default_row)
        worksheet.write(row, col + 3, finding_cvss, format_default_row)

        row += 1

    workbook.close()


if __name__ == "__main__":
    run()
