import utils


def test_pdf_content(zip_archive: str):
    pdf_filename = "Python_Testing_with_pytest.pdf"
    expected_text = "Python Testing with pytest,\nSecond Edition"
    page_number = 5

    page_text = utils.extract_pdf_text_from_zip(zip_archive, pdf_filename, page_number)
    assert (
        expected_text in page_text
    ), f"Текст '{expected_text}' не найден на странице {page_number} файла {pdf_filename}"


def test_csv_content(zip_archive: str):
    csv_filename = "Cities.csv"
    expected_dict = {
        "LatD": "   37",
        ' "LatM"': "   46",
        ' "LatS"': "   47",
        ' "NS"': ' "N"',
        ' "LonD"': "    122",
        ' "LonM"': "   25",
        ' "LonS"': "   11",
        ' "EW"': ' "W"',
        ' "City"': ' "San Francisco"',
        ' "State"': " CA",
    }

    csv_content = utils.extract_csv_content_from_zip(zip_archive, csv_filename)
    assert (
        expected_dict in csv_content
    ), f"Ожидаемая строка {expected_dict} не найдена в файле {csv_filename}"


def test_xlsx_content(zip_archive: str):
    xlsx_filename = "Financial_Sample.xlsx"
    expected_row = [
        "Midmarket",
        "Germany",
        "VTT",
        "None",
        888,
        250,
        15,
        13320,
        0,
        13320,
        8880,
        4440,
        6,
        "June",
        "2014",
    ]
    row_number = 43

    actual_row = utils.extract_xlsx_row_from_zip(
        zip_archive, xlsx_filename, row_number=row_number
    )

    assert (
        actual_row == expected_row
    ), f"Ожидалась строка {expected_row}, но получена {actual_row}"
