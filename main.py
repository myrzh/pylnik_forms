import zipfile
import csv

from docx import Document
from docx.shared import Pt


GROUP_NUMBER = "5151003/30002"


def extract_csv(practice_number: str) -> None:
    with zipfile.ZipFile(f"ORG_{practice_number}.csv.zip", "r") as zip_ref:
        zip_ref.extractall("responses_extracted")


def actualize_records(records: list) -> list:
    used_names = []
    responses = []

    for record in records[::-1]:
        if record["Имя"] not in used_names:
            responses.append(record)
            used_names.append(record["Имя"])
    responses.reverse()

    return responses


def init_document(practice_number: str) -> Document:
    document = Document()

    style = document.styles["Normal"]
    font = style.font
    font.name = "Times New Roman"
    font.size = Pt(12)

    document.add_paragraph(f"Практика {practice_number}, группа {GROUP_NUMBER}\n")

    return document


def write_response(document: Document, response: dict) -> Document:
    document.add_paragraph(response["Имя"])

    for index in range(2, 4):
        document.add_paragraph(
            list(response.keys())[index].replace(" без заголовка", "")
        )
        document.add_paragraph(list(response.values())[index])

    document.add_page_break()

    return document


def main():
    # practice_number = input("Enter practice number: ")
    practice_number = "TEMPLATE"

    with open(
        f"responses_extracted/ORG_{practice_number}.csv", encoding="utf8"
    ) as csvfile:
        reader = csv.DictReader(csvfile, delimiter=",", quotechar='"')
        responses = actualize_records(list(reader))

    document = init_document(practice_number)

    for response in responses:
        write_response(document, response)

    document.save(f"Практика {practice_number}.docx")


if __name__ == "__main__":
    main()
