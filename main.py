import zipfile
import csv

from pprint import pprint
from docx import Document
from docx.shared import Pt


GROUP_NUMBER = "5151003/30002"


def main():
    # practice_number = input("Enter practice number: ")
    practice_number = "TEMPLATE"

    with zipfile.ZipFile(f"ORG_{practice_number}.csv.zip", "r") as zip_ref:
        zip_ref.extractall("responses_extracted")

    with open(
        f"responses_extracted/ORG_{practice_number}.csv", encoding="utf8"
    ) as csvfile:
        reader = csv.DictReader(csvfile, delimiter=",", quotechar='"')
        records = list(reader)

    responses = records.copy()
    # for record in records:
    #     if responses:
    #         abnormal_exit = False
    #         for index, item in enumerate(responses):
    #             if item["Имя"] == record["Имя"]:
    #                 print("YASS")
    #                 responses[index] == record
    #                 abnormal_exit = True
    #                 break
    #         if not abnormal_exit:
    #             responses.append(record)
    #     else:
    #         responses.append(record)

    # pprint(responses)

    document = Document()

    document.add_paragraph(f"Практика {practice_number}, группа {GROUP_NUMBER}")
    document.add_paragraph("\n")

    style = document.styles["Normal"]
    font = style.font
    font.name = "Times New Roman"
    font.size = Pt(12)

    for response in responses:
        document.add_paragraph(response["Имя"])
        document.add_paragraph(list(response.keys())[2].replace(" без заголовка", ""))
        document.add_paragraph(list(response.values())[2])
        document.add_paragraph(list(response.keys())[3].replace(" без заголовка", ""))
        document.add_paragraph(list(response.values())[3])
        document.add_page_break()

    document.save(f"Практика {practice_number}.docx")


if __name__ == "__main__":
    main()
