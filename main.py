import zipfile
import csv


def main():
    practice_number = input("Enter practice number: ")

    with zipfile.ZipFile(f"responses/ORG_{practice_number}.csv.zip", "r") as zip_ref:
        zip_ref.extractall("responses_extracted")

    with open(
        f"responses_extracted/ORG_{practice_number}.csv", encoding="utf8"
    ) as csvfile:
        reader = csv.reader(csvfile, delimiter=";", quotechar='"')
        # print(reader)
        for index, row in enumerate(reader):
            print(index, row)


if __name__ == "__main__":
    main()
