import argparse
import json
import re
import time

import openpyxl


def save_json(lesson_json, fp_json):
    with open(fp_json, "w", encoding="utf-8") as f:
        json.dump(lesson_json, f, ensure_ascii=False, indent=2)


def get_item(item: str, default="") -> str:
    return (
        (item.strip() if type(item) == str else item)
        if item
        else default
    )


def xlsx2json(fp_xlsx):
    def previous_row():
        nonlocal row_idx
        row_idx -= 1

    def next_row():
        nonlocal row_idx
        row_idx += 1

    lesson_json = {}

    workbook = openpyxl.load_workbook(fp_xlsx)
    sheet = workbook.active

    row_idx = 0

    next_row()
    lesson_json["id"] = int(get_item(
        sheet["B1"].value,
        default=re.search(r"\d+", fp_xlsx).group(0))
    )

    next_row()
    lesson_json["title"] = get_item(sheet["B2"].value)

    lesson_json["vocab"] = []
    while True:
        next_row()
        if sheet[f"B{row_idx}"].value != "FORMS":
            previous_row()
            break

        word = {}

        stem_list = list(filter(bool, get_item(sheet[f"C{row_idx}"].value, "").split("\n")))
        prefixes_list = list(filter(bool, get_item(sheet[f"D{row_idx}"].value, "").split("\n")))
        suffices_list = list(filter(bool, get_item(sheet[f"E{row_idx}"].value, "").split("\n")))
        articles_list = list(filter(bool, get_item(sheet[f"F{row_idx}"].value, "").split("\n")))
        assert len(stem_list) == len(prefixes_list) == len(suffices_list) == len(articles_list)

        # print(stem_list)
        # print(prefixes_list)
        # print(suffices_list)
        # print(articles_list)
        # next_row()
        # next_row()
        # continue

        forms = []
        for stem, prefixes, suffices, articles in zip(
                stem_list,
                prefixes_list,
                suffices_list,
                articles_list
        ):
            if not stem.strip():
                continue

            form_dict = {}
            if stem != '-':
                form_dict["stem"] = stem
            if prefixes.strip() and prefixes != '-':
                form_dict["prefixes"] = prefixes.split(",")
            if suffices.strip() and suffices != '-':
                form_dict["suffices"] = suffices.split(",")
            if articles.strip() and articles != '-':
                form_dict["articles"] = articles.split(",")
            forms.append(form_dict)
        word["forms"] = forms

        next_row()

        pos_list = list(filter(bool, get_item(sheet[f"C{row_idx}"].value, "").split("\n")))
        meanings_list = list(filter(bool, get_item(sheet[f"D{row_idx}"].value, "").split("\n")))
        labels_list = list(filter(bool, get_item(sheet[f"E{row_idx}"].value, "").split("\n")))
        usage_list = list(filter(bool, get_item(sheet[f"F{row_idx}"].value, "").split("\n")))
        assert len(pos_list) == len(meanings_list) == len(usage_list)

        meanings = []
        for pos, meanings_, labels, usage in zip(
                pos_list,
                meanings_list,
                labels_list,
                usage_list
        ):
            # if pos == "-":
            #     continue
            #
            # meanings_dict = {"pos": pos, "meanings": meanings_}
            meanings_dict = {}
            if pos.strip() != "-":
                meanings_dict["pos"] = pos
            meanings_dict["meanings"] = meanings_
            if labels.strip() != "-":
                meanings_dict["labels"] = labels.split(",")
            if usage.strip() != "-":
                # meanings_dict["usage"] = usage
                meanings_dict["usage"] = usage.split(",")
            meanings.append(meanings_dict)
        word["meanings"] = meanings

        next_row()

        explanation = sheet[f"C{row_idx}"].value
        # if explanation.strip() != "":
        if explanation and explanation.strip():
            word["explanation"] = explanation

        lesson_json["vocab"].append(word)

    lesson_json["sentences"] = []
    while True:
        next_row()
        if sheet[f"B{row_idx}"].value not in ["GREEK", "ENGLISH"]:
            previous_row()
            break

        if sheet[f"B{row_idx}"].value == "GREEK":
            sentence_dict = {"greek": get_item(sheet[f"C{row_idx}"].value)}
            next_row()
            sentence_dict["english_"] = get_item(sheet[f"C{row_idx}"].value)
        elif sheet[f"B{row_idx}"].value == "ENGLISH":
            sentence_dict = {"english": get_item(sheet[f"C{row_idx}"].value)}
            next_row()
            sentence_dict["greek_"] = get_item(sheet[f"C{row_idx}"].value)
        else:
            raise NotImplementedError

        lesson_json["sentences"].append(sentence_dict)

    lesson_json["reading"] = {}

    next_row()
    lesson_json["reading"]["title"] = get_item(sheet[f"C{row_idx}"].value)

    next_row()
    lesson_json["reading"]["text"] = get_item(sheet[f"C{row_idx}"].value)

    lesson_json["reading"]["vocab"] = []
    while True:
        next_row()
        if sheet[f"B{row_idx}"].value == "TRANSLATION":
            previous_row()
            break

        word_dict = {
            "word": get_item(sheet[f"C{row_idx}"].value),
            "explanation": get_item(sheet[f"D{row_idx}"].value)
        }
        lesson_json["reading"]["vocab"].append(word_dict)

    next_row()
    lesson_json["reading"]["translation"] = get_item(sheet[f"C{row_idx}"].value)

    # For lessons wo reading.
    if not lesson_json["reading"]["title"].strip():
        del lesson_json["reading"]

    return lesson_json


def main(fp_xlsx, fp_json):
    lesson_json = xlsx2json(fp_xlsx=fp_xlsx)
    save_json(lesson_json, fp_json)


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument("--xlsx", required=True)
    parser.add_argument("--json")
    args = parser.parse_args()

    if not args.json:
        args.json = args.xlsx.replace(".xlsx", f"_{time.strftime('%y%m%d%H%M%S')}.json")

    main(fp_xlsx=args.xlsx, fp_json=args.json)
