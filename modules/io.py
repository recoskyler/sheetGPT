#!/usr/bin/python

__all__ = ["save_progress", "save_and_close"]

def save_progress(workbook, path: str):
    if workbook == None:
        raise TypeError("workbook must not be NoneType")

    print("\nSaving file...")

    workbook.save(path)

    print("\nResults saved to: " + path)

def save_and_close(workbook, result_book, result_path: str):
    try:
        if workbook != None:
            workbook.close()
    except Exception as e:
        print("\nFailed to close workbook\n")
        print(e)


    if result_book != None:
        print("\nSaving file...")

        result_book.save(result_path)
        result_book.close()

        print("\nResults saved to: " + result_path)
    else:
        print("\n\nNot saving...")
