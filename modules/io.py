#!/usr/bin/python

__all__ = ["save_progress", "save_and_close"]

def save_progress(workbook, path: str):
    if workbook == None:
        raise TypeError("workbook must not be NoneType")

    try:
        print("\nSaving file...")

        workbook.save(path)

        print("\nResults saved to: " + path)
    except Exception as e:
        print("\nAn error occurred while saving\n\n")
        print(e)

def save_and_close(workbook, result_book, result_path: str):
    try:
        if workbook != None:
            workbook.close()
    except Exception as e:
        print("\nFailed to close workbook\n")
        print(e)


    try:
        if result_book != None:
            print("\nSaving file...")

            result_book.save(result_path)
            result_book.close()

            print("\nResults saved to: " + result_path)
        else:
            print("\n\nNot saving...")
    except Exception as e:
        print("\nFailed to save and close result book\n")
        print(e)
