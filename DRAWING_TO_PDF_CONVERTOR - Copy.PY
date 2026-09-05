import win32com.client
import pythoncom
import os
import time


# ============================================================
# SETTINGS
# ============================================================

# Folder containing the SolidWorks drawing files
INPUT_FOLDER = r"D:\LOKESH\PROGRAMMING\sw automation\drawing to pdf"

# Output folder
OUTPUT_FOLDER = os.path.join(INPUT_FOLDER, "PDF")


# ============================================================
# CREATE OUTPUT FOLDER
# ============================================================

os.makedirs(OUTPUT_FOLDER, exist_ok=True)


# ============================================================
# CONNECT TO SOLIDWORKS 2026
# ============================================================

print("Connecting to SolidWorks 2026...")

sw = win32com.client.Dispatch("SldWorks.Application.34")

sw.Visible = True

print("Connected!")
print("SolidWorks version:", sw.RevisionNumber)


# ============================================================
# FIND ALL SLDDRW FILES
# ============================================================

drawings = [
    f
    for f in os.listdir(INPUT_FOLDER)
    if f.lower().endswith(".slddrw")
]

print()
print("Found", len(drawings), "drawing(s).")


if not drawings:
    print("No SLDDRW files found.")
    input("Press Enter to exit...")
    exit()


# ============================================================
# COUNTERS
# ============================================================

successful = 0
failed = 0


# ============================================================
# PROCESS EACH DRAWING
# ============================================================

for index, filename in enumerate(drawings, start=1):

    print()
    print("=" * 50)
    print(f"PROCESSING {index}/{len(drawings)}")
    print("File:", filename)
    print("=" * 50)

    # --------------------------------------------------------
    # FILE PATHS
    # --------------------------------------------------------

    drawing_path = os.path.join(
        INPUT_FOLDER,
        filename
    )

    pdf_name = os.path.splitext(filename)[0] + ".pdf"

    pdf_path = os.path.join(
        OUTPUT_FOLDER,
        pdf_name
    )


    # --------------------------------------------------------
    # OPEN DRAWING
    # --------------------------------------------------------

    print("Opening drawing...")

    errors = win32com.client.VARIANT(
        pythoncom.VT_BYREF | pythoncom.VT_I4,
        0
    )

    warnings = win32com.client.VARIANT(
        pythoncom.VT_BYREF | pythoncom.VT_I4,
        0
    )

    model = sw.OpenDoc6(
        drawing_path,
        3,
        1,
        "",
        errors,
        warnings
    )


    # --------------------------------------------------------
    # CHECK IF DRAWING OPENED
    # --------------------------------------------------------

    if model is None:

        print("FAILED TO OPEN DRAWING")

        print("Error:", errors.value)
        print("Warning:", warnings.value)

        failed += 1

        continue


    print("Drawing opened successfully.")

    # Give SolidWorks some time to finish loading
    time.sleep(2)


    # --------------------------------------------------------
    # EXPORT PDF
    # --------------------------------------------------------

    print("Exporting PDF...")

    save_errors = win32com.client.VARIANT(
        pythoncom.VT_BYREF | pythoncom.VT_I4,
        0
    )

    save_warnings = win32com.client.VARIANT(
        pythoncom.VT_BYREF | pythoncom.VT_I4,
        0
    )


    success = model.SaveAs4(
        pdf_path,
        0,
        1,
        save_errors,
        save_warnings
    )


    # --------------------------------------------------------
    # CHECK PDF
    # --------------------------------------------------------

    if success and os.path.exists(pdf_path):

        file_size = os.path.getsize(pdf_path)

        print("PDF CREATED SUCCESSFULLY!")

        print("PDF:", pdf_path)

        print(
            "Size:",
            file_size,
            "bytes"
        )

        successful += 1

    else:

        print("PDF CREATION FAILED")

        print(
            "Success:",
            success
        )

        print(
            "Save error:",
            save_errors.value
        )

        print(
            "Save warning:",
            save_warnings.value
        )

        failed += 1


    # --------------------------------------------------------
    # CLOSE DRAWING
    # --------------------------------------------------------

    print("Closing drawing...")

    # GetTitle is a PROPERTY in Python COM,
    # so do NOT use model.GetTitle()
    sw.CloseDoc(model.GetTitle)

    time.sleep(1)


# ============================================================
# FINAL RESULT
# ============================================================

print()
print("=" * 50)
print("CONVERSION COMPLETE")
print("=" * 50)

print("Successful:", successful)
print("Failed:", failed)

print()

print("PDF folder:")
print(OUTPUT_FOLDER)

print()

input("Press Enter to exit...")