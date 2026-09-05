import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import win32com.client
import pythoncom
import os
import time
import threading


# ============================================================
# SOLIDWORKS DRAWING TO PDF CONVERTER
# ============================================================


class DrawingToPDFConverter:

    def __init__(self, root):

        self.root = root
        self.root.title("SolidWorks Drawing to PDF Converter")
        self.root.geometry("750x550")
        self.root.resizable(False, False)

        # ----------------------------------------------------
        # VARIABLES
        # ----------------------------------------------------

        self.input_folder = tk.StringVar()
        self.output_folder = tk.StringVar()

        self.is_converting = False


        # ----------------------------------------------------
        # TITLE
        # ----------------------------------------------------

        title = tk.Label(
            root,
            text="SOLIDWORKS DRAWING → PDF",
            font=("Segoe UI", 20, "bold")
        )

        title.pack(pady=(25, 5))


        subtitle = tk.Label(
            root,
            text="Batch convert SLDDRW drawings to PDF",
            font=("Segoe UI", 10)
        )

        subtitle.pack(pady=(0, 20))


        # ----------------------------------------------------
        # INPUT FOLDER
        # ----------------------------------------------------

        input_frame = tk.Frame(root)

        input_frame.pack(
            fill="x",
            padx=30,
            pady=8
        )


        tk.Label(
            input_frame,
            text="Input Folder",
            font=("Segoe UI", 10, "bold")
        ).pack(
            anchor="w"
        )


        input_row = tk.Frame(input_frame)

        input_row.pack(
            fill="x",
            pady=5
        )


        self.input_entry = tk.Entry(
            input_row,
            textvariable=self.input_folder,
            font=("Segoe UI", 10)
        )

        self.input_entry.pack(
            side="left",
            fill="x",
            expand=True,
            ipady=7
        )


        tk.Button(
            input_row,
            text="Browse",
            width=12,
            command=self.select_input_folder
        ).pack(
            side="right",
            padx=(10, 0),
            ipady=4
        )


        # ----------------------------------------------------
        # OUTPUT FOLDER
        # ----------------------------------------------------

        output_frame = tk.Frame(root)

        output_frame.pack(
            fill="x",
            padx=30,
            pady=8
        )


        tk.Label(
            output_frame,
            text="Output Folder",
            font=("Segoe UI", 10, "bold")
        ).pack(
            anchor="w"
        )


        output_row = tk.Frame(output_frame)

        output_row.pack(
            fill="x",
            pady=5
        )


        self.output_entry = tk.Entry(
            output_row,
            textvariable=self.output_folder,
            font=("Segoe UI", 10)
        )

        self.output_entry.pack(
            side="left",
            fill="x",
            expand=True,
            ipady=7
        )


        tk.Button(
            output_row,
            text="Browse",
            width=12,
            command=self.select_output_folder
        ).pack(
            side="right",
            padx=(10, 0),
            ipady=4
        )


        # ----------------------------------------------------
        # CONVERT BUTTON
        # ----------------------------------------------------

        self.convert_button = tk.Button(
            root,
            text="CONVERT TO PDF",
            font=("Segoe UI", 11, "bold"),
            width=25,
            height=2,
            command=self.start_conversion
        )

        self.convert_button.pack(
            pady=20
        )


        # ----------------------------------------------------
        # PROGRESS
        # ----------------------------------------------------

        tk.Label(
            root,
            text="Progress",
            font=("Segoe UI", 10, "bold")
        ).pack(
            anchor="w",
            padx=30
        )


        self.progress = ttk.Progressbar(
            root,
            orient="horizontal",
            length=690,
            mode="determinate"
        )

        self.progress.pack(
            padx=30,
            pady=8
        )


        # ----------------------------------------------------
        # STATUS
        # ----------------------------------------------------

        self.status_label = tk.Label(
            root,
            text="Ready",
            font=("Segoe UI", 9)
        )

        self.status_label.pack(
            anchor="w",
            padx=30
        )


        # ----------------------------------------------------
        # LOG
        # ----------------------------------------------------

        tk.Label(
            root,
            text="Conversion Log",
            font=("Segoe UI", 10, "bold")
        ).pack(
            anchor="w",
            padx=30,
            pady=(15, 5)
        )


        self.log_box = tk.Text(
            root,
            height=8,
            width=85,
            font=("Consolas", 9)
        )

        self.log_box.pack(
            padx=30
        )

        self.log_box.config(
            state="disabled"
        )


    # ========================================================
    # SELECT INPUT FOLDER
    # ========================================================

    def select_input_folder(self):

        folder = filedialog.askdirectory(
            title="Select folder containing SLDDRW files"
        )

        if folder:

            self.input_folder.set(folder)

            # Automatically suggest PDF folder
            default_output = os.path.join(
                folder,
                "PDF"
            )

            self.output_folder.set(default_output)


    # ========================================================
    # SELECT OUTPUT FOLDER
    # ========================================================

    def select_output_folder(self):

        folder = filedialog.askdirectory(
            title="Select PDF output folder"
        )

        if folder:

            self.output_folder.set(folder)


    # ========================================================
    # LOG MESSAGE
    # ========================================================

    def log(self, message):

        self.log_box.config(
            state="normal"
        )

        self.log_box.insert(
            tk.END,
            message + "\n"
        )

        self.log_box.see(
            tk.END
        )

        self.log_box.config(
            state="disabled"
        )


    # ========================================================
    # UPDATE STATUS
    # ========================================================

    def update_status(self, message):

        self.status_label.config(
            text=message
        )

        self.root.update_idletasks()


    # ========================================================
    # START CONVERSION
    # ========================================================

    def start_conversion(self):

        if self.is_converting:

            return


        input_folder = self.input_folder.get().strip()
        output_folder = self.output_folder.get().strip()


        # ----------------------------------------------------
        # CHECK INPUT
        # ----------------------------------------------------

        if not input_folder:

            messagebox.showerror(
                "Error",
                "Please select an input folder."
            )

            return


        if not os.path.isdir(input_folder):

            messagebox.showerror(
                "Error",
                "Input folder does not exist."
            )

            return


        # ----------------------------------------------------
        # CHECK OUTPUT
        # ----------------------------------------------------

        if not output_folder:

            messagebox.showerror(
                "Error",
                "Please select an output folder."
            )

            return


        os.makedirs(
            output_folder,
            exist_ok=True
        )


        # ----------------------------------------------------
        # FIND DRAWINGS
        # ----------------------------------------------------

        drawings = [
            f
            for f in os.listdir(input_folder)
            if f.lower().endswith(".slddrw")
        ]


        if not drawings:

            messagebox.showwarning(
                "No Drawings",
                "No SLDDRW files were found in the selected folder."
            )

            return


        # ----------------------------------------------------
        # RESET UI
        # ----------------------------------------------------

        self.progress["value"] = 0
        self.progress["maximum"] = len(drawings)

        self.log_box.config(
            state="normal"
        )

        self.log_box.delete(
            "1.0",
            tk.END
        )

        self.log_box.config(
            state="disabled"
        )


        self.is_converting = True

        self.convert_button.config(
            state="disabled"
        )


        # ----------------------------------------------------
        # RUN CONVERSION IN BACKGROUND
        # ----------------------------------------------------

        thread = threading.Thread(
            target=self.convert_drawings,
            args=(input_folder, output_folder, drawings),
            daemon=True
        )

        thread.start()


    # ========================================================
    # CONVERT DRAWINGS
    # ========================================================

    def convert_drawings(
        self,
        input_folder,
        output_folder,
        drawings
    ):

        pythoncom.CoInitialize()

        successful = 0
        failed = 0

        sw = None


        try:

            # ------------------------------------------------
            # CONNECT TO SOLIDWORKS 2026
            # ------------------------------------------------

            self.update_status(
                "Connecting to SolidWorks 2026..."
            )

            self.log(
                "Connecting to SolidWorks 2026..."
            )


            sw = win32com.client.Dispatch(
                "SldWorks.Application.34"
            )

            sw.Visible = True


            self.log(
                "Connected to SolidWorks."
            )

            self.log(
                "SolidWorks version: "
                + sw.RevisionNumber
            )


            # ------------------------------------------------
            # PROCESS EACH DRAWING
            # ------------------------------------------------

            for index, filename in enumerate(
                drawings,
                start=1
            ):

                self.update_status(
                    f"Converting {index}/{len(drawings)}: {filename}"
                )


                self.log(
                    ""
                )

                self.log(
                    f"[{index}/{len(drawings)}] {filename}"
                )


                drawing_path = os.path.join(
                    input_folder,
                    filename
                )


                pdf_name = (
                    os.path.splitext(filename)[0]
                    + ".pdf"
                )


                pdf_path = os.path.join(
                    output_folder,
                    pdf_name
                )


                # --------------------------------------------
                # OPEN DRAWING
                # --------------------------------------------

                self.log(
                    "Opening drawing..."
                )


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


                if model is None:

                    self.log(
                        "FAILED TO OPEN DRAWING"
                    )

                    self.log(
                        "Error: "
                        + str(errors.value)
                    )

                    self.log(
                        "Warning: "
                        + str(warnings.value)
                    )

                    failed += 1

                    self.root.after(
                        0,
                        self.update_progress,
                        index
                    )

                    continue


                self.log(
                    "Drawing opened successfully."
                )


                time.sleep(2)


                # --------------------------------------------
                # EXPORT PDF
                # --------------------------------------------

                self.log(
                    "Exporting PDF..."
                )


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


                # --------------------------------------------
                # CHECK RESULT
                # --------------------------------------------

                if (
                    success
                    and
                    os.path.exists(pdf_path)
                ):

                    file_size = os.path.getsize(
                        pdf_path
                    )


                    self.log(
                        "✓ PDF CREATED SUCCESSFULLY"
                    )

                    self.log(
                        f"PDF: {pdf_name}"
                    )

                    self.log(
                        f"Size: {file_size:,} bytes"
                    )


                    successful += 1


                else:

                    self.log(
                        "✗ PDF CREATION FAILED"
                    )

                    self.log(
                        "Save error: "
                        + str(save_errors.value)
                    )

                    self.log(
                        "Save warning: "
                        + str(save_warnings.value)
                    )


                    failed += 1


                # --------------------------------------------
                # CLOSE DRAWING
                # --------------------------------------------

                self.log(
                    "Closing drawing..."
                )


                sw.CloseDoc(
                    model.GetTitle
                )


                time.sleep(1)


                # --------------------------------------------
                # UPDATE PROGRESS
                # --------------------------------------------

                self.root.after(
                    0,
                    self.update_progress,
                    index
                )


            # ------------------------------------------------
            # FINISHED
            # ------------------------------------------------

            self.root.after(
                0,
                self.conversion_finished,
                successful,
                failed,
                output_folder
            )


        except Exception as e:

            self.root.after(
                0,
                self.conversion_error,
                str(e)
            )


        finally:

            pythoncom.CoUninitialize()


    # ========================================================
    # UPDATE PROGRESS BAR
    # ========================================================

    def update_progress(self, value):

        self.progress["value"] = value


    # ========================================================
    # CONVERSION FINISHED
    # ========================================================

    def conversion_finished(
        self,
        successful,
        failed,
        output_folder
    ):

        self.is_converting = False

        self.convert_button.config(
            state="normal"
        )


        self.progress["value"] = self.progress["maximum"]


        self.update_status(
            "Conversion complete."
        )


        self.log(
            ""
        )

        self.log(
            "========================================"
        )

        self.log(
            "CONVERSION COMPLETE"
        )

        self.log(
            f"Successful: {successful}"
        )

        self.log(
            f"Failed: {failed}"
        )

        self.log(
            f"Output folder: {output_folder}"
        )

        self.log(
            "========================================"
        )


        if failed == 0:

            messagebox.showinfo(
                "Conversion Complete",
                f"{successful} drawing(s) converted successfully."
            )

        else:

            messagebox.showwarning(
                "Conversion Complete",
                f"Successful: {successful}\n"
                f"Failed: {failed}"
            )


    # ========================================================
    # CONVERSION ERROR
    # ========================================================

    def conversion_error(self, error):

        self.is_converting = False

        self.convert_button.config(
            state="normal"
        )


        self.update_status(
            "Conversion failed."
        )


        self.log(
            ""
        )

        self.log(
            "ERROR:"
        )

        self.log(
            error
        )


        messagebox.showerror(
            "Conversion Error",
            error
        )


# ============================================================
# START APPLICATION
# ============================================================

if __name__ == "__main__":

    root = tk.Tk()

    app = DrawingToPDFConverter(
        root
    )

    root.mainloop()