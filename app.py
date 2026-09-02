import random
import pandas as pd

from flask import Flask, render_template, request, send_file
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Side

app = Flask(__name__)


def myRandom(start, stop, step):
    # Integer-based calculation avoids floating-point bias/errors
    start_i = round(start / step)
    stop_i = round(stop / step)

    return random.randint(start_i, stop_i) * step


@app.route("/", methods=["GET", "POST"])
def home():

    data = []

    # Default values
    students = ""
    starting_marks = 22.0
    maximum_marks = 24.5
    total_marks = 25.0

    if request.method == "POST":

        students = int(request.form["students"])
        starting_marks = float(request.form["starting_marks"])
        maximum_marks = float(request.form["maximum_marks"])
        total_marks = float(request.form["total_marks"])

        # Validation
        if starting_marks > maximum_marks:
            return render_template(
                "index.html",
                error="Starting marks cannot be greater than maximum marks.",
                students=students,
                starting_marks=starting_marks,
                maximum_marks=maximum_marks,
                total_marks=total_marks
            )

        if maximum_marks > total_marks:
            return render_template(
                "index.html",
                error="Maximum marks cannot be greater than total marks.",
                students=students,
                starting_marks=starting_marks,
                maximum_marks=maximum_marks,
                total_marks=total_marks
            )

        # Generate marks
        for i in range(1, students + 1):

            x = myRandom(
                starting_marks,
                maximum_marks,
                0.5
            )

            data.append({
                "Roll": i,
                "Marks": round(x, 1)
            })

        # Create Excel
        df = pd.DataFrame(data)
        df.to_excel("marks.xlsx", index=False)

        # Format Excel
        wb = load_workbook("marks.xlsx")
        ws = wb.active

        thin = Side(style="thin")

        border = Border(
            left=thin,
            right=thin,
            top=thin,
            bottom=thin
        )

        for row in ws.iter_rows():
            for cell in row:

                cell.alignment = Alignment(
                    horizontal="center",
                    vertical="center"
                )

                cell.border = border

        # Column width
        ws.column_dimensions["A"].width = 12
        ws.column_dimensions["B"].width = 12

        wb.save("marks.xlsx")

    return render_template(
        "index.html",
        data=data,
        students=students,
        starting_marks=starting_marks,
        maximum_marks=maximum_marks,
        total_marks=total_marks
    )


@app.route("/download")
def download():

    return send_file(
        "marks.xlsx",
        as_attachment=True,
        download_name="marks.xlsx"
    )


if __name__ == "__main__":
    app.run(debug=True)
