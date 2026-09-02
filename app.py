import random
import pandas as pd

from flask import Flask, render_template, request, send_file
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Side

app = Flask(__name__)


def myRandom(start, stop, step):
    steps = int((stop - start) / step)
    val = start + step * random.randint(0, steps)
    return val


@app.route("/", methods=["GET", "POST"])
def home():

    data = []

    if request.method == "POST":

        roll = int(request.form["roll"])

        marks1 = 22
        marks2 = 24.5

        for i in range(1, roll + 1):
            x = myRandom(marks1, marks2, 0.5)

            data.append({
                "Roll": i,
                "Marks": x
            })

        # Create marks.xlsx
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

        wb.save("marks.xlsx")

    return render_template("index.html", data=data)


# DOWNLOAD ROUTE
@app.route("/download")
def download():

    return send_file(
        "marks.xlsx",
        as_attachment=True,
        download_name="marks.xlsx"
    )


if __name__ == "__main__":
    app.run(debug=True)
