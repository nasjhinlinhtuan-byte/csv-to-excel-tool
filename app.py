from flask import Flask, render_template, request, send_file
import pandas as pd
from openpyxl import load_workbook
import tempfile

app = Flask(__name__)

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/process", methods=["POST"])
def process():
    # Lấy file từ form
    csv_file = request.files["csv_file"]
    excel_template = request.files["excel_template"]

    # Đọc CSV
    df = pd.read_csv(csv_file)

    # Mở Excel mẫu
    wb = load_workbook(excel_template)
    ws = wb.active  # hoặc wb["TênSheet"] nếu bạn có sheet cụ thể

    # Ví dụ: điền dòng đầu tiên của CSV vào Excel
    # Bạn có thể sửa lại theo nhu cầu của bạn
    ws["B3"] = df["Date demande"][0]
    ws["G3"] = df["Date mise en service"][0]
    ws["G5"] = df["Nb expéditions"][0]
    ws["G6"] = df["Nb colis"][0]
    ws["B7"] = df["Produit"][0]
    ws["B9"] = df["Agence"][0]
    ws["B14"] = df["Commercial"][0]
    ws["G14"] = df["Présence installation"][0]
    ws["B16"] = df["Raison sociale"][0]
    ws["G16"] = df["SIRET administratif"][0]
    ws["B18"] = df["Adresse d'installation"][0]
    ws["B21"] = df["UTILISATEUR"][0]
    ws["F21"] = df["DECIDEUR"][0]
    ws["B22"] = df["Téléphone UTILISATEUR"][0]
    ws["F22"] = df["Téléphone DECIDEUR"][0]
    ws["B23"] = df["Email"][0]
values = df["Compte ERM"][0]  # là list: ["ERM11", "", "", "", ""]

for i, val in enumerate(values):
    ws[f"A{26 + i}"] = val


    ws["B26"] = df["Nom"][0]
    ws["C26"] = df["Kg"][0]
    ws["B14"] = df["Mes"][0]
    ws["B14"] = df["Exp"][0]
    ws["B14"] = df["Euro"][0]
    ws["B14"] = df["Pal"][0]
    ws["B14"] = df["Lot"][0]
    ws["B14"] = df["S+"][0]
    ws["B14"] = df["Hpd"][0]
    ws["D26"] = "☑" if df["Euro"][0] == "On" else "☐"
    ws["E26"] = "☑" if df["Euro"][0] == "On" else "☐"
    ws["D27"] = "☑" if df["Euro"][0] == "On" else "☐"










    # Tạo file tạm để trả về
    temp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    wb.save(temp.name)

    return send_file(temp.name, as_attachment=True, download_name="result.xlsx")

if __name__ == "__main__":
    app.run()
