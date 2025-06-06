from flask import Flask, request, render_template
import pandas as pd
import os

app = Flask(__name__)

@app.route("/", methods=["GET", "POST"])
def upload():
    message = None
    if request.method == "POST":
        try:
            new_row = {
                "Region": request.form["region"],
                "PROGRAM": request.form["program"],
                "Quarter": request.form["quarter"],
                "hh": int(request.form["hh"]),
                "Weight": float(request.form["weight"]),
                "Cost": float(request.form["cost"])
            }

            excel_path = "uploaded_data.xlsx"

            if os.path.exists(excel_path):
                df = pd.read_excel(excel_path)
            else:
                df = pd.DataFrame(columns=["Region", "PROGRAM", "Quarter", "hh", "Weight", "Cost"])

            df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
            df.to_excel(excel_path, index=False)

            message = "✅ Upload successful!"
        except Exception as e:
            message = f"❌ Error: {str(e)}"

    return render_template("upload.html", message=message)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=True)



