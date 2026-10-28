
from pathlib import Path
from flask import Flask, jsonify
import traceback

from profiler import profile_workbook
from exporter import export_profile

app = Flask(__name__)


@app.route("/", methods=["GET"])
def home():
    return jsonify({
        "message": "VELOOP Data Profiling API is running",
        "status": "success"
    })


@app.route("/profile", methods=["GET", "POST"])
def run_profiling():
    try:
        input_file = Path("data/VELoop_Rewards_Data_01.xlsx")
        output_file = Path("output/VELOOP_Profile_Output.xlsx")

        if not input_file.exists():
            return jsonify({
                "error": "Input Excel file not found",
                "path": str(input_file)
            }), 404

        results = profile_workbook(input_file)

        output_file.parent.mkdir(
            parents=True,
            exist_ok=True
        )

        export_profile(results, output_file)

        return jsonify({
            "message": "Data profiling completed successfully",
            "output": str(output_file)
        })

    except Exception as e:
        app.logger.exception("Profiling failed")
        return jsonify({
            "error": str(e)
        }), 500

