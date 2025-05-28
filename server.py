from flask import Flask, request, jsonify
from threading import Thread
from gen_and_req import start_process

app = Flask(__name__)


@app.route("/start-process", methods=["POST"])
def trigger_process():
    data = request.get_json()
    sheet = data.get("sheet")

    if not sheet:
        return jsonify({"error": "Missing 'sheet' field in request body"}), 400

    Thread(target=start_process, args=(sheet,)).start()
    return jsonify({"status": "Process started"}), 200


if __name__ == "__main__":
    app.run(debug=True, port=8000)
