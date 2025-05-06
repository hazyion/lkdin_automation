from flask import Flask, jsonify
from threading import Thread
from gen_and_req import start_process

app = Flask(__name__)

@app.route("/start-process", methods=["GET"])
def trigger_process():
    Thread(target=start_process).start()
    return jsonify({"status": "Process started"}), 200

if __name__ == "__main__":
    app.run(debug=True, port=5000)
