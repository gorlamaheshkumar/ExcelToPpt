from flask import Flask

# Create the Flask web application object
app = Flask(__name__)

# Create a route for the homepage ("/")
@app.route("/")
def hello():
    # This is what will be displayed on the webpage
    return "<h1>✅ Hello world from a real web app!</h1>"