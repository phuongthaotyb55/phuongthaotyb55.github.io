from flask_frozen import Freezer

from app import app

app.config["FREEZER_DESTINATION"] = "build"
app.config["FREEZER_RELATIVE_URLS"] = False
app.config["FREEZER_TRAILING_SLASH"] = True

freezer = Freezer(app)

if __name__ == "__main__":
    freezer.freeze()
