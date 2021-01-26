#! /usr/bin/env ipython3
from json import dump, load
from threading import Timer
from time import sleep
from webbrowser import open as browse

from flask import Flask, flash, redirect, render_template, request
from flask_wtf import FlaskForm
from werkzeug.datastructures import ImmutableMultiDict, MultiDict
from wtforms import (BooleanField, SelectField, SubmitField,  # StringField,
                     TextField, validators)

from gdrive_tools import authorisedDrive, deAuthorise, getMainFolder

# App config.
DEBUG = True
app = Flask(__name__)
app.config.from_object(__name__)
app.config["SECRET_KEY"] = "7d441f27d441f27567d441f2b6176a"


class ConfigForm(FlaskForm):
    delay = TextField(
        "Delay:", description="Time between program updates/runs (in minutes)"
    )
    autoformat = BooleanField(
        "Auto-format Documents?",
        description="Whether or not the app should auto-format resolution documents",
    )
    folderpath = TextField(
        "Folder Path:",
        description="The Google Drive folder where the app should look for MUN documents. Use '/' to denote the root folder, and as a separator.\n Eg: /Documents/MUN means it's in a subfolder MUN which itself is within a subfolder 'Documents",
    )
    submit = SubmitField("Save Settings")

    # rule=RuleForm()

    @app.route("/", methods=["GET", "POST"])
    def hello():
        """Creates a form, renders and auto-fills it, and (when the form is saved) writes the data to config.json in the requisite format.

        :returns: A rendered flask/WTForms template
        :rtype:

        """
        form = ConfigForm(gen_multi_dict())
        if request.method == "POST":
            delay = (
                request.form["delay"]
                if "delay" in request.form
                else load("./config.json")["delay"]
            )
            autoformat = "autoformat" in request.form
            folderpath = (
                request.form["folderpath"]
                if "folderpath" in request.form
                else load("./config.json")["folderpath"]
            )
            print(request.form)
            rule_matrix = [
                request.form.getlist("type"),
                request.form.getlist("text"),
                request.form.getlist("doctype"),
            ]
            rule_count = len(rule_matrix[0])
            rule_json = {"name": [], "contains": []}
            for i in range(rule_count):
                formatted_rule = {"regex": rule_matrix[1][i], "type": rule_matrix[2][i]}
                rule_json[rule_matrix[0][i]].append(formatted_rule)

            conf_dict = {
                "delay": delay,
                "autoformat": autoformat,
                "folderpath": folderpath,
                "folderlink": getMainFolder(folderpath)["alternateLink"],
                "custom-rules": rule_json,
            }
            dump(conf_dict, open("./config.json", "w"))

        if form.validate():
            # Save the comment here.
            flash("Changes saved")

        return render_template("config-form.jinja.html", form=form)

    @app.context_processor
    def utility_processor():
        """ A set of processors that expose certain values to the Jinja2 template, so they can be used in the form more easily.

        :returns: A dictionary, with functions as values.
        :rtype: Dict

        """
        def currentJson():
            return load(open("./config.json"))

        def defaultJson():
            return {"delay": "10", "autoformat": False, "root": "/MUN"}

        def linkFromJson():
            folderpath = load(open("./config.json"))["folderpath"]
            return getMainFolder(folderpath)["alternateLink"]

        def customRules():
            data = []
            with open("config.json") as conf_json:
                conf = load(conf_json)["custom-rules"]
            for i in conf["name"]:
                data.append(["name", i["regex"], i["type"]])
            for i in conf["contains"]:
                data.append(["contains", i["regex"], i["type"]])
            return data

        return dict(
            current=currentJson,
            default=defaultJson,
            link=linkFromJson,
            rules=customRules,
        )

    @app.route("/auth", methods=["GET", "POST"])
    def auth():
        authorisedDrive()
        return redirect("/")

    @app.route("/deauth", methods=["GET", "POST"])
    def deauth():
        deAuthorise()
        return redirect("/")


def gen_multi_dict(config_file="config.json"):
    """Converts data from the config.json file into a MultiDict format, so that it can be read by flask and passed to the form

    :param config_file: Path to the config file to read from
    :returns: A MultiDict containing the rules from the config.json file
    :rtype: MultiDict

    """
    data = MultiDict()
    with open(config_file) as conf_json:
        conf = load(conf_json)
    for i in ("delay", "autoformat", "folderpath"):
        data.add(i, conf[i])
    # Now we have to parse the custom rules.
    rules = conf["custom-rules"]
    name_rules = rules["name"]
    contain_rules = rules["contains"]
    for i in name_rules:
        data.add("type", "name")
        data.add("text", i["regex"])
        data.add("doctype", i["type"])
        data.add("type", "regex")
    for i in contain_rules:
        data.add("text", i["regex"])
        data.add("doctype", i["type"])
    return data


# https://stackoverflow.com/questions/54235347/open-browser-automatically-when-python-code-is-executed
if __name__ == "__main__":
    browse("http://127.0.0.1:5000")
    app.run()
