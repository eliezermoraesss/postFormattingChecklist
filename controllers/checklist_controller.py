from flask import Blueprint, render_template, request, redirect, url_for, send_file
from services.excel_service import (
    load_data,
    save_data,
    add_row,
    update_row,
    delete_row,
    EXCEL_PATH,
)

checklist_bp = Blueprint("checklist", __name__)


@checklist_bp.route("/checklist")
def checklist():
    df = load_data()
    return render_template("index.html", items=df.to_dict(orient="records"))


@checklist_bp.route("/add", methods=["POST"])
def add():
    descricao = request.form.get("descricao")
    programa = request.form.get("programa")
    add_row(descricao, programa)
    return redirect(url_for("checklist.checklist"))


@checklist_bp.route("/edit/<int:index>")
def edit(index):
    df = load_data()
    item = df.iloc[index].to_dict()
    return render_template("edit.html", item=item, index=index)


@checklist_bp.route("/update/<int:index>", methods=["POST"])
def update(index):
    descricao = request.form.get("descricao")
    programa = request.form.get("programa")
    checked = request.form.get("checked") == "on"
    update_row(index, descricao, programa, checked)
    return redirect(url_for("checklist.checklist"))


@checklist_bp.route("/delete/<int:index>")
def delete(index):
    delete_row(index)
    return redirect(url_for("checklist.checklist"))


@checklist_bp.route("/download")
def download():
    return send_file(EXCEL_PATH, as_attachment=True)
