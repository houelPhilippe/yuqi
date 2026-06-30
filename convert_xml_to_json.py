import xmlschema
import json
import argparse
import os
import sys
from decimal import Decimal
import xml.etree.ElementTree as ET


def json_default(value):
    if isinstance(value, Decimal):
        return float(value)
    return str(value)


XSD_DIR = "xsd"
RESULT_DIR = "result"


def element_to_dict(element):
    children = list(element)
    if not children:
        return element.text.strip() if element.text and element.text.strip() else None

    result = {}
    for child in children:
        child_value = element_to_dict(child)
        if child.tag in result:
            if not isinstance(result[child.tag], list):
                result[child.tag] = [result[child.tag]]
            result[child.tag].append(child_value)
        else:
            result[child.tag] = child_value
    return result


def format_occurs(min_occurs, max_occurs):
    if min_occurs == 1 and max_occurs == 1:
        return "(1)"
    return f"({min_occurs}..{'n' if max_occurs is None else max_occurs})"


def local_tag(tag):
    return tag.split("}")[-1]


def describe_children(element_type, xml_element, lines, depth, type_stack, parent_requirement, chain_intact=True, ancestors_required=True):
    content = getattr(element_type, "content", None)
    if content is None:
        return

    type_id = id(element_type)
    if type_id in type_stack:
        lines.append(f"{'  ' * depth}- ... (type récursif, voir ci-dessus)")
        return

    if content.model != "sequence":
        lines.append(f"{'  ' * depth}- *({content.model} : un seul des éléments suivants)*")

    actual_children = list(xml_element) if xml_element is not None else []
    actual_tags = [local_tag(child.tag) for child in actual_children]
    expected_names = {schema_child.local_name for schema_child in content.iter_elements()}

    for position, tag in enumerate(actual_tags, start=1):
        if tag not in expected_names:
            lines.append(
                f"{'  ' * depth}- `{tag}` (position {position} dans le fichier) "
                f"— **ERREUR** (balise non autorisée à cet emplacement)"
            )

    for expected_position, schema_child in enumerate(content.iter_elements(), start=1):
        occurs = format_occurs(schema_child.min_occurs, schema_child.max_occurs)
        requirement = "R" if schema_child.min_occurs >= 1 else "O"
        positions = [i + 1 for i, tag in enumerate(actual_tags) if tag == schema_child.local_name]

        if positions:
            actual_label = ", ".join(str(p) for p in positions)
            representative = actual_children[positions[0] - 1]
        else:
            actual_label = "absente"
            representative = None

        is_error = requirement == "R" and actual_label == "absente" and parent_requirement == "R" and chain_intact
        error_label = " — **ERREUR**" if is_error else ""

        is_array = schema_child.max_occurs is None or schema_child.max_occurs > 1
        is_tableau_json = (
            is_array and requirement == "R" and actual_label != "absente" and ancestors_required
        )
        tableau_label = " — Tableau JSON" if is_tableau_json else ""

        lines.append(
            f"{'  ' * depth}- `{schema_child.local_name}` {occurs} — ({requirement}) "
            f"({actual_label}, {expected_position}){tableau_label}{error_label}"
        )

        child_chain_intact = chain_intact and not (requirement == "O" and actual_label == "absente")
        child_ancestors_required = ancestors_required and requirement == "R"

        if not schema_child.type.is_simple():
            describe_children(
                schema_child.type, representative, lines, depth + 1,
                type_stack | {type_id}, requirement, child_chain_intact, child_ancestors_required,
            )


def write_order_doc(schema, xml_path, md_path):
    root_xml_element = ET.parse(xml_path).getroot()
    root_local_name = local_tag(root_xml_element.tag)
    root_element = schema.elements[root_local_name]

    lines = [
        f"# Ordonnancement attendu des balises XML pour `{root_local_name}`",
        "",
        f"- `{root_local_name}` — racine du fichier",
    ]
    describe_children(root_element.type, root_xml_element, lines, 1, frozenset(), "R")

    with open(md_path, "w", encoding="utf-8") as md_file:
        md_file.write("\n".join(lines) + "\n")
    print("Ordonnancement des balises généré ici :", md_path)


def run_conversion(xml_path, json_path_arg, xsd_path_arg, work_dir=None, xsd_dir=None, log=print, output_dir=None):
    base = os.path.abspath(work_dir) if work_dir else os.getcwd()

    effective_xsd_dir = os.path.abspath(xsd_dir) if xsd_dir else os.path.join(base, XSD_DIR)

    xsd_path = xsd_path_arg if xsd_path_arg else None
    if xsd_path and not os.path.isabs(xsd_path) and not os.path.dirname(xsd_path):
        xsd_path = os.path.join(effective_xsd_dir, xsd_path)

    if output_dir:
        dest_dir = os.path.abspath(output_dir)
    else:
        dest_dir = os.path.join(base, RESULT_DIR)
    os.makedirs(dest_dir, exist_ok=True)
    json_basename = os.path.splitext(os.path.basename(json_path_arg))[0] + ".json"
    json_path = os.path.join(dest_dir, json_basename)

    if xsd_path:
        schema = xmlschema.XMLSchema11(xsd_path)

        md_path = os.path.join(dest_dir, os.path.splitext(os.path.basename(json_path_arg))[0] + "_ordre.md")
        write_order_doc(schema, xml_path, md_path)

        errors = list(schema.iter_errors(xml_path))
        if errors:
            log("Le fichier XML n'est pas valide selon le XSD. Détails des erreurs :")
            for error in errors:
                path = f"Chemin XML : {error.path}" if getattr(error, 'path', None) else ""
                reason = f" -> {error.reason}" if getattr(error, 'reason', None) else ""
                log(f"- {path} {error.message}{reason}")
            return False
        else:
            log("Le fichier XML est valide selon le XSD.")

        data = schema.to_dict(xml_path, process_namespaces=True)
    else:
        log("Aucun fichier XSD fourni : conversion sans validation ni typage.")
        tree = ET.parse(xml_path)
        root = tree.getroot()
        data = {root.tag: element_to_dict(root)}

    with open(json_path, "w", encoding="utf-8") as json_file:
        json.dump(data, json_file, indent=2, ensure_ascii=False, default=json_default)

    log("Conversion terminée ! Le fichier JSON est disponible ici : " + json_path)
    return True


# ── Interface graphique ────────────────────────────────────────────────────────

def launch_gui():
    import tkinter as tk
    from tkinter import filedialog, messagebox, scrolledtext

    root = tk.Tk()
    root.title("Convertisseur XML → JSON")
    root.resizable(False, False)

    pad = {"padx": 8, "pady": 4}

    def make_row(row, label):
        tk.Label(root, text=label, anchor="w").grid(row=row, column=0, sticky="w", **pad)
        var = tk.StringVar()
        entry = tk.Entry(root, textvariable=var, width=60)
        entry.grid(row=row, column=1, **pad)
        return var, entry

    # ── Répertoire de travail ──────────────────────────────────────────────────
    workdir_var, _ = make_row(0, "Répertoire de travail :")

    def get_dir(var):
        d = var.get().strip()
        return d if d and os.path.isdir(d) else None

    def browse_dir(var, title):
        path = filedialog.askdirectory(title=title)
        if path:
            var.set(path)

    tk.Button(root, text="Parcourir…", command=lambda: browse_dir(workdir_var, "Répertoire de travail")).grid(row=0, column=2, **pad)

    # ── Répertoire XSD ─────────────────────────────────────────────────────────
    xsddir_var, _ = make_row(1, "Répertoire XSD :")
    tk.Button(root, text="Parcourir…", command=lambda: browse_dir(xsddir_var, "Répertoire des fichiers XSD")).grid(row=1, column=2, **pad)

    # ── Fichier XML ────────────────────────────────────────────────────────────
    xml_var, _ = make_row(2, "Fichier XML :")

    def browse_xml():
        initial = get_dir(workdir_var) or "/"
        path = filedialog.askopenfilename(
            title="Sélectionner le fichier XML",
            initialdir=initial,
            filetypes=[("Fichiers XML", "*.xml"), ("Tous les fichiers", "*.*")],
        )
        if path:
            xml_var.set(path)
            if not json_var.get():
                json_var.set(os.path.basename(path))

    tk.Button(root, text="Parcourir…", command=browse_xml).grid(row=2, column=2, **pad)

    # ── Nom JSON ───────────────────────────────────────────────────────────────
    json_var, json_entry = make_row(3, "Nom fichier JSON :")

    def compute_result_path():
        workdir = get_dir(workdir_var)
        base = os.path.abspath(workdir) if workdir else os.getcwd()
        return os.path.join(base, RESULT_DIR) + os.sep

    def on_json_focus_in(_):
        status_var.set("Dossier de sortie : " + compute_result_path())

    def on_json_focus_out(_):
        status_var.set("")

    json_entry.bind("<FocusIn>", on_json_focus_in)
    json_entry.bind("<FocusOut>", on_json_focus_out)

    # ── Fichier XSD ────────────────────────────────────────────────────────────
    xsd_var, _ = make_row(4, "Fichier XSD (optionnel) :")

    def browse_xsd():
        xsd_dir = get_dir(xsddir_var)
        initial = xsd_dir or get_dir(workdir_var) or "/"
        path = filedialog.askopenfilename(
            title="Sélectionner le fichier XSD",
            initialdir=initial,
            filetypes=[("Fichiers XSD", "*.xsd"), ("Tous les fichiers", "*.*")],
        )
        if path:
            xsd_var.set(path)

    tk.Button(root, text="Parcourir…", command=browse_xsd).grid(row=4, column=2, **pad)

    # ── Zone de log ────────────────────────────────────────────────────────────
    log_box = scrolledtext.ScrolledText(root, width=80, height=14, state="disabled", wrap="word")
    log_box.grid(row=5, column=0, columnspan=3, padx=8, pady=8)

    def log(msg):
        log_box.config(state="normal")
        log_box.insert("end", msg + "\n")
        log_box.see("end")
        log_box.config(state="disabled")
        root.update_idletasks()

    # ── Boutons ────────────────────────────────────────────────────────────────
    def on_convert():
        xml_path = xml_var.get().strip()
        json_name = json_var.get().strip()
        xsd_path = xsd_var.get().strip() or None
        workdir = get_dir(workdir_var)
        xsddir = get_dir(xsddir_var)

        if not xml_path:
            messagebox.showerror("Champ manquant", "Veuillez sélectionner un fichier XML.")
            return
        if not json_name:
            messagebox.showerror("Champ manquant", "Veuillez saisir un nom pour le fichier JSON.")
            return
        if not os.path.isfile(xml_path):
            messagebox.showerror("Fichier introuvable", f"Le fichier XML est introuvable :\n{xml_path}")
            return
        if xsd_path and not os.path.isfile(xsd_path):
            messagebox.showerror("Fichier introuvable", f"Le fichier XSD est introuvable :\n{xsd_path}")
            return

        log_box.config(state="normal")
        log_box.delete("1.0", "end")
        log_box.config(state="disabled")

        try:
            ok = run_conversion(xml_path, json_name, xsd_path, work_dir=workdir, xsd_dir=xsddir, log=log)
            if ok:
                messagebox.showinfo("Succès", "Conversion terminée avec succès !")
            else:
                messagebox.showerror("Erreur de validation", "Le XML ne respecte pas le XSD.\nConsultez le journal ci-dessous.")
        except Exception as exc:
            log(f"ERREUR : {exc}")
            messagebox.showerror("Erreur", str(exc))

    btn_frame = tk.Frame(root)
    btn_frame.grid(row=6, column=0, columnspan=3, pady=(0, 4))
    tk.Button(btn_frame, text="Convertir", command=on_convert, width=20, bg="#0078d4", fg="white",
              font=("Segoe UI", 10, "bold")).pack(side="left", padx=4)
    tk.Button(btn_frame, text="Quitter", command=root.destroy, width=10).pack(side="left", padx=4)

    # ── Barre de statut ────────────────────────────────────────────────────────
    status_var = tk.StringVar(value="")
    tk.Label(root, textvariable=status_var, anchor="w", relief="sunken", bd=1, fg="gray",
             font=("Segoe UI", 8)).grid(row=7, column=0, columnspan=3, sticky="ew", padx=2, pady=(0, 2))

    root.mainloop()


# ── Point d'entrée ─────────────────────────────────────────────────────────────

if __name__ == '__main__':
    if len(sys.argv) == 1:
        launch_gui()
    else:
        parser = argparse.ArgumentParser(description="Convertir un fichier XML en JSON, avec validation et typage optionnels via un XSD.")
        parser.add_argument("xml_path", help="Chemin du fichier XML à convertir")
        parser.add_argument("json_path", help="Nom (ou chemin) du fichier JSON de sortie, enregistré dans le répertoire 'result'")
        parser.add_argument("xsd_path", nargs="?", help="Nom (ou chemin) du fichier XSD")
        parser.add_argument("--work-dir", help="Répertoire de travail (par défaut : répertoire courant)")
        parser.add_argument("--xsd-dir", help="Répertoire des fichiers XSD (par défaut : <work-dir>/xsd/)")
        args = parser.parse_args()

        ok = run_conversion(args.xml_path, args.json_path, args.xsd_path,
                            work_dir=args.work_dir, xsd_dir=args.xsd_dir)
        if not ok:
            sys.exit(1)
