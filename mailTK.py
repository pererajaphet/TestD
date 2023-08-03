import tkinter as tk
from tkinter import filedialog
import mailbox
import csv
import re
import subprocess
import os
from email import message_from_file
import sys

ARCHIVE_FILE = "archive.csv"


def convert_pst_ost_to_mbox(pst_ost_file, output_dir):
    print("Converting PST/OST to MBOX...")
    subprocess.run(['readpst', '-o', output_dir, pst_ost_file], check=True)


def process_mbox(mbox_dir):
    print("Processing MBOX...")
    mbox = mailbox.mbox(mbox_dir)

    data_list = []
    message_cache = {}

    for message in mbox:
        message_id = message['Message-ID']
        data_dict = process_message(message)
        data_dict.update(get_message_tree(message, message_cache))
        data_list.append(data_dict)

        # Ajouter le message actuel au cache
        message_cache[message_id] = message

    return data_list


def process_message(message):
    attribs = ['subject', 'from', 'to', 'cc', 'date']
    data_dict = {}
    for attrib in attribs:
        data_dict[attrib] = message.get(attrib, "N/A")

    if message.get_all('X-Transport') is not None:
        data_dict.update(process_headers(message.get_all('X-Transport')))

    attachments_size = get_attachments_size(message)
    data_dict['attachments_size'] = attachments_size

    return data_dict


def process_headers(header):
    key_pattern = re.compile("^([A-Za-z\-]+:)(.*)$")
    header_data = {}
    for line in header:
        line = line.decode('utf-8')
        if len(line) == 0:
            continue

        reg_result = key_pattern.match(line)
        if reg_result:
            key = reg_result.group(1).strip(":").strip()
            value = reg_result.group(2).strip()
        else:
            value = line

        if key.lower() in header_data:
            if isinstance(header_data[key.lower()], list):
                header_data[key.lower()].append(value)
            else:
                header_data[key.lower()] = [header_data[key.lower()], value]
        else:
            header_data[key.lower()] = value
    return header_data


def get_attachments_size(message):
    attachments_size = 0
    for part in message.walk():
        if part.get_filename() is not None:
            attachments_size += len(part.get_payload(decode=True))
    return attachments_size


def get_message_tree(message, message_cache):
    tree = []
    while True:
        message_id = message['In-Reply-To']
        if not message_id:
            break

        parent_message = message_cache.get(message_id)
        if not parent_message:
            break

        tree.append(parent_message['subject'])
        message = parent_message

    tree.append(message['subject'])
    tree.reverse()
    return {'message_tree': tree, 'message_status': get_message_status(message)}


def get_message_status(message):
    # Vérifier le statut du message (lu/non lu)
    return 'SEEN' if message['Status'] == 'RO' else 'UNSEEN'


def write_data(outfile, data_list):
    print("Writing Report: ", outfile)
    columns = ['subject', 'from', 'to', 'cc', 'date',
               'attachments_size', 'message_tree', 'message_status']
    formatted_data_list = []
    for entry in data_list:
        tmp_entry = {}
        for k, v in entry.items():
            if k not in columns:
                columns.append(k)

            if isinstance(v, list):
                tmp_entry[k] = "; ".join(v)
            else:
                tmp_entry[k] = v
        formatted_data_list.append(tmp_entry)

    with open(outfile, 'w', newline='') as openfile:
        csvfile = csv.DictWriter(openfile, delimiter=';', fieldnames=columns)
        csvfile.writeheader()
        csvfile.writerows(formatted_data_list)


def load_pst_file():
    file_path = filedialog.askopenfilename(
        filetypes=[("Fichier PST", "*.pst")])
    if file_path:
        output_dir = "files"
        convert_pst_ost_to_mbox(file_path, output_dir)
        parsed_data = []
        for root, dirs, files in os.walk(output_dir):
            for file in files:
                mbox_path = os.path.join(root, file)
                mbox_data = process_mbox(mbox_path)
                parsed_data.extend(mbox_data)
        report_path = os.path.join(output_dir, "report.csv")
        write_data(report_path, parsed_data)
        print("PST file loaded and processed. Report saved at:", report_path)
        update_archive(report_path)


def update_archive(report_path):
    if not os.path.exists(ARCHIVE_FILE):
        os.rename(report_path, ARCHIVE_FILE)
    else:
        existing_data = read_archive()
        new_data = read_report(report_path)
        merged_data = merge_data(existing_data, new_data)
        write_data(ARCHIVE_FILE, merged_data)
        # Sauvegarder une copie du rapport actuel
        report_copy_path = report_path.replace(".csv", "_Current.csv")
        os.rename(report_path, report_copy_path)
        print("Data merged and saved in archive.")
        print("Current report saved as:", report_copy_path)


def read_archive():
    with open(ARCHIVE_FILE, 'r') as openfile:
        csvfile = csv.DictReader(openfile, delimiter=';')
        data_list = []
        for row in csvfile:
            data_list.append(row)
        return data_list


def read_report(report_path):
    with open(report_path, 'r') as openfile:
        csvfile = csv.DictReader(openfile, delimiter=';')
        data_list = []
        for row in csvfile:
            data_list.append(row)
        return data_list


def merge_data(existing_data, new_data):
    merged_data = existing_data + new_data
    unique_data = [dict(t) for t in {tuple(d.items()) for d in merged_data}]
    return unique_data


# Création de la fenêtre principale
window = tk.Tk()

# Fonction pour mettre à jour l'étiquette du chemin du fichier PST


def update_path_label(path):
    path_label.config(text="PST File: " + path)

# Fonction pour afficher un message de succès dans la fenêtre


def show_success_message():
    success_label = tk.Label(
        window, text="File loaded and processed successfully!", font=("Arial", 12), fg="green")
    success_label.pack()

# Fonction pour quitter l'application


def exit_application():
    sys.exit()


# Création de l'étiquette pour afficher le chemin du fichier PST
path_label = tk.Label(window, text="PST File: ", font=("Arial", 12))
path_label.pack()

# Création du bouton "Load PST File"
load_button = tk.Button(window, text="Load PST File", command=lambda: [
                        load_pst_file(), show_success_message()], font=("Arial", 12))
load_button.pack()

# Création du bouton "Exit"
exit_button = tk.Button(window, text="Exit",
                        command=exit_application, font=("Arial", 12))
exit_button.pack()

# Configuration de la taille de la fenêtre
window.geometry("400x200")

# Lancement de la boucle principale de l'interface graphique
window.mainloop()
