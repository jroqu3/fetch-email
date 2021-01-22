import os
import re
import email
import imaplib
import datetime
import asyncio
import configparser
from aioimaplib import aioimaplib
# from app.principal.models import Medicamento


class FetchEmail():
    def __init__(self, mail_server, username, password, path_email):
        self.list_response_pattern = re.compile(
            r'\((?P<flags>.*?)\) "(?P<delimiter>.*)" (?P<name>.*)'
        )

        print("Connecting to...", mail_server)
        self.connection = imaplib.IMAP4_SSL(mail_server)
        print("Logging in as...", username)
        self.connection.login(username, password)

        typ, mbox_data = self.connection.list(pattern=path_email)
        flags, delimiter, mbox_name = self.parse_list_response(mbox_data[0])
        self.connection.select('"{}"'.format(mbox_name), readonly=False)

    def close_connection(self):
        self.connection.close()

    def parse_list_response(self, line):
        match = self.list_response_pattern.match(line.decode("utf-8"))
        flags, delimiter, mailbox_name = match.groups()
        mailbox_name = mailbox_name.strip('"')
        return (flags, delimiter, mailbox_name)

    def save_attachment(self, msg, download_folder):
        ok, att_path = False, "No attachment csv found."
        for part in msg.walk():
            if part.get_content_maintype() == "multipart":
                continue
            if part.get("Content-Disposition") is None:
                continue

            self.download_folder = download_folder
            self.filename = part.get_filename()
            ext = self.filename.split(".")[-1]

            if ext not in ("csv"):
                continue

            att_path = os.path.join(self.download_folder, self.filename)

            if not os.path.isfile(att_path):
                fp = open(att_path, "wb")
                fp.write(part.get_payload(decode=True))
                fp.close()
            ok = True
        return ok, att_path

    def fetch_unread_messages(self):
        emails = []
        (result, messages) = self.connection.search(None, '(UnSeen SUBJECT "prueba")')
        if result == "OK":
            for message in messages[0].decode("utf-8").split(" "):
                try:
                    ret, data = self.connection.fetch(message, "(RFC822)")

                    msg = email.message_from_bytes(data[0][1])

                    print("=" * 50)
                    for header in ["date", "subject", "to", "from"]:
                        print("{:^8}: {}".format(
                            header.upper(), msg[header]
                        ))
                    print("=" * 50)

                    if not isinstance(msg, str):
                        emails.append(msg)
                    response, data = self.connection.store(message, "+FLAGS", "\\Seen")
                except:
                    print("No new emails to read.")
                finally:
                    self.close_connection()
                    exit()
            return emails

        self.error = "Failed to retreive emails."
        return emails

    def load_in_database(self):
        print("Saving in database...")
        with open("{}/{}".format(self.download_folder, self.filename), mode="r") as csv_file:
            csv_reader = csv.DictReader(csv_file)
            total_rows = 0
            total_created = 0
            total_updated = 0
            Medicamento.objects.all().update(estado="D")

            for row in csv_reader:
                try:
                    updated = False
                    total_rows += 1
                    medicamento_obj, created = Medicamento.objects.get_or_create(clave=row["ID"])
                    if medicamento_obj.nombre != row["STRNOMBRE"]:
                        medicamento_obj.nombre = row["STRNOMBRE"]
                        updated = True
                    if medicamento_obj.sustancia_activa != row["STRSUSTANCIAACTIVA"]:
                        medicamento_obj.sustancia_activa = row["STRSUSTANCIAACTIVA"]
                        updated = True
                    if medicamento_obj.gpi != row["STRGPI"]:
                        medicamento_obj.gpi = row["STRGPI"]
                        updated = True
                    if medicamento_obj.controlado != row["ICONTROLADO"]:
                        medicamento_obj.controlado = row["ICONTROLADO"]
                        updated = True
                    if medicamento_obj.biologico != row["IBIOLOGICO"]:
                        medicamento_obj.biologico = row["IBIOLOGICO"]
                        updated = True
                    if medicamento_obj.antibiotico != row["IANTIBIOTICO"]:
                        medicamento_obj.antibiotico = row["IANTIBIOTICO"]
                        updated = True
                    if updated:
                        total_updated += 1
                        medicamento_obj.fecha_modificacion = datetime.datetime.now()
                        medicamento_obj.save()

                    if created:
                        total_created += 1
                        print("Object {} created".format(row["ID"]))
                except Exception as e:
                    print("Exception in load_in_database (ID: {}) => {}".format(row["ID"], e))
            print("Total read => {}".format(total_rows))
            print("Total created => {}".format(total_created))
            print("Total updated => {}".format(total_updated))
        print("Data updated...")


config = configparser.ConfigParser()
config.read([os.path.expanduser("~/.pymotw")])
hostname = config.get("server", "hostname")
username = config.get("account", "username")
password = config.get("account", "password")
path_email = config.get("path", "email")
download_folder = config.get("path", "download")

fetch_email = FetchEmail(hostname, username, password, path_email)
unseen_emails = fetch_email.fetch_unread_messages()
for msg in unseen_emails:
    ok, attach = fetch_email.save_attachment(msg, download_folder)
    if ok:
        print("Attachment saved in...", attach)
        # fetch_email.load_in_database()
    else:
        print(attach)
