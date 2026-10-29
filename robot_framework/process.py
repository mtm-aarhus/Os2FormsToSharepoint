"""This module contains the main process of the robot."""

from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
from OpenOrchestrator.database.queues import QueueElement
import json
import requests
from office365.sharepoint.client_context import ClientContext
import pyodbc
import smtplib
from email.message import EmailMessage
from datetime import datetime
from zoneinfo import ZoneInfo
import re
import os
from urllib.parse import urlparse, urljoin, unquote
from office365.sharepoint.attachments.creation_information import AttachmentCreationInformation

COPENHAGEN_TZ = ZoneInfo("Europe/Copenhagen")


def local_to_sharepoint_utc(date_str: str, time_str: str = "00:00:00") -> str:
    """
    Konverterer en lokal dansk dato/tid (Europe/Copenhagen, inkl. sommer-/vintertid)
    til en UTC ISO8601-streng, som SharePoints DateTime-felter forventer.
    Uden denne konvertering risikerer man at datoen forskydes en dag ved
    visning, fordi SharePoint gemmer DateTime-felter internt i UTC.
    """
    naive_dt = datetime.strptime(f"{date_str} {time_str}", "%Y-%m-%d %H:%M:%S")
    local_dt = naive_dt.replace(tzinfo=COPENHAGEN_TZ)
    utc_dt = local_dt.astimezone(ZoneInfo("UTC"))
    return utc_dt.strftime("%Y-%m-%dT%H:%M:%SZ")


def format_time_text(time_str: str) -> str:
    """Formatterer et tidspunkt fra formularen ('11:00:00') til 'HH:MM' til Text-feltet."""
    return datetime.strptime(time_str, "%H:%M:%S").strftime("%H:%M")


def os2forms_tid_to_sharepoint_utc(value) -> str | None:
    """Konverterer created/completed fra OS2Forms til SharePoints UTC-format.
    Kommer som ISO-streng med tidszone ('2026-10-29T09:46:06+00:00'),
    men Unix-sekunder håndteres også."""
    if not value or str(value) == "0":
        return None
    value = str(value)
    if value.isdigit():
        dt = datetime.fromtimestamp(int(value), ZoneInfo("UTC"))
    else:
        dt = datetime.fromisoformat(value)
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=COPENHAGEN_TZ)
    return dt.astimezone(ZoneInfo("UTC")).strftime("%Y-%m-%dT%H:%M:%SZ")


def tekst(value) -> str | None:
    """Trimmer en formularværdi og gør tomme strenge til None."""
    if value is None:
        return None
    return str(value).strip() or None


def sp_link(url: str) -> dict:
    """Værdi til en SharePoint-kolonne af typen 'Hyperlink eller billede'."""
    url = url.replace(" ", "%20")[:255]
    return {"__metadata": {"type": "SP.FieldUrlValue"}, "Url": url, "Description": url}


def hent_os2forms_fil(form: dict, fil_id: str, os2forms_url: str, api_key: str) -> tuple[str, bytes]:
    """Henter en uploadet fil fra OS2Forms. Returnerer (filnavn, indhold)."""
    headers = {"api-key": api_key}
    filnavn = None

    # 1) OS2Forms' REST-modul lægger normalt fil-URL'en under data.linked
    info = ((form.get("linked") or {}).get("upload_dokumenter") or {}).get(fil_id) or {}
    fil_url = info.get("url")
    print(f"[P8DEBUG] fil {fil_id}: linked-felter={sorted(info.keys())}")

    # 2) Ellers slås filen op via entity/file
    if not fil_url:
        r = requests.get(f"{os2forms_url}entity/file/{fil_id}?_format=json", headers=headers, timeout=60)
        print(f"[P8DEBUG] fil {fil_id}: entity/file -> {r.status_code} {r.headers.get('Content-Type')}")
        r.raise_for_status()
        entity = r.json()
        print(f"[P8DEBUG] fil {fil_id}: entity-nøgler={sorted(entity.keys())}")
        fil_url = entity["uri"][0]["url"]
        filnavn = entity.get("filename", [{}])[0].get("value")

    fil_url = urljoin(os2forms_url, fil_url)  # relativ /system/files/... -> fuld URL
    r = requests.get(fil_url, headers=headers, timeout=120)
    content_type = r.headers.get("Content-Type", "")
    print(f"[P8DEBUG] fil {fil_id}: download {fil_url} -> {r.status_code} {content_type} {len(r.content)} bytes")
    if not r.ok:
        print(f"[P8DEBUG] fil {fil_id}: svar-body: {r.text[:500]!r}")
    r.raise_for_status()

    # Mangler API-brugeren adgang til private filer, svarer Drupal med en loginside og status 200
    if "text/html" in content_type:
        print(f"[P8DEBUG] fil {fil_id}: HTML i stedet for fil: {r.text[:300]!r}")
        raise RuntimeError(f"Download af fil {fil_id} gav HTML i stedet for filen – tjek API-brugerens adgang til private filer.")

    filnavn = filnavn or unquote(os.path.basename(urlparse(fil_url).path)) or f"fil_{fil_id}"
    return filnavn, r.content

def send_forkert_mag_mail(az, orchestrator_connection):
    SMTP_SERVER = "smtp.adm.aarhuskommune.dk"
    SMTP_PORT = 25
    SCREENSHOT_SENDER = "altingetopretter@aarhus.dk"
    html_failed = f"""
    <html>
    <body>
        <p>Der er oprettet adgang til altinget fra en bruger udenfor MTM. Az er {az}</p>
    </body>
    </html>
    """
    # Create the email message
    UdviklerMail = orchestrator_connection.get_constant('balas').value

    msg_failed = EmailMessage()
    msg_failed['To'] = UdviklerMail
    msg_failed['From'] = SCREENSHOT_SENDER
    msg_failed['Subject'] = "Altinget oprettelse - forkert magistratsafdeling"
    msg_failed.set_content("Please enable HTML to view this message.")
    msg_failed.add_alternative(html_failed, subtype='html')

    # Send the email using SMTP
    try:
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as smtp:
            smtp.send_message(msg_failed)
    except Exception as e:
        orchestrator_connection.log_info(f"Failed to send error email: {e}")


def get_magistrat_from_fdw(orchestrator_connection: OrchestratorConnection, az_ident: str) -> str | None:
    """Slår magistratsafdeling op i FDW baseret på az-ident, hvis feltet er tomt i formularen."""
    az_ident = az_ident.upper()

    sql_server_f = orchestrator_connection.get_constant("sqlserverf").value
    conn_string_f = f"DRIVER={{SQL Server}};SERVER={sql_server_f};DATABASE=FDW;Trusted_Connection=yes;"

    query = """
        SELECT Niveau2_Navn
        FROM FDW.pdb.PersonLight_udvidet
        WHERE Azident = ?
    """

    with pyodbc.connect(conn_string_f, timeout=30) as conn:
        with conn.cursor() as cursor:
            cursor.execute(query, az_ident)
            row = cursor.fetchone()
            return row[0] if row and row[0] else None
        
# pylint: disable-next=unused-argument
def process(orchestrator_connection: OrchestratorConnection, queue_element: QueueElement | None = None) -> None:
    """Do the primary process of the robot."""
    orchestrator_connection.log_trace("Running process.")

    queue_data = json.loads(queue_element.data)
    anmodnings_id = queue_data.get("application_uuid")
    formular_titel = queue_data.get("formular")

    # --- Credentials / SharePoint-kontekst ---
    os2forms_user = orchestrator_connection.get_credential("OS2FormsAPI")
    os2forms_url = os2forms_user.username
    os2forms_api_key = os2forms_user.password

    certification = orchestrator_connection.get_credential("SharePointCert")
    api = orchestrator_connection.get_credential("SharePointAPI")
    base_url = f'{orchestrator_connection.get_constant("AarhusKommuneSharePoint").value}'
    altinget_endelse = '/Teams/tea-teamsite12592'
    cykel_endelse = '/Teams/tea-teamsite11485'

    cert_credentials = {
        "tenant": api.username,
        "client_id": api.password,
        "thumbprint": certification.username,
        "cert_path": certification.password,
    }
    if  formular_titel.strip().lower() == "Tilmelding til Altinget".lower():
        ctx = ClientContext(f'{base_url}{altinget_endelse}').with_client_certificate(**cert_credentials)
        ctx.load(ctx.web)
        ctx.execute_query()

        # --- Hent submission fra OS2Forms ---
        url = f"{os2forms_url}tilmelding_til_altinget/submission/{anmodnings_id}"
        headers = {"api-key": os2forms_api_key}
        response = requests.get(url, headers=headers, timeout=60)
        response.raise_for_status()

        payload = response.json()
        try:
            form = payload["data"]["mine_medarbejder_data"]
        except Exception as e:
            orchestrator_connection.log_error(f'Ingen data fundet {e}')
            return
        if not form.get("magistrat"):
            az_ident = form.get("az")
            magistrat = get_magistrat_from_fdw(orchestrator_connection, az_ident)
            if magistrat != "Teknik og Miljø":
                send_forkert_mag_mail(az_ident, orchestrator_connection= orchestrator_connection)
            if magistrat:
                form["magistrat"] = magistrat
                orchestrator_connection.log_info(f"Magistrat var tom, hentet fra FDW: {magistrat}")
            else:
                orchestrator_connection.log_info("Magistrat var tom, og kunne ikke findes i FDW.")
        submission_uuid = payload["entity"]["uuid"][0]["value"]  # global unik – til evt. dedup

        # Venstre side = de PRÆCISE nøgler fra form-JSON'en.
        # Højre side  = de interne SharePoint-kolonnenavne.
        column_mapping = {
            "name":                  "Title",
            "az":                    "Az_x002d_ident",
            "organisation_enhed":    "Afdeling",                          
            "organisation_niveau_2": "Organisatoriskenhedovermedarbejd",  
            "magistrat":             "Magistratsafdeling1",                
        }

        add_item_to_sharepoint_list(
            orchestrator_connection=orchestrator_connection,
            client=ctx,
            list_title="Tilmeldte medarbejdere",
            data=form,
            column_mapping=column_mapping,
        )
    elif formular_titel.strip().lower() == "Blanket: Booking af Cykeløen".lower():
        ctx = ClientContext(f'{base_url}{cykel_endelse}').with_client_certificate(**cert_credentials)
        ctx.load(ctx.web)
        ctx.execute_query()

        url = f"{os2forms_url}cykeludlaan_i_mobilitet/submission/{anmodnings_id}"
        headers = {"api-key": os2forms_api_key}
        response = requests.get(url, headers=headers, timeout=60)
        response.raise_for_status()

        payload = response.json()
        form = payload["data"]
        submission_uuid = payload["entity"]["uuid"][0]["value"]  # global unik – til dedup

        # --- Beregnet Title: institution + dato ---
        institution = form.get("hvilken_skole_boernehave_eller_forening_booker_du_for_skriv_navn", "")
        dato_raw = form.get("bookingdato", "")
        form["titel_beregnet"] = f"{institution} - {dato_raw}".strip(" -")

        # --- Dato (DateTime): kombinér dato + starttid, konvertér til UTC ---
        starttid_raw = form.get("starttidspunkt_for_booking_", "00:00:00")
        sluttid_raw = form.get("sluttidspunkt_for_booking", "00:00:00")
        if dato_raw:
            form["dato_beregnet"] = local_to_sharepoint_utc(dato_raw, starttid_raw)

        # --- Starttid / Sluttid (Text): pæn visning uden sekunder ---
        if starttid_raw:
            form["starttidspunkt_for_booking_"] = format_time_text(starttid_raw)
        if sluttid_raw:
            form["sluttidspunkt_for_booking"] = format_time_text(sluttid_raw)

        # --- Boolean-konvertering: formularen sender 'Ja'/'Nej' som streng ---
        form["oensker_cykelundervisning_hvis_muligt_1"] = (
            form.get("oensker_cykelundervisning_hvis_muligt_1") == "Ja"
        )

        for felt in ("booking_af_cykelbane_r_", "booking_af_cykler"):
            vaerdi = form.get(felt)

            if isinstance(vaerdi, list):
                form[felt] = {"results": vaerdi}
            elif vaerdi:
                form[felt] = {"results": [vaerdi]}
            else:
                form[felt] = {"results": []}

        # --- Dedup-nøgle til idempotens ved retry af samme submission ---
        form["submission_uuid"] = submission_uuid

        column_mapping = {
            "titel_beregnet":                  "Title",
            "dit_fulde_navn":                  "Fulde_navn",
            "telefonnummer":                    "Telefonnummer",
            "e_mailadresse":    "Mailadresse",
            "hvilken_skole_boernehave_eller_forening_booker_du_for_skriv_navn": "Institution_navn",
            "klassetrin_aldersgruppe":             "Klassetrin",
            "antal_deltagende_boern_": "Antal_deltagere",
            "dato_beregnet": "Dato",
            "starttidspunkt_for_booking_": "Starttid",
            "sluttidspunkt_for_booking": "Sluttid",
            "booking_af_cykelbane_r_": "Cykelbaner",
            "booking_af_cykler": "Cykler",
            "oensker_cykelundervisning_hvis_muligt_1": "Cykelundervisning",
            "kommentarer_til_din_booking": "Kommentar_til_booking",
            "submission_uuid": "ID_os2forms",
        }

        add_item_to_sharepoint_list(
            orchestrator_connection=orchestrator_connection,
            client=ctx,
            list_title="OS2data",
            data=form,
            column_mapping=column_mapping,
            dedup_field="ID_os2forms",
            dedup_value=submission_uuid,
        )

    elif formular_titel.strip().lower() == "jordforurening":
        # §8-ansøgning. Skriver til fire forbundne lister på Jordportalen:
        # P8Ansogninger + P8Adresser, P8Kontakter og P8Vedhaeftninger, der hver har en
        # Lookup ("Ansogning", sættes som AnsogningId) til ansøgningen.
        # Bemærk lille t i /teams/ her.
        ctx = ClientContext(f'{base_url}/teams/Jordportalen').with_client_certificate(**cert_credentials)
        ctx.load(ctx.web)
        ctx.execute_query()

        # Maskinnavnet er afkortet af Drupal ved 32 tegn – "_beb" er ikke en tastefejl
        url = f"{os2forms_url}ansoegning_om_ss8_tilladelse_beb/submission/{anmodnings_id}"
        headers = {"api-key": os2forms_api_key}
        response = requests.get(url, headers=headers, timeout=60)
        response.raise_for_status()

        payload = response.json()
        form = payload["data"]
        entity = payload["entity"]
        submission_uuid = entity["uuid"][0]["value"]
        sid = (entity.get("sid") or [{}])[0].get("value")
        serial = (entity.get("serial") or [{}])[0].get("value")
        print(f"[P8DEBUG] data-nøgler: {sorted(form.keys())}")
        print(f"[P8DEBUG] upload_dokumenter: {form.get('upload_dokumenter')!r}")

        # --- Dedup: findes ansøgningen allerede, er den skrevet færdig ---
        eksisterende = (
            ctx.web.lists.get_by_title("P8Ansogninger")
            .items.filter(f"SubmissionUUID eq '{submission_uuid}'")
            .top(1).get().execute_query()
        )
        if len(eksisterende) > 0:
            orchestrator_connection.log_info(f"§8: {submission_uuid} findes allerede – springer over.")
            return

        # --- Adresser (ubegrænset antal) ---
        adresser = []
        for blok in form.get("adresser") or []:
            adresse = tekst(blok.get("vaelg_adresse"))
            matrikel = tekst(blok.get("mat"))
            lokalitet = tekst(blok.get("lokalitets_nummer"))
            if adresse or matrikel or lokalitet:
                adresser.append({
                    "Title": adresse or "(uden adresse)",
                    "Adresse": adresse,
                    "Matrikel": matrikel,
                    "LokalitetsNummer": lokalitet,
                })

        # --- Kontakter: blanketten er flad, hver part har sit eget suffiks ---
        flere_grundejere = form.get("er_der_flere_grundejere") == "Ja"
        bygherre_er_grundejer = form.get("er_bygherre_den_samme_som_grundejer") == "Ja"
        parter = [("", "Grundejer")]
        if flere_grundejere:
            parter.append(("_2", "Grundejer"))   # grundejer nr. 2 har intet navnefelt
        if not bygherre_er_grundejer:
            parter.append(("_byg", "Bygherre"))
        parter.append(("_raad", "Rådgiver"))

        kontakter = []
        for suffix, rolle in parter:
            kontakt = {
                "Navn": tekst(form.get(f"navn_kontaktperson{suffix}")),
                "Firma": tekst(form.get(f"evt_firma{suffix}")),
                "CVR": tekst(form.get(f"cvr{suffix}")),
                # Bygherrens mailfelt har dobbelt underscore i blanketten
                "Email": tekst(form.get("e_mail__byg" if suffix == "_byg" else f"e_mail{suffix}")),
                "Telefon": tekst(form.get(f"telefon_nr{suffix}")),
                "Adresse": tekst(form.get(f"adresse{suffix}")),
            }
            if not any(kontakt.values()):
                continue
            kontakt["KontaktType"] = rolle
            kontakt["Title"] = kontakt["Navn"] or kontakt["Firma"] or rolle
            kontakt["ErUdfylder"] = rolle == form.get("ansoegning_indsendt_af") and suffix != "_2"
            kontakter.append(kontakt)

        # --- Uploadede filer: kommer som fil-id'er (én fil kan komme som streng i stedet for liste) ---
        fil_ids = form.get("upload_dokumenter") or []
        if not isinstance(fil_ids, list):
            fil_ids = [fil_ids]
        fil_ids = [str(f) for f in fil_ids if f]

        # --- Hovedrækken ---
        adresse_tekster = [a["Adresse"] for a in adresser if a["Adresse"]]
        if adresse_tekster:
            titel = adresse_tekster[0]
            if len(adresse_tekster) > 1:
                titel += f" (+{len(adresse_tekster) - 1} flere)"
        else:
            titel = f"§8 – {serial}"
        grundejere = [k["Navn"] or k["Firma"] for k in kontakter if k["KontaktType"] == "Grundejer"]
        dato = form.get("vaelg_dato_for_ansoegning")
        os2formsurl_basic = os2forms_url.rsplit('/da')[0]

        ansogning = {
            "Title": titel,
            "SubmissionUUID": submission_uuid,
            "SubmissionSerial": int(serial) if serial else None,
            "SubmissionSid": int(sid) if sid else None,
            "OS2FormsUrl": sp_link(f"{os2formsurl_basic}/admin/structure/webform/submission/{sid}") if sid else None,
            "Udfylder": form.get("udfylder"),
            "IndsendtAf": form.get("ansoegning_indsendt_af"),
            "AnsogningsDato": local_to_sharepoint_utc(dato) if dato else None,
            "Bemaerkninger": form.get("bemaerkninger"),
            "ModtagetDato": os2forms_tid_to_sharepoint_utc((entity.get("created") or [{}])[0].get("value")),
            "AfsluttetDato": os2forms_tid_to_sharepoint_utc((entity.get("completed") or [{}])[0].get("value")),
            "FlereGrundejere": flere_grundejere,
            "BygherreSammeSomGrundejer": bygherre_er_grundejer,
            "Status": "Ny",  # sagsbehandlerens felt – sættes kun ved oprettelse
            "AntalAdresser": len(adresser),
            "AntalKontakter": len(kontakter),
            "AntalVedhaeftninger": len(fil_ids),
            "AdresserTekst": "\n".join(adresse_tekster) or None,
            "Grundejere": "\n".join(grundejere) or None,
        }

        parent = add_item_to_sharepoint_list(
            orchestrator_connection=orchestrator_connection,
            client=ctx,
            list_title="P8Ansogninger",
            data=ansogning,
        )

        # --- Filer og detaljerækker. Fejler noget, slettes det oprettede igen – ellers ville
        # dedup-tjekket ved næste retry springe ansøgningen over med manglende data. ---
        oprettet = [parent]
        try:
            vedhaeftninger = []
            brugte_navne = set()
            for fil_id in fil_ids:
                filnavn, indhold = hent_os2forms_fil(form, fil_id, os2forms_url, os2forms_api_key)
                filnavn = re.sub(r'[\\/:*?"<>|#%]', "_", filnavn).strip() or f"fil_{fil_id}"
                if filnavn.lower() in brugte_navne:
                    filnavn = f"{fil_id}_{filnavn}"
                brugte_navne.add(filnavn.lower())

                attachment = parent.attachment_files.add(
                    AttachmentCreationInformation(filnavn, indhold)
                ).execute_query()
                server_url = attachment.properties.get("ServerRelativeUrl")
                print(f"[P8DEBUG] vedhæftet '{filnavn}' ({len(indhold)} bytes) -> {server_url}")

                vedhaeftninger.append({
                    "Title": filnavn,
                    "FilId": fil_id,
                    "Filnavn": filnavn,
                    "FilUrl": sp_link(f"{base_url}{server_url}") if server_url else None,
                })

            for list_title, raekker in (
                ("P8Adresser", adresser),
                ("P8Kontakter", kontakter),
                ("P8Vedhaeftninger", vedhaeftninger),
            ):
                for raekke in raekker:
                    raekke["AnsogningId"] = parent.id
                    raekke["SubmissionUUID"] = submission_uuid
                    oprettet.append(add_item_to_sharepoint_list(
                        orchestrator_connection=orchestrator_connection,
                        client=ctx,
                        list_title=list_title,
                        data=raekke,
                    ))
        except Exception:
            orchestrator_connection.log_error(f"§8: fejl – sletter {len(oprettet)} oprettede rækker for {submission_uuid}.")
            for item in reversed(oprettet):
                try:
                    item.delete_object().execute_query()
                except Exception as e:  # pylint: disable=broad-exception-caught
                    orchestrator_connection.log_error(f"§8: kunne ikke slette item {item.id}: {e}")
            raise

        orchestrator_connection.log_info(
            f"§8: {submission_uuid} skrevet som item {parent.id} med {len(adresser)} adresser, "
            f"{len(kontakter)} kontakter og {len(vedhaeftninger)} filer."
        )


def get_internal_column_names(client: ClientContext, list_title: str) -> dict:
    """Returnerer {visningsnavn: internt navn} for skrivbare kolonner.
    Kør én gang for at finde de korrekte interne API-navne."""
    sp_list = client.web.lists.get_by_title(list_title)
    fields = sp_list.fields.get().execute_query()
    result = {}
    for f in fields:
        p = f.properties
        if not p.get("Hidden") and not p.get("ReadOnlyField"):
            result[p.get("Title")] = p.get("InternalName")
            print(f"{p.get('Title'):<35} -> {p.get('InternalName'):<35} ({p.get('TypeAsString')})")
    return result


def add_item_to_sharepoint_list(
    orchestrator_connection: OrchestratorConnection,
    client: ClientContext,
    list_title: str,
    data: dict,
    column_mapping: dict = None,
    dedup_field: str = None,
    dedup_value: str = None,
):
    """
    Opretter en ny række i en SharePoint-liste.

    Args:
        orchestrator_connection: Bruges til logning.
        client:         Autentificeret ClientContext.
        list_title:     Den præcise titel på SharePoint-listen.
        data:           Dict med form-data (flade nøgler).
        column_mapping: Mapper form-nøgler -> interne SharePoint-kolonnenavne.
                        Venstre side SKAL matche nøglerne i `data`.
        dedup_field:    Valgfrit. Internt kolonnenavn der bruges til at tjekke
                        for eksisterende række (idempotens).
        dedup_value:    Valgfrit. Værdien der slås op i dedup_field.

    Returns:
        Det oprettede (eller eksisterende) list item.
    """
    try:
        sp_list = client.web.lists.get_by_title(list_title)

        # --- Dedup: undgå dubletter ved retry af samme submission ---
        if dedup_field and dedup_value is not None:
            existing = (
                sp_list.items
                .filter(f"{dedup_field} eq '{dedup_value}'")
                .top(1)
                .get()
                .execute_query()
            )
            if len(existing) > 0:
                orchestrator_connection.log_info(
                    f"⏭️ Række findes allerede ({dedup_field}={dedup_value}, "
                    f"ID {existing[0].id}) – springer over."
                )
                return existing[0]

        # --- Byg det dict der sendes til SharePoint ---
        if column_mapping:
            item_properties = {
                sp_col: data[json_key]
                for json_key, sp_col in column_mapping.items()
                if json_key in data and data[json_key] not in (None, "")
            }
        else:
            item_properties = {k: v for k, v in data.items() if v not in (None, "")}

        new_item = sp_list.add_item(item_properties).execute_query()

        orchestrator_connection.log_info(
            f"✅ Ny række oprettet i '{list_title}' med ID: {new_item.id}"
        )
        return new_item

    except Exception as e:
        orchestrator_connection.log_error(
            f"❌ Fejl ved oprettelse af række i '{list_title}': {str(e)}"
        )
        raise