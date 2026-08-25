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

def send_forkert_mag_mail(az, orchestrator_connection):
    SMTP_SERVER = "smtp.adm.aarhuskommune.dk"
    SMTP_PORT = 25
    SCREENSHOT_SENDER = "PersonaleAktindsigtssag@aarhus.dk"
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

        # --- MultiChoice-felter: SharePoint REST API forventer {'results': [...]} ---
        for felt in ("booking_af_cykelbane_r_", "booking_af_cykler"):
            vaerdi = form.get(felt)
            if vaerdi:
                form[felt] = {"results": vaerdi if isinstance(vaerdi, list) else [vaerdi]}

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
