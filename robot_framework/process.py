"""This module contains the main process of the robot."""

from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
from OpenOrchestrator.database.queues import QueueElement
import json
import requests
from office365.sharepoint.client_context import ClientContext


# pylint: disable-next=unused-argument
def process(orchestrator_connection: OrchestratorConnection, queue_element: QueueElement | None = None) -> None:
    """Do the primary process of the robot."""
    orchestrator_connection.log_trace("Running process.")

    queue_data = json.loads(queue_element.data)
    anmodnings_id = queue_data.get("application_uuid")

    # --- Credentials / SharePoint-kontekst ---
    os2forms_user = orchestrator_connection.get_credential("OS2FormsAPI")
    os2forms_url = os2forms_user.username
    os2forms_api_key = os2forms_user.password

    certification = orchestrator_connection.get_credential("SharePointCert")
    api = orchestrator_connection.get_credential("SharePointAPI")
    base_url = f'{orchestrator_connection.get_constant("AarhusKommuneSharePoint").value}/Teams/tea-teamsite12592'

    cert_credentials = {
        "tenant": api.username,
        "client_id": api.password,
        "thumbprint": certification.username,
        "cert_path": certification.password,
    }
    ctx = ClientContext(base_url).with_client_certificate(**cert_credentials)
    ctx.load(ctx.web)
    ctx.execute_query()

    # --- Hent submission fra OS2Forms ---
    url = f"{os2forms_url}tilmelding_til_altinget/submission/{anmodnings_id}"
    headers = {"api-key": os2forms_api_key}
    response = requests.get(url, headers=headers, timeout=60)
    response.raise_for_status()

    payload = response.json()
    # "data" ligger på top-niveau ved siden af "entity" (ikke inde i entity)
    form = payload["data"]["mine_medarbejder_data"]
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