# Tilmelding til Altinget – OpenOrchestrator Robot

An [OpenOrchestrator](https://github.com/itk-dev-rpa/OpenOrchestrator) robot that registers
employee sign-ups for *Altinget* in a SharePoint list. When an employee submits the
**"Tilmelding til Altinget"** form in OS2Forms, the robot picks up the submission, reads the
employee data, and creates a corresponding row in a SharePoint list.

## How it works

The robot runs as a queue-based process in openorchestrator. Each queue element references a single OS2Forms
submission, and the robot processes it as follows:

1. **Read the queue element.** The element's `data` contains an `application_uuid` — the ID
   of the OS2Forms submission to process.
2. **Fetch the submission** from the OS2Forms REST API
   (`<api-url>/tilmelding_til_altinget/submission/<application_uuid>`), authenticated with an
   API key. The employee fields live under `data.mine_medarbejder_data`.
3. **Authenticate to SharePoint** using an app-only (client certificate) login against the
   team site.
4. **Map the form fields** to the internal SharePoint column names (see table below).
5. **Create a new list item** in *Tilmeldte medarbejdere*. On any failure the error is logged
   and re-raised, so the queue element is marked **FAILED** rather than silently completing.

## Field mapping

The form fields from `mine_medarbejder_data` are mapped to the **internal** SharePoint column
names (not the display names):

| OS2Forms field         | SharePoint column (internal name)  | Example value        |
| ---------------------- | ---------------------------------- | -------------------- |
| `name`                 | `Title`                            | Employee name        |
| `az`                   | `Az_x002d_ident`                   | AZ identifier        |
| `organisation_enhed`   | `Afdeling`                         | Digital Udvikling    |
| `organisation_niveau_2`| `Organisatoriskenhedovermedarbejd` | Digitalisering MTM   |
| `magistrat`            | `Magistratsafdeling`               | Teknik og Miljø      |

## Requirements

- Python **3.11+** (required by OpenOrchestrator 3.x)
- Dependencies (see `pyproject.toml`):
  - `OpenOrchestrator`
  - `Office365-REST-Python-Client`
  - `requests`

```toml
[project]
requires-python = ">=3.11"
dependencies = [
    "OpenOrchestrator>=3.0.0",
    "Office365-REST-Python-Client>=2.6.2",
    "requests>=2.34.2",
]
```

## Configuration

The robot reads its credentials and settings from OpenOrchestrator. These must exist before
the robot is run:

| Type       | Name                      | Contents                                          |
| ---------- | ------------------------- | ------------------------------------------------- |
| Credential | `OS2FormsAPI`             | username = API base URL, password = API key       |
| Credential | `SharePointAPI`           | username = tenant, password = client ID           |
| Credential | `SharePointCert`          | username = certificate thumbprint, password = path to certificate (PEM) |
| Constant   | `AarhusKommuneSharePoint` | Base SharePoint URL                               |

**Target site and list**

- Site: `<AarhusKommuneSharePoint>/Teams/tea-teamsite12592`
- List: `Tilmeldte medarbejdere`

The Azure AD app behind the certificate must have **write** permission to the target site
(e.g. `Sites.Selected` granted with Write, or broader). Read-only access is not enough to
create list items.

## Setup

1. Install the dependencies (`pip install -e .` or your usual workflow).
2. Create the credentials and constant listed above in OpenOrchestrator.
3. Make sure the SharePoint app registration has write access to the target site.
4. Configure a trigger that enqueues OS2Forms submissions and points OpenOrchestrator at this
   process.

## Known gotchas

- **Internal column names are not display names.** SharePoint encodes spaces and special
  characters at column-creation time (`-` becomes `_x002d_`, `ø` becomes `_x00f8_`, etc.) and
  truncates internal names at 32 characters — which is why `Organisatoriskenhedovermedarbejd`
  looks cut off but is correct. Use the `get_internal_column_names()` helper to list the
  current internal names before changing the mapping.

- **There are three columns displaying as "Magistratsafdeling"** (`Magistratsafdeling`,
  `Magistratsafdeling0`, `Magistratsafdeling1`). Writing to the wrong one succeeds *without an
  error* but lands the value in a column the view doesn't show. Confirm which internal name
  backs the visible column before relying on the mapping.

- **Empty form fields** come back as empty strings (`""`), not `null`. The robot skips empty
  values so blank fields are not written.
