from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
from OpenOrchestrator.database.queues import QueueElement
from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext
from unittest.mock import MagicMock
from robot_framework.process import process
import os
import json

orchestrator_connection = OrchestratorConnection('os2forms2sharepoint', os.getenv('OpenOrchestratorSQL'), os.getenv('OpenOrchestratorKey'), None, None)

queue_element = MagicMock()
queue_element.data = json.dumps({})

process(orchestrator_connection= orchestrator_connection, queue_element= queue_element)