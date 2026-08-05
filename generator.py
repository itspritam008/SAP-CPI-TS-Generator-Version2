"""
Visteon SAP CPI Technical Specification Generator
==================================================

Parses SAP Integration Suite iFlow ZIP exports and generates a formatted
Technical Specification DOCX.

Key implementation rules:
- BPMN sequenceFlow is the source of truth for execution order.
- BPMN messageFlow direction is the source of truth for sender/receiver roles.
- Main process, error event subprocess, and local exception process are parsed
  separately.
- Local and global/shared Groovy scripts are extracted and deduplicated.
- Optional LLM enrichment can improve wording, but cannot overwrite the
  deterministic BPMN graph, Groovy inventory, success-notification branch,
  archive branch, or exception flow.
"""

import glob
import json
import os
import re
import textwrap
import zipfile
import xml.etree.ElementTree as ET
from collections import defaultdict, deque

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as patches

try:
    from google import genai
except Exception:
    genai = None

from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Inches, Pt, RGBColor


# ---------------------------------------------------------------------------
# Paths and optional LLM client
# ---------------------------------------------------------------------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PRIMARY_LOGO_PATH = os.path.join(BASE_DIR, "Visteon_Logo.jpeg")
ZIP_DIRECTORY_PATH = os.environ.get("ZIP_DIRECTORY_PATH", BASE_DIR)

try:
    if genai and os.environ.get("GEMINI_API_KEY"):
        llm_client = genai.Client(api_key=os.environ.get("GEMINI_API_KEY"))
    else:
        llm_client = None
except Exception as exc:
    llm_client = None
    print(f"Warning: optional LLM client was not initialized: {exc}")


# ---------------------------------------------------------------------------
# Theme
# ---------------------------------------------------------------------------
NAVY = "1F3864"
NAVY_SOFT = "2E5395"
ORANGE = "F2661D"
SLATE = "44546A"
GREEN = "2E7D32"
RED = "C0392B"
PURPLE = "7A4FC0"
AMBER = "B08900"
TEAL = "0E7C86"
BG_LIGHT = "F4F7FB"
BG_ERR = "FDF1EF"
GRID = "D7DEE8"
WHITE = "FFFFFF"
TEXT_BODY = "333333"
FONT = "Calibri"
CONTENT_WIDTH_IN = 6.9

TYPE_COLORS = {
    "Groovy Script": ORANGE,
    "Data Mapping": PURPLE,
    "Lookup / VM": TEAL,
    "Converter / Store": AMBER,
    "Data Store": AMBER,
    "PGP Encryption": PURPLE,
    "Decision Gateway": AMBER,
    "External Call": TEAL,
    "Receiver Call": GREEN,
    "Content Modifier": NAVY_SOFT,
    "Local Process Call": RED,
    "Processing Step": NAVY_SOFT,
    "Exception Handling": RED,
}


def rgb(value):
    return RGBColor.from_string(value)


def sanitize_output_name(name):
    safe = re.sub(r"[^A-Za-z0-9._-]+", "_", name or "iflow").strip("._-")
    return safe or "iflow"


def resolve_asset_path(path):
    if not path:
        return None
    if os.path.isabs(path):
        return path
    for candidate in (path, os.path.join(BASE_DIR, path), os.path.join(os.getcwd(), path)):
        if os.path.exists(candidate):
            return candidate
    return os.path.join(BASE_DIR, path)


def _normalize_text(value):
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value).strip())


def _looks_like_placeholder(value):
    normalized = re.sub(r"[^a-z0-9]+", "", _normalize_text(value).lower())
    return not normalized or normalized in {
        "test", "sample", "dummy", "temp", "tbd", "unknown", "na",
        "sender", "receiver", "sendersystem", "receiversystem",
    }


def _local_name(element_or_tag):
    tag = element_or_tag.tag if hasattr(element_or_tag, "tag") else str(element_or_tag)
    return tag.split("}")[-1]


def _direct_child_text(element, child_name):
    for child in list(element):
        if _local_name(child) == child_name:
            return (child.text or "").strip()
    return ""


def _element_properties(element):
    properties = {}
    for child in element.iter():
        if _local_name(child) != "property":
            continue
        key = child.attrib.get("key") or _direct_child_text(child, "key")
        value = child.attrib.get("value") or _direct_child_text(child, "value")
        if key:
            properties[key.strip()] = (value or "").strip()
    return properties


def _is_timer_start_element(element):
    tag = _local_name(element)
    if tag == "timerStartEvent":
        return True
    if tag == "startEvent":
        return any(_local_name(child) == "timerEventDefinition" for child in element)
    return False


def normalize_adapter_type(adapter_type):
    text = (adapter_type or "Adapter").strip()
    lowered = text.lower()
    mappings = [
        ("idoc", "IDoc"), ("odata", "OData"), ("https", "HTTPS"),
        ("http", "HTTP"), ("sftp", "SFTP"), ("ftp", "FTP"),
        ("soap", "SOAP"), ("amqp", "AMQP"), ("processdirect", "ProcessDirect"),
        ("mail", "Mail"), ("smtp", "Mail"), ("jdbc", "JDBC"),
        ("rfc", "RFC"), ("as2", "AS2"), ("jms", "JMS"),
        ("timer", "Timer"), ("schedule", "Timer"), ("file", "File"),
    ]
    for token, result in mappings:
        if token in lowered:
            return result
    if any(token in lowered for token in ("rest", "api", "webservice")):
        return "REST"
    return text


def _adapter_from_properties(properties, fallback="Adapter"):
    """Return the user-facing adapter family from CPI channel properties."""
    properties = properties or {}
    component = normalize_adapter_type(properties.get("ComponentType", ""))
    message = normalize_adapter_type(properties.get("MessageProtocol", ""))
    transport = normalize_adapter_type(properties.get("TransportProtocol", ""))
    variant = normalize_adapter_type(properties.get("cmdVariantUri", ""))

    # Adapter family is normally ComponentType. OData uses HCIOData internally,
    # which normalize_adapter_type correctly maps to OData.
    if component and component.lower() not in {"adapter", ""}:
        return component
    if message and message.lower() not in {"adapter", "file", ""}:
        return message
    if transport and transport.lower() not in {"adapter", ""}:
        return transport
    if message and message.lower() not in {"adapter", ""}:
        return message
    if variant and variant.lower() not in {"adapter", ""}:
        return variant
    return normalize_adapter_type(fallback)


# ---------------------------------------------------------------------------
# Security material helpers
# ---------------------------------------------------------------------------
def _security_material_label(key):
    if not key:
        return "Security Material"
    label = key.replace("_", " ").replace("-", " ").title()
    return "Security Material" if label.lower() in {"value", "string", "text"} else label


def _add_security_material(data, key, value):
    if value is None:
        return
    value = str(value).strip()
    if not value:
        return
    if any(value == existing for _, existing in data["security_materials"]):
        return
    data["security_materials"].append((_security_material_label(key), value))


# ---------------------------------------------------------------------------
# Groovy script extraction
# ---------------------------------------------------------------------------
def _normalize_script_path(path):
    path = (path or "").strip().replace("\\", "/")
    return re.sub(r"/+", "/", path).lstrip("./")


def _script_identity(path, bundle_id):
    return ((bundle_id or "").strip().lower(), _normalize_script_path(path).lower())


def _add_script_reference(data, path, scope, bundle_id="", step_name="", physical_path=""):
    path = _normalize_script_path(path or physical_path)
    physical_path = _normalize_script_path(physical_path)
    if not path:
        return

    identity = _script_identity(path, bundle_id)
    existing = None
    for item in data["script_references"]:
        if _script_identity(item.get("path"), item.get("bundle_id")) == identity:
            existing = item
            break

    if existing is None:
        existing = {
            "path": path,
            "scope": scope,
            "bundle_id": (bundle_id or "").strip(),
            "used_by": [],
            "physical_paths": [],
        }
        data["script_references"].append(existing)

    if step_name and step_name not in existing["used_by"]:
        existing["used_by"].append(step_name)
    if physical_path and physical_path not in existing["physical_paths"]:
        existing["physical_paths"].append(physical_path)


def _merge_physical_groovy_files(data, physical_files):
    """Merge physical ZIP files with BPMN references without losing shared scripts."""
    for physical in physical_files:
        normalized = _normalize_script_path(physical)
        basename = os.path.basename(normalized).lower()
        matches = [
            item for item in data["script_references"]
            if os.path.basename(item.get("path", "")).lower() == basename
        ]
        if matches:
            for match in matches:
                if normalized not in match["physical_paths"]:
                    match["physical_paths"].append(normalized)
        else:
            _add_script_reference(
                data,
                path=normalized,
                scope="Local",
                physical_path=normalized,
            )

    data["global_groovy_scripts"] = []
    data["local_groovy_scripts"] = []
    data["groovy_scripts"] = []

    for item in data["script_references"]:
        path = item["path"]
        target = (
            data["global_groovy_scripts"]
            if item.get("scope") == "Global"
            else data["local_groovy_scripts"]
        )
        if path not in target:
            target.append(path)
        if path not in data["groovy_scripts"]:
            data["groovy_scripts"].append(path)


def format_groovy_artifacts(iflow_data):
    sections = []
    references = iflow_data.get("script_references", [])

    global_items = [item for item in references if item.get("scope") == "Global"]
    local_items = [item for item in references if item.get("scope") != "Global"]

    if global_items:
        lines = ["GLOBAL / SHARED SCRIPTS"]
        for index, item in enumerate(global_items, 1):
            suffix = f" [Library: {item['bundle_id']}]" if item.get("bundle_id") else ""
            used = f"\n   Used by: {', '.join(item['used_by'])}" if item.get("used_by") else ""
            lines.append(f"{index}. {item['path']}{suffix}{used}")
        sections.append("\n".join(lines))

    if local_items:
        lines = ["LOCAL IFLOW SCRIPTS"]
        for index, item in enumerate(local_items, 1):
            used = f"\n   Used by: {', '.join(item['used_by'])}" if item.get("used_by") else ""
            lines.append(f"{index}. {item['path']}{used}")
        sections.append("\n".join(lines))

    return "\n\n".join(sections) or "None"


# ---------------------------------------------------------------------------
# BPMN graph parsing
# ---------------------------------------------------------------------------
def _classify_bpmn_step(tag_name, name, properties):
    lowered = (name or "").lower()
    activity = properties.get("activityType", "").lower()
    subactivity = properties.get("subActivityType", "").lower()

    if tag_name == "startEvent":
        return ("Error Start Event", "ErrorStart") if "error" in lowered else ("Start Event", "Start")
    if tag_name == "endEvent":
        if "error" in lowered or activity == "enderrorevent":
            return "Error End Event", "ErrorEnd"
        return "End Event", "End"
    if tag_name == "exclusiveGateway":
        return "Decision Gateway", "Gateway"
    if activity == "script" or subactivity == "groovyscript" or properties.get("script"):
        return "Groovy Script", "Processing"
    if activity == "pgpencrypt":
        return "PGP Encryption", "Processing"
    if activity == "pgpdecrypt":
        return "PGP Decryption", "Processing"
    if activity == "dbstorage":
        return "Data Store", "Processing"
    if activity == "send":
        return "Receiver Call", "Processing"
    if activity == "externalcall":
        return "External Call", "Processing"
    if activity == "enricher":
        return "Content Modifier", "Processing"
    if activity == "processcallelement":
        return "Local Process Call", "Processing"
    if "mapping" in lowered or "xslt" in lowered:
        return "Data Mapping", "Processing"
    return "Processing Step", "Processing"


def _parse_process_nodes(container):
    supported = {
        "startEvent", "endEvent", "callActivity", "serviceTask",
        "scriptTask", "sendTask", "task", "exclusiveGateway",
    }
    nodes = {}
    for element in list(container):
        tag = _local_name(element)
        if tag not in supported:
            continue
        node_id = element.attrib.get("id", "")
        properties = _element_properties(element)
        subtype, node_type = _classify_bpmn_step(
            tag, element.attrib.get("name", ""), properties
        )
        nodes[node_id] = {
            "id": node_id,
            "name": element.attrib.get("name", "").strip() or subtype,
            "subtype": subtype,
            "node_type": node_type,
            "tag": tag,
            "properties": properties,
        }
    return nodes


def _parse_sequence_flows(container):
    edges = []
    for element in list(container):
        if _local_name(element) != "sequenceFlow":
            continue
        condition = ""
        for child in list(element):
            if _local_name(child) == "conditionExpression":
                condition = _normalize_text(child.text)
        edges.append({
            "id": element.attrib.get("id", ""),
            "name": element.attrib.get("name", "").strip(),
            "source": element.attrib.get("sourceRef", ""),
            "target": element.attrib.get("targetRef", ""),
            "condition": condition,
        })
    return edges


def _parse_di_positions(tree):
    positions = {}
    for shape in tree.iter():
        if _local_name(shape) != "BPMNShape":
            continue
        element_id = shape.attrib.get("bpmnElement", "")
        for child in list(shape):
            if _local_name(child) == "Bounds":
                try:
                    positions[element_id] = {
                        "x": float(child.attrib.get("x", 0)),
                        "y": float(child.attrib.get("y", 0)),
                        "width": float(child.attrib.get("width", 100)),
                        "height": float(child.attrib.get("height", 60)),
                    }
                except ValueError:
                    pass
                break
    return positions


def _parse_message_flows(tree, participants):
    flows = []
    for element in tree.iter():
        if _local_name(element) != "messageFlow":
            continue
        properties = _element_properties(element)
        source_ref = element.attrib.get("sourceRef", "")
        target_ref = element.attrib.get("targetRef", "")
        direction = properties.get("direction", "").strip()
        adapter = _adapter_from_properties(
            properties,
            element.attrib.get("name") or "Adapter",
        )
        if direction.lower() == "sender":
            system = properties.get("system") or participants.get(source_ref) or element.attrib.get("name", "")
        else:
            system = properties.get("system") or participants.get(target_ref) or element.attrib.get("name", "")
        flows.append({
            "id": element.attrib.get("id", ""),
            "name": element.attrib.get("name", ""),
            "source_ref": source_ref,
            "target_ref": target_ref,
            "direction": direction,
            "adapter_type": adapter,
            "system": system,
            "properties": properties,
        })
    return flows


def _resolve_interface_adapters(data):
    """Resolve source and target adapters from actual main-process message calls.

    CPI external-call adapters are often stored with direction=Receiver even
    when they retrieve business data from the source system. Therefore adapter
    role is determined by BPMN position and business purpose, not direction alone.
    """
    nodes = data.get("main_flow_nodes", {})
    flows = data.get("message_flows", [])
    if not nodes or not flows:
        return

    ordered_ids = _ordered_process_node_ids(data)
    order = {node_id: index for index, node_id in enumerate(ordered_ids)}
    auxiliary_terms = (
        "exception", "error", "event mesh", "event-mesh", "amqp", "mail",
        "email", "archive", "logging", "monitor",
    )

    source_candidates = []
    target_candidates = []
    for flow in flows:
        source_ref = flow.get("source_ref", "")
        target_ref = flow.get("target_ref", "")
        props = flow.get("properties", {}) or {}
        searchable = " ".join([
            flow.get("name", ""), flow.get("system", ""),
            props.get("Name", ""), props.get("MessageProtocol", ""),
            props.get("ComponentType", ""), props.get("TransportProtocol", ""),
            props.get("operation", ""),
        ]).lower()
        if any(term in searchable for term in auxiliary_terms):
            continue

        node = nodes.get(source_ref) or nodes.get(target_ref) or {}
        node_text = " ".join([
            node.get("name", ""), node.get("subtype", ""),
            (node.get("properties", {}) or {}).get("activityType", ""),
        ]).lower()
        adapter = _adapter_from_properties(props, flow.get("adapter_type"))
        system = props.get("system") or flow.get("system") or flow.get("name")
        candidate = (order.get(source_ref, order.get(target_ref, 9999)), flow, adapter, system)

        # Query/read/get calls retrieve data from a source even though CPI marks
        # the adapter channel itself as Receiver.
        is_retrieval = any(term in searchable or term in node_text for term in (
            "odata", "query(get)", "read data", "read ", "get data", "fetch",
            "retrieve", "select",
        ))
        is_delivery = any(term in searchable or term in node_text for term in (
            "send", "write", "deliver", "receiver", "ftp", "sftp", "as2",
        )) and not is_retrieval
        if is_retrieval:
            source_candidates.append(candidate)
        elif is_delivery:
            target_candidates.append(candidate)

    if source_candidates:
        _, flow, adapter, system = sorted(source_candidates, key=lambda item: item[0])[0]
        data["sender_system"] = system or data.get("sender_system", "Source")
        data["sender_adapter_type"] = adapter or data.get("sender_adapter_type", "Adapter")
        data["sender_params"] = list((flow.get("properties", {}) or {}).items())
    if target_candidates:
        _, flow, adapter, system = sorted(target_candidates, key=lambda item: item[0])[-1]
        data["receiver_system"] = system or data.get("receiver_system", "Target")
        data["receiver_adapter_type"] = adapter or data.get("receiver_adapter_type", "Adapter")
        data["receiver_params"] = list((flow.get("properties", {}) or {}).items())


def _is_exception_process(process):
    name = process.attrib.get("name", "").lower()
    process_id = process.attrib.get("id", "").lower()
    return any(token in name or token in process_id for token in ("exception", "error", "fault"))


def _attach_positions(nodes, positions):
    for node_id, node in nodes.items():
        if node_id in positions:
            node["position"] = positions[node_id]


def _topological_layers(nodes, edges):
    incoming = defaultdict(int)
    outgoing = defaultdict(list)
    for edge in edges:
        if edge["source"] in nodes and edge["target"] in nodes:
            outgoing[edge["source"]].append(edge["target"])
            incoming[edge["target"]] += 1
    starts = [node_id for node_id in nodes if incoming[node_id] == 0]
    queue = deque((node_id, 0) for node_id in starts)
    layer = {}
    while queue:
        node_id, depth = queue.popleft()
        if depth <= layer.get(node_id, -1):
            continue
        layer[node_id] = depth
        for target in outgoing[node_id]:
            queue.append((target, depth + 1))
    for node_id in nodes:
        layer.setdefault(node_id, 0)
    return layer


# ---------------------------------------------------------------------------
# Deterministic descriptive content
# ---------------------------------------------------------------------------
def _property_value(properties, key):
    for existing_key, value in properties.items():
        if existing_key.lower() == key.lower():
            return value
    return ""


def _detect_security_operation(data):
    """Detect the actual PGP operation from BPMN activity metadata and names."""
    decrypt_tokens = ("pgpdecrypt", "decrypt", "decryption", "decryptor")
    encrypt_tokens = ("pgpencrypt", "encrypt", "encryption", "encryptor")

    for node in data.get("main_flow_nodes", {}).values():
        properties = node.get("properties", {}) or {}
        searchable = " ".join([
            node.get("name", ""),
            node.get("subtype", ""),
            properties.get("activityType", ""),
            properties.get("subActivityType", ""),
            properties.get("cmdVariantUri", ""),
        ]).lower()
        if any(token in searchable for token in decrypt_tokens):
            return {
                "mode": "decrypt",
                "action": "PGP Decrypt",
                "noun": "PGP Decryption",
                "result": "decrypted",
            }
        if any(token in searchable for token in encrypt_tokens):
            return {
                "mode": "encrypt",
                "action": "PGP Encrypt",
                "noun": "PGP Encryption",
                "result": "encrypted",
            }

    return {
        "mode": "none",
        "action": "Secure File Processing",
        "noun": "Security Processing",
        "result": "processed",
    }


def _extract_timer_events(tree):
    timer_events = []
    for element in tree.iter():
        if _local_name(element) != "startEvent":
            continue
        timer_def = next((child for child in list(element) if _local_name(child) == "timerEventDefinition"), None)
        if timer_def is None:
            continue
        details = {}
        for child in list(timer_def):
            tag = _local_name(child)
            value = _normalize_text(child.text) or _normalize_text(child.attrib.get("expression")) or _normalize_text(child.attrib.get("value"))
            if value:
                details[tag] = value
        timer_events.append({
            "name": _normalize_text(element.attrib.get("name") or "Timer Start Event"),
            "details": details,
        })
    return timer_events


def _guess_auth_type(properties, adapter_type=None):
    properties = properties or {}
    adapter_type = (adapter_type or "").strip()
    auth_type = ""
    for key, value in properties.items():
        key_lower = (key or "").lower()
        value_lower = _normalize_text(value).lower()
        if any(token in key_lower or token in value_lower for token in ("oauth", "bearer", "token")):
            return "OAuth / Token"
        if any(token in key_lower or token in value_lower for token in ("basic", "password", "username", "user")):
            auth_type = "User/password"
        if any(token in key_lower or token in value_lower for token in ("certificate", "cert", "keystore", "truststore", "p12")):
            return "Certificate-based"
        if any(token in key_lower or token in value_lower for token in ("privatekey", "private key", "pgp", "secret")):
            return "Key-based"
        if "ssh" in key_lower or "ssh" in value_lower:
            auth_type = auth_type or "SSH key or password"
    if auth_type:
        return auth_type
    protocol = adapter_type.upper()
    if protocol in {"HTTPS", "HTTP", "SOAP", "REST"}:
        return "SSL / Basic auth"
    if protocol in {"SFTP", "FTP"}:
        return "SSH key or password"
    if protocol == "AS2":
        return "AS2 certificate"
    if protocol == "IDOC":
        return "SAP credentials"
    return "Not explicit"


def _normalize_flow_protocol(flow):
    properties = flow.get("properties", {}) or {}
    return normalize_adapter_type(
        properties.get("MessageProtocol")
        or properties.get("TransportProtocol")
        or properties.get("ComponentType")
        or flow.get("adapter_type")
    )


def _get_flow_block_name(flow, data):
    node = None
    if flow.get("direction", "").lower() == "sender":
        node = data.get("main_flow_nodes", {}).get(flow.get("target_ref", ""))
    else:
        node = data.get("main_flow_nodes", {}).get(flow.get("source_ref", ""))
    if node:
        return node.get("name") or node.get("subtype") or "Integration process"
    return flow.get("name") or flow.get("system") or "Integration process"


def _describe_flow_purpose(flow):
    name = _normalize_text(flow.get("name", "")).lower()
    direction = flow.get("direction", "").lower()
    system = flow.get("system") or "target system"
    if any(token in name for token in ("email", "mail")):
        return "Sends a notification email as part of the flow."
    if any(token in name for token in ("archive", "store", "save")):
        return "Stores a copy of the payload for retention or audit."
    if direction == "sender":
        return "Receives inbound data from the source system and delivers it to the integration process."
    if direction == "receiver":
        return f"Delivers processed data to {system}."
    return "Moves data through the integration flow toward its destination."


def _build_connectivity_rows(data):
    """
    Build a connectivity list suitable for point-wise document rendering.
    Returns a list of dicts: {title, adapter_type, authentication, purpose, details}
    """
    items = []
    timer_events = data.get("timer_events", []) or []
    if not timer_events and "scheduled" in (data.get("execution_mode", "") or "").lower():
        timer_events = [{
            "name": "Scheduled trigger",
            "details": {"Frequency": _normalize_text(data.get("frequency", "Configured schedule"))}
        }]

    for timer in timer_events:
        details = timer.get("details") or {}
        details_text = "; ".join(f"{key}: {value}" for key, value in details.items() if value) or "Configured timer trigger"
        items.append({
            "title": f"{timer['name']} -> Integration flow",
            "adapter_type": "Timer",
            "authentication": "N/A",
            "purpose": f"Starts the flow on schedule: {details_text}.",
            "details": details,
        })

    flows = data.get("message_flows", []) or []
    # Prepare a simple context for optional LLM enrichment
    simple_flows = []
    for flow in flows:
        title = f"{flow.get('direction', '').title() or 'Adapter'} adapter to {_get_flow_block_name(flow, data)}"
        adapter_type = normalize_adapter_type(flow.get("adapter_type") or _normalize_flow_protocol(flow))
        auth = _guess_auth_type(flow.get("properties", {}), adapter_type)
        purpose = _describe_flow_purpose(flow)
        items.append({
            "title": title,
            "adapter_type": adapter_type,
            "authentication": auth,
            "purpose": purpose,
            "details": flow.get("properties", {}) or {},
        })
        simple_flows.append({"title": title, "adapter_type": adapter_type})

    # Optionally ask the LLM to rewrite concise purposes or enrich adapter descriptions
    if llm_client and items:
        try:
            model_name = os.environ.get("GEMINI_MODEL", "gemini-3.5-flash-lite")
            prompt = {
                "items": [{"title": it["title"], "adapter_type": it["adapter_type"]} for it in items]
            }
            llm_prompt = (
                "You are an SAP Integration Suite technical writer. Given the list of adapter entries, "
                "return a JSON array of objects with fields: title, adapter_type, authentication, purpose. "
                "Each purpose must be a single concise sentence grounded only in the provided title and adapter type.\n"
                + json.dumps(prompt, ensure_ascii=False)
            )
            response = llm_client.models.generate_content(model=model_name, contents=llm_prompt)
            raw = getattr(response, "text", "") or ""
            match = re.search(r"\[\s*\{.*\}\s*\]", raw, re.DOTALL)
            if match:
                entries = json.loads(match.group(0))
                if isinstance(entries, list):
                    # merge back enriched purposes where available
                    for enriched in entries:
                        for it in items:
                            if it["title"] == enriched.get("title"):
                                if enriched.get("adapter_type"):
                                    it["adapter_type"] = enriched.get("adapter_type")
                                if enriched.get("authentication"):
                                    it["authentication"] = enriched.get("authentication")
                                if enriched.get("purpose"):
                                    it["purpose"] = _normalize_text(enriched.get("purpose"))
        except Exception:
            pass

    if not items:
        items.append({
            "title": "None detected",
            "adapter_type": "-",
            "authentication": "-",
            "purpose": "No connectivity or authentication details were detected in the parsed iFlow.",
            "details": {},
        })

    return items


def _guess_mapping_purpose_from_name(filename, data):
    name = os.path.basename(filename).lower()
    sender = data.get("sender_system", "source system")
    receiver = data.get("receiver_system", "target system")
    if "idoc" in name:
        return f"Maps source fields into the IDoc structure required by {receiver}."
    if "invoice" in name:
        return f"Maps invoice-related fields into the format expected by {receiver}."
    if "payment" in name:
        return f"Maps payment data from {sender} to the expected receiver format."
    if "order" in name:
        return f"Transforms order data into the structure required by {receiver}."
    if "status" in name:
        return f"Maps status update fields from {sender} to {receiver}."
    if "header" in name or "field" in name or "segment" in name:
        return f"Transforms individual source fields into the receiver mapping format."
    return f"Maps source payload fields into the target format required by {receiver}."


def _summarize_mapping_artifacts(iflow_data):
    artifacts = (iflow_data.get("xslt_scripts", []) or []) + (iflow_data.get("mappings", []) or [])
    rows = []
    if not artifacts:
        return rows
    if llm_client:
        model_name = os.environ.get("GEMINI_MODEL", "gemini-3.5-flash-lite")
        context = {
            "sender_system": iflow_data.get("sender_system", "source system"),
            "receiver_system": iflow_data.get("receiver_system", "target system"),
            "artifacts": artifacts,
        }
        prompt = f"""
You are an SAP CPI technical author. Analyze the mapping artifacts below and return a JSON array of objects with fields 'artifact' and 'purpose'.
Each purpose should be one concise sentence describing the mapping intent in the context of the sender and receiver systems.
Do not invent artifact details beyond the artifact name.
Context: {json.dumps(context, ensure_ascii=False)}
"""
        try:
            response = llm_client.models.generate_content(model=model_name, contents=prompt)
            raw_text = getattr(response, "text", "") or ""
            match = re.search(r"\[\s*\{.*\}\s*\]", raw_text, re.DOTALL)
            if match:
                entries = json.loads(match.group(0))
                if isinstance(entries, list):
                    for entry in entries:
                        if isinstance(entry, dict) and entry.get("artifact"):
                            rows.append([
                                entry["artifact"],
                                "XSLT" if entry["artifact"].lower().endswith((".xsl", ".xslt")) else "Message Mapping",
                                _normalize_text(entry.get("purpose", "")) or _guess_mapping_purpose_from_name(entry["artifact"], iflow_data),
                            ])
                    if rows:
                        return rows
        except Exception:
            pass
    for artifact in artifacts:
        rows.append([
            artifact,
            "XSLT" if artifact.lower().endswith((".xsl", ".xslt")) else "Message Mapping",
            _guess_mapping_purpose_from_name(artifact, iflow_data),
        ])
    return rows


def _detect_features(data):
    names = " ".join(
        node.get("name", "")
        for node in data.get("main_flow_nodes", {}).values()
    ).lower()
    conditions = " ".join(
        edge.get("condition", "")
        for edge in data.get("main_flow_edges", [])
    ).lower()
    properties = data.get("properties", {})
    security = _detect_security_operation(data)
    return {
        "filename": "filename" in names,
        "file_check": "file present" in names,
        "datastore": "data store" in names or "datastore" in names,
        "security_mode": security["mode"],
        "security_action": security["action"],
        "security_noun": security["noun"],
        "security_result": security["result"],
        "success_email": "success email" in names or "sendsuccessfullnotification" in conditions,
        "archive": "archive" in names or "archivefile" in conditions,
        "success_email_value": _property_value(properties, "SendSuccessfullNotification"),
        "archive_value": _property_value(properties, "ArchiveFile"),
    }


def _ordered_node_ids_for_lane(nodes, edges):
    if not nodes:
        return []
    outgoing = {}
    incoming = {node_id: 0 for node_id in nodes}
    for edge in edges:
        source = edge.get("source")
        target = edge.get("target")
        if source in nodes and target in nodes:
            outgoing.setdefault(source, []).append(target)
            incoming[target] = incoming.get(target, 0) + 1
    queue = [node_id for node_id in nodes if incoming.get(node_id, 0) == 0]
    ordered = []
    seen = set()
    while queue:
        node_id = queue.pop(0)
        if node_id in seen:
            continue
        seen.add(node_id)
        ordered.append(node_id)
        queue.extend(outgoing.get(node_id, []))
    ordered.extend(node_id for node_id in nodes if node_id not in seen)
    return ordered


def _ordered_process_node_ids(data):
    """Return process nodes in source order using BPMN DI coordinates when available."""
    nodes = data.get("main_flow_nodes", {})
    edges = data.get("main_flow_edges", [])
    if not nodes:
        return []

    positioned = [
        (node_id, node.get("position", {}).get("x"), node.get("position", {}).get("y"))
        for node_id, node in nodes.items()
        if node.get("position") and node.get("position", {}).get("x") is not None
    ]
    if len(positioned) >= max(2, len(nodes) // 2):
        positioned.sort(key=lambda item: (item[1], item[2] if item[2] is not None else 0))
        ordered = [item[0] for item in positioned]
        ordered.extend(node_id for node_id in nodes if node_id not in ordered)
        return ordered

    layers = _topological_layers(nodes, edges)
    return sorted(nodes, key=lambda node_id: (layers.get(node_id, 0), node_id))


def _compress_actual_steps(items, target_count):
    """Compress adjacent BPMN items without inventing replacement business steps."""
    if len(items) <= target_count:
        return [[item] for item in items]
    groups = []
    total = len(items)
    for group_index in range(target_count):
        start = round(group_index * total / target_count)
        end = round((group_index + 1) * total / target_count)
        if start < end:
            groups.append(items[start:end])
    return groups


def _build_compact_main_stages(data, max_stages=None):
    """Create concise diagram stages exclusively from parsed BPMN nodes and participants."""
    nodes = data.get("main_flow_nodes", {})
    edges = data.get("main_flow_edges", [])
    ordered_ids = _ordered_process_node_ids(data)

    process_items = []
    for node_id in ordered_ids:
        node = nodes[node_id]
        if node.get("node_type") in {"Start", "End", "ErrorStart", "ErrorEnd"}:
            continue
        process_items.append({
            "id": node_id,
            "name": node.get("name") or node.get("subtype") or node_id,
            "subtype": node.get("subtype") or "BPMN Step",
            "node_type": node.get("node_type") or "Processing",
        })

    # No semantic cap: preserve every parsed processing item unless an explicit
    # caller limit is intentionally supplied for a non-document preview.
    if max_stages is None:
        grouped = [[item] for item in process_items]
    else:
        internal_limit = max(1, max_stages - 2)
        grouped = _compress_actual_steps(process_items, internal_limit)

    stages = []
    sender_name = data.get("sender_system") or "Source"
    sender_adapter = data.get("sender_adapter_type") or "Sender Adapter"
    stages.append({
        "name": sender_name,
        "subtype": sender_adapter,
        "node_type": "Sender",
        "source_ids": [],
    })

    for group in grouped:
        names = [item["name"] for item in group]
        subtypes = list(dict.fromkeys(item["subtype"] for item in group))
        node_types = {item["node_type"] for item in group}
        # The body preserves the actual activity names. The banner uses the
        # represented technical categories, never a generic "N BPMN Steps" label.
        if len(subtypes) == 1:
            stage_subtype = subtypes[0]
        else:
            stage_subtype = " / ".join(subtypes[:2])
            if len(subtypes) > 2:
                stage_subtype += " / Processing"
        stages.append({
            "name": " -> ".join(names),
            "subtype": stage_subtype,
            "node_type": "Gateway" if node_types == {"Gateway"} else "Processing",
            "visual_role": "validation" if "Gateway" in node_types else "processing",
            "detail": "",
            "source_ids": [item["id"] for item in group],
        })

    end_names = [
        node.get("name") or "End"
        for node in nodes.values()
        if node.get("node_type") == "End"
    ]
    stages.append({
        "name": end_names[0] if end_names else "End",
        "subtype": "End Event",
        "node_type": "End",
        "source_ids": [],
    })

    # Attach conditions found on actual sequence flows to the stage containing the source gateway.
    stage_by_source = {}
    for stage in stages:
        for source_id in stage.get("source_ids", []):
            stage_by_source[source_id] = stage
    for edge in edges:
        stage = stage_by_source.get(edge.get("source"))
        if not stage:
            continue
        branch = edge.get("name") or edge.get("condition")
        if branch:
            stage.setdefault("branches", [])
            branch = _normalize_text(branch)
            if branch and branch not in stage["branches"]:
                stage["branches"].append(branch)

    return stages


def _describe_parsed_node(node):
    """Explain a parsed node from its actual name and CPI metadata without flow-specific assumptions."""
    name = _normalize_text(node.get("name")) or _normalize_text(node.get("subtype"))
    subtype = _normalize_text(node.get("subtype"))
    properties = node.get("properties", {}) or {}
    activity = _normalize_text(properties.get("activityType"))
    details = []
    if properties.get("script"):
        details.append(f"executes {properties['script']}")
    if properties.get("scriptBundleId"):
        details.append(f"from shared library {properties['scriptBundleId']}")
    if properties.get("processId"):
        details.append(f"calls process {properties['processId']}")
    technical = subtype or activity
    explanation = f"Performs {name}"
    if technical and technical.lower() not in name.lower():
        explanation += f" as a {technical} operation"
    if details:
        explanation += " and " + " ".join(details)
    return explanation + "."


def build_detailed_processing_logic(data):
    """Fallback only: concise explanations derived from actual BPMN nodes and routes."""
    nodes = data.get("main_flow_nodes", {})
    edges = data.get("main_flow_edges", [])
    ordered_ids = _ordered_process_node_ids(data)
    items = [
        nodes[node_id] for node_id in ordered_ids
        if nodes[node_id].get("node_type") not in {"Start", "End", "ErrorStart", "ErrorEnd"}
    ]
    groups = _compress_actual_steps(items, 6)
    logic = []
    edge_by_source = defaultdict(list)
    for edge in edges:
        edge_by_source[edge.get("source")].append(edge)

    for index, group in enumerate(groups, 1):
        names = [_normalize_text(node.get("name")) for node in group if _normalize_text(node.get("name"))]
        title = names[0] if len(names) == 1 else " and ".join(names)
        explanation = " ".join(_describe_parsed_node(node) for node in group)
        routes = []
        for node in group:
            for edge in edge_by_source.get(node.get("id"), []):
                route = _normalize_text(edge.get("name") or edge.get("condition"))
                if route and route not in routes:
                    routes.append(route)
        if routes:
            explanation += " Configured route(s): " + "; ".join(routes) + "."
        logic.append(f"{index}. {title}: {explanation}")

    exception_ids = _ordered_node_ids_for_lane(
        data.get("exception_flow_nodes", {}), data.get("exception_flow_edges", [])
    )
    if exception_ids:
        exception_names = [
            data["exception_flow_nodes"][node_id].get("name", "")
            for node_id in exception_ids
            if data["exception_flow_nodes"][node_id].get("name")
        ]
        logic.append(
            f"{len(logic) + 1}. Exception Handling: The exception process follows "
            + " -> ".join(exception_names)
            + "."
        )
    return logic

def _compact_llm_summary(data):
    return {
        "iflow_name": data.get("iflow_name"),
        "sender_system": data.get("sender_system"),
        "receiver_system": data.get("receiver_system"),
        "sender_adapter": data.get("sender_adapter_type"),
        "receiver_adapter": data.get("receiver_adapter_type"),
        "main_nodes": [node.get("name") for node in data.get("main_flow_nodes", {}).values()],
        "main_edges": data.get("main_flow_edges", []),
        "exception_nodes": [node.get("name") for node in data.get("exception_flow_nodes", {}).values()],
        "scripts": data.get("script_references", []),
        "compact_stages": _build_compact_main_stages(data, max_stages=7),
        "conditional_features": data.get("conditional_features", {}),
    }


def extract_interface_wording_with_llm(data):
    """Two-stage LLM workflow: explain the flow, then design a grounded visual story."""
    if not llm_client:
        print("[LLM] SKIPPED: Gemini client/API key is not configured. Parsed BPMN fallback will be used.")
        return None

    model_name = os.environ.get("GEMINI_MODEL", "gemini-3.5-flash-lite")
    iflow_name = data.get("iflow_name", "Unknown iFlow")
    graph_payload = _llm_graph_payload(data)
    print(f"[LLM] STARTED: Grounded flow analysis for '{iflow_name}' using '{model_name}'.")

    logic_prompt = f"""
You are an expert SAP Integration Suite architect writing a Technical
Specification. Analyze ONLY the structured BPMN graph below.

Return JSON with:
- business_overview: 2 to 3 concise business sentences
- interface_description: 2 to 4 concise technical sentences
- processing_logic: 4 to 7 objects containing title, explanation, and
  source_node_ids

The processing logic must explain what happens and why it matters. Group only
adjacent related activities. Include all critical scripts, mappings, converters,
PGP operations, Data Stores, receiver calls, gateways, notifications, archive
routes, and process calls that exist. Follow sequence_flows. Never invent a step.
Distinguish PGP encryption from PGP decryption. Each explanation must be one or
two useful sentences, not a copy of the BPMN captions. Return JSON only.

Structured BPMN graph:
{json.dumps(graph_payload, ensure_ascii=False)}
"""

    try:
        response = llm_client.models.generate_content(model=model_name, contents=logic_prompt)
        raw_logic = getattr(response, "text", "") or ""
        match = re.search(r"\{.*\}", raw_logic, re.DOTALL)
        if not match:
            print(f"[LLM] FAILED: No logic JSON returned for '{iflow_name}'.")
            return None
        logic_result = json.loads(match.group(0))
        logic_result, reason = _validate_llm_logic(logic_result, data)
        if not logic_result:
            print(f"[LLM] REJECTED: High-level logic for '{iflow_name}': {reason}")
            return None
        print(f"[LLM] LOGIC SUCCESS: Generated {len(logic_result['processing_logic'])} explanatory logic points.")

        diagram_prompt = f"""
You are a Business Process Analyst. Convert the APPROVED HIGH-LEVEL PROCESSING
LOGIC into a clean, professional, layperson-friendly business process flow.

IMPORTANT
- Do not rewrite, shorten, or alter the approved high-level processing logic.
- The diagram must show the core business journey plus the important operational
  outcomes requested below.
- All names, protocols, actions, decisions, and security operations must come
  dynamically from the supplied business context and BPMN graph. Never hardcode
  QAD, Fifth Third Bank, SFTP, payment, encryption, decryption, email, archive,
  or any other flow-specific value unless supported by the current input.

Return JSON only with diagram_stages and connections.

Each stage must contain:
- stage_id
- heading: short phase name such as Source, Receive, Check, Protect, Convert,
  Send, Notify, Save Copy, Stop, or Finish
- label: concise plain-language action, normally 3 to 9 words
- detail: empty string
- node_type: Sender, Processing, Gateway, Receiver, or End
- visual_role: input, validation, processing, transformation, communication,
  governance, completion, or exception
- lane: main, upper, or lower
- source_node_ids: BPMN IDs represented by the stage
- optional: true only when the activity is conditional in the current flow

Each connection must contain from, to, label, and path_type.
For arrows leaving a Gateway, label must be exactly Yes or No. Other arrows
should normally have an empty label.

REQUIRED BUSINESS STORY
1. Identify the business object from the approved logic, such as payment file,
   bank report, order, invoice, employee data, or another supported object.
2. Start with a source stage written in this form where supported:
   "<Business object> from <sender system>".
   Example pattern only: "Payment file from Source ERP".
3. Show how the information arrives. Use the actual sender protocol dynamically,
   for example "Receive via SFTP", "Receive via HTTPS", or "Receive via AS2".
4. Show the meaningful validation question in plain language.
5. Show encryption or decryption when present, translated into everyday wording:
   - encryption -> "Encrypt the file" or "Protect sensitive information"
   - decryption -> "Decrypt the file" or "Open the protected file"
   Do not reverse encryption and decryption.
6. Show conversion or transformation only when it changes information into the
   format required by the receiver. Use plain wording such as
   "Convert to the required format".
7. Determine the actual business outcome from the approved logic and graph.
   Supported outcome patterns include, but are not limited to:
   - receiver delivery: "Send to <receiver system> via <receiver protocol>"
   - notification only: "Notify stakeholders by email"
   - archive only: "Save a copy for records"
   - notification plus archive: show both actions in sequence
   Do not invent a receiver-delivery stage when the flow does not send business
   data to an external receiver.
8. Show an email step when the current flow sends a notification, using wording
   such as "Notify stakeholders by email". Email may be the primary outcome.
9. Show a copy-retention step when the current flow archives or stores a copy,
   using wording such as "Save a copy for records". Archive may be the primary
   or final outcome.
10. Every terminal completion stage must use the exact visible label "END".
    Do not use outcome-specific ending labels such as Notification complete,
    Archive complete, Transfer complete, Process complete, or Success.

LAYOUT RULES
- Main business journey: one dominant left-to-right row.
- No-data or rejected route: upper branch ending in a stop outcome.
- For receiver-delivery flows, notification and saved-copy activities may use a
  compact lower row after delivery.
- For notification-only, archive-only, or notification-plus-archive flows, keep
  those outcome actions on the main row. Do not create an empty receiver stage.
- Keep the diagram simple. Combine configuration and implementation substeps
  that support the same business action.
- Do not add separate boxes for parameters, filenames, headers, properties,
  scripts, local calls, technical logging setup, or adapter configuration.
- Use Yes on the branch that satisfies the decision question and No on the
  branch that does not. Determine this from the target outcome, not arrow order.
- Avoid crossed lines and unnecessary arrow labels.

VISIBLE LANGUAGE RULES
- Allowed when present in the input: actual sender/receiver names, transfer
  protocols such as SFTP/HTTPS/AS2, encryption/decryption, email, and saving a copy.
- Do not show SAP CPI, iFlow, Groovy, Content Modifier, ProcessDirect,
  Data Store, properties, headers, receiver call, gateway, router, BPMN,
  activity IDs, or other implementation terminology.
- Do not show captions such as PROCESS START, FINAL OUTCOME, BUSINESS SOURCE,
  or BUSINESS OUTCOME.

APPROVED HIGH-LEVEL PROCESSING LOGIC:
{json.dumps(logic_result['processing_logic'], ensure_ascii=False)}

DYNAMIC BUSINESS CONTEXT:
{json.dumps({
    'sender_system': graph_payload.get('sender', {}).get('system'),
    'sender_protocol': graph_payload.get('sender', {}).get('adapter'),
    'receiver_system': graph_payload.get('primary_receiver', {}).get('system'),
    'receiver_protocol': graph_payload.get('primary_receiver', {}).get('adapter'),
    'message_flows': graph_payload.get('message_flows', []),
}, ensure_ascii=False)}

TECHNICAL GRAPH FOR FACT CHECKING ONLY:
{json.dumps(graph_payload, ensure_ascii=False)}
"""
        diagram_spec = None
        reason = ""
        previous_diagram_raw = ""
        for diagram_attempt in (1, 2):
            current_prompt = diagram_prompt
            if diagram_attempt == 2:
                current_prompt = f"""
The previous diagram specification was rejected for this precise reason:
{reason}

Repair the diagram JSON. Do not shorten or simplify the approved logic. Include
all missing decisions, outcomes, terminal End stages, and connections. Keep every
stage grounded in the original BPMN node IDs.

Previous diagram response:
{previous_diagram_raw}

Original diagram task:
{diagram_prompt}
"""
                print(f"[LLM] DIAGRAM RETRY: Repairing diagram for '{iflow_name}'.")

            diagram_response = llm_client.models.generate_content(
                model=model_name,
                contents=current_prompt,
            )
            previous_diagram_raw = getattr(diagram_response, "text", "") or ""
            match = re.search(r"\{.*\}", previous_diagram_raw, re.DOTALL)
            if not match:
                reason = "no diagram JSON was returned"
                continue
            diagram_result = json.loads(match.group(0))
            diagram_spec, reason = _validate_diagram_spec(
                diagram_result,
                data,
                processing_logic=logic_result["processing_logic"],
            )
            if diagram_spec:
                visible_text = " ".join(
                    f"{stage.get('heading', '')} {stage.get('name', '')}"
                    for stage in diagram_spec["diagram_stages"]
                ).lower()
                # Sender context is expected when available. Receiver context
                # is not mandatory because some valid flows end with email,
                # archiving, persistence, or another internal business outcome.
                required_context = [
                    data.get("sender_system", ""),
                    data.get("sender_adapter_type", ""),
                ]
                missing_context = [
                    value for value in required_context
                    if _normalize_text(value) and _normalize_text(value).lower() not in visible_text
                ]
                if missing_context:
                    diagram_spec = None
                    reason = "diagram omitted dynamic source context: " + ", ".join(missing_context)
                else:
                    break
            print(
                f"[LLM] DIAGRAM REJECTED: Attempt {diagram_attempt} "
                f"for '{iflow_name}': {reason}"
            )

        if not diagram_spec:
            print(f"[LLM] DIAGRAM FALLBACK for '{iflow_name}': {reason}")
            logic_result["diagram_stages"] = []
            return logic_result

        logic_result.update(diagram_spec)
        print(
            f"[LLM] DIAGRAM SUCCESS: Generated "
            f"{len(diagram_spec['diagram_stages'])} explanatory visual stages "
            f"for '{iflow_name}'."
        )
        print(f"[LLM] SUCCESS: LLM worked properly for '{iflow_name}'.")
        return logic_result

    except json.JSONDecodeError as exc:
        print(f"[LLM] FAILED: Invalid JSON for '{iflow_name}': {exc}")
    except Exception as exc:
        print(f"[LLM] FAILED: {type(exc).__name__}: {exc}")
    print(f"[LLM] FALLBACK: Continuing with BPMN-derived content for '{iflow_name}'.")
    return None


# ---------------------------------------------------------------------------
# User metadata
# ---------------------------------------------------------------------------
def apply_user_metadata(iflow_data, user_inputs=None):
    merged = dict(iflow_data or {})
    user_inputs = user_inputs or {}

    def choose(*values):
        return next((_normalize_text(value) for value in values if _normalize_text(value)), "")

    prepared_by = choose(user_inputs.get("prepared_by"), user_inputs.get("preparedBy"), merged.get("prepared_by")) or "Not provided"
    reviewed_by = choose(user_inputs.get("reviewed_by"), user_inputs.get("reviewedBy"), merged.get("reviewed_by")) or "Not provided"
    approved_by = choose(user_inputs.get("approved_by"), user_inputs.get("approvedBy"), merged.get("approved_by")) or "Not provided"
    effective_date = choose(user_inputs.get("effective_date"), user_inputs.get("effectiveDate"), merged.get("effective_date")) or "TBD"
    description = choose(user_inputs.get("description"), merged.get("description")) or "No description provided."

    direction = choose(user_inputs.get("direction"), merged.get("direction"))
    direction = direction.title() if direction.lower() in {"inbound", "outbound"} else (merged.get("direction") or "Inbound")

    sync_async = choose(user_inputs.get("sync_async"), user_inputs.get("syncAsync"), merged.get("synchronous_asynchronous"))
    if sync_async.lower() in {"sync", "synchronous"}:
        sync_async = "Synchronous"
    elif sync_async.lower() in {"async", "asynchronous"}:
        sync_async = "Asynchronous"
    else:
        sync_async = merged.get("synchronous_asynchronous") or "Synchronous"

    sender = choose(user_inputs.get("source_system"), user_inputs.get("sourceSystem"), merged.get("sender_system"))
    receiver = choose(user_inputs.get("target_system"), user_inputs.get("targetSystem"), merged.get("receiver_system"))
    if _looks_like_placeholder(sender):
        sender = "Source System"
    if _looks_like_placeholder(receiver):
        receiver = "Target System"

    merged.update({
        "package_name": choose(user_inputs.get("package_name"), merged.get("package_name")) or "SAP Integration Package",
        "doc_version": choose(user_inputs.get("doc_version"), merged.get("doc_version")) or "1.0",
        "prepared_by": prepared_by,
        "reviewed_by": reviewed_by,
        "approved_by": approved_by,
        "effective_date": effective_date,
        "description": description,
        "direction": direction,
        "synchronous_asynchronous": sync_async,
        "sender_system": sender,
        "receiver_system": receiver,
        "execution_mode": merged.get("execution_mode") or "Real-time / Event-driven",
        "frequency": merged.get("frequency") or "On demand / operational schedule",
    })
    return merged


def _select_main_process(processes, message_flows):
    """Select the actual message-processing process, not a local email/helper process."""
    if not processes:
        return None, []

    process_by_id = {process.attrib.get("id", ""): process for process in processes}
    node_to_process = {}
    for process in processes:
        process_id = process.attrib.get("id", "")
        for element in process.iter():
            element_id = element.attrib.get("id", "")
            if element_id:
                node_to_process[element_id] = process_id

    exception_processes = [process for process in processes if _is_exception_process(process)]
    non_exception = [process for process in processes if process not in exception_processes]

    # The strongest signal is the process containing the target of the explicit Sender messageFlow.
    for flow in message_flows:
        if (flow.get("direction") or "").lower() != "sender":
            continue
        process_id = node_to_process.get(flow.get("target_ref", ""))
        process = process_by_id.get(process_id)
        if process is not None and process in non_exception:
            return process, exception_processes

    # Next prefer the process containing the source of a non-auxiliary receiver messageFlow.
    auxiliary_terms = ("email", "mail", "archive", "event", "amqp", "exception", "monitor")
    for flow in message_flows:
        if (flow.get("direction") or "").lower() != "receiver":
            continue
        searchable = " ".join([flow.get("name", ""), flow.get("system", "")]).lower()
        if any(term in searchable for term in auxiliary_terms):
            continue
        process_id = node_to_process.get(flow.get("source_ref", ""))
        process = process_by_id.get(process_id)
        if process is not None and process in non_exception:
            return process, exception_processes

    # Final fallback: choose the largest non-exception executable process.
    def score(process):
        executable = sum(
            1 for element in list(process)
            if _local_name(element) in {
                "startEvent", "endEvent", "callActivity", "serviceTask",
                "scriptTask", "sendTask", "task", "exclusiveGateway", "sequenceFlow"
            }
        )
        properties = _element_properties(process)
        if properties.get("processType", "").lower() == "directcall":
            executable -= 20
        return executable

    return (max(non_exception, key=score) if non_exception else processes[0]), exception_processes


def _llm_graph_payload(data):
    """Create a complete, bounded graph payload for grounded LLM summarization."""
    nodes = []
    for node_id in _ordered_process_node_ids(data):
        node = data.get("main_flow_nodes", {}).get(node_id, {})
        properties = node.get("properties", {}) or {}
        nodes.append({
            "id": node_id,
            "name": node.get("name", ""),
            "subtype": node.get("subtype", ""),
            "node_type": node.get("node_type", ""),
            "activity_type": properties.get("activityType", ""),
            "sub_activity_type": properties.get("subActivityType", ""),
            "script": properties.get("script", ""),
            "script_bundle": properties.get("scriptBundleId", ""),
            "process_id": properties.get("processId", ""),
        })

    edges = [
        {
            "source": edge.get("source", ""),
            "target": edge.get("target", ""),
            "route_name": edge.get("name", ""),
            "condition": edge.get("condition", ""),
        }
        for edge in data.get("main_flow_edges", [])
    ]

    receivers = [
        {
            "name": flow.get("name", ""),
            "system": flow.get("system", ""),
            "adapter": flow.get("adapter_type", ""),
            "message_protocol": (flow.get("properties", {}) or {}).get("MessageProtocol", ""),
            "transport_protocol": (flow.get("properties", {}) or {}).get("TransportProtocol", ""),
            "component_type": (flow.get("properties", {}) or {}).get("ComponentType", ""),
            "operation": (flow.get("properties", {}) or {}).get("operation", ""),
            "source_ref": flow.get("source_ref", ""),
            "target_ref": flow.get("target_ref", ""),
            "direction": flow.get("direction", ""),
        }
        for flow in data.get("message_flows", [])
    ]

    exception_nodes = [
        node.get("name", "")
        for node in data.get("error_subprocess_nodes", {}).values()
    ] + [
        node.get("name", "")
        for node in data.get("exception_flow_nodes", {}).values()
    ]

    return {
        "iflow_name": data.get("iflow_name", ""),
        "sender": {
            "system": data.get("sender_system", ""),
            "adapter": data.get("sender_adapter_type", ""),
        },
        "primary_receiver": {
            "system": data.get("receiver_system", ""),
            "adapter": data.get("receiver_adapter_type", ""),
        },
        "nodes": nodes,
        "sequence_flows": edges,
        "message_flows": receivers,
        "exception_steps": exception_nodes,
        "groovy_scripts": data.get("script_references", []),
        "externalized_properties": data.get("properties", {}),
    }


def _validate_llm_logic(result, data):
    """Validate explanatory logic against actual BPMN node IDs."""
    if not isinstance(result, dict):
        return None, "response is not a JSON object"
    logic = result.get("processing_logic")
    if not isinstance(logic, list) or not 4 <= len(logic) <= 7:
        return None, "processing_logic must contain 4 to 7 points"

    valid_ids = set(data.get("main_flow_nodes", {}))
    cleaned_logic = []
    covered_ids = set()
    for item in logic:
        if not isinstance(item, dict):
            return None, "each logic item must be an object"
        title = _normalize_text(item.get("title"))
        explanation = _normalize_text(item.get("explanation"))
        source_ids = [str(value) for value in item.get("source_node_ids", []) if str(value) in valid_ids]
        if not title or len(explanation) < 45 or not source_ids:
            return None, "logic point lacks a useful explanation or grounded node IDs"
        if any(term in f"{title} {explanation}".lower() for term in ("bpmn step", "processing step 1", "processing step 2")):
            return None, "logic contains generic step language"
        covered_ids.update(source_ids)
        cleaned_logic.append({"title": title, "explanation": explanation, "source_node_ids": source_ids})

    critical_ids = {
        node_id for node_id, node in data.get("main_flow_nodes", {}).items()
        if node.get("node_type") == "Gateway"
        or node.get("subtype") in {
            "Groovy Script", "Data Mapping", "PGP Encryption", "PGP Decryption",
            "Data Store", "Receiver Call", "External Call", "Local Process Call",
            "Converter / Store"
        }
    }
    if critical_ids and not critical_ids.issubset(covered_ids):
        return None, "logic omitted critical BPMN nodes: " + ", ".join(sorted(critical_ids - covered_ids))

    cleaned = {
        "business_overview": _normalize_text(result.get("business_overview")),
        "interface_description": _normalize_text(result.get("interface_description")),
        "processing_logic": cleaned_logic,
    }
    if not cleaned["business_overview"] or not cleaned["interface_description"]:
        return None, "overview or interface description is missing"
    return cleaned, ""


def _yes_no_label(value):
    """Convert a conditional route label into a simple Yes or No label."""
    text = _normalize_text(value).lower()
    negative_tokens = (
        "not exist", "does not exist", "no record", "no file", "zero length",
        "false", "disabled", "disable", "skip", "not enabled", "missing",
        "empty", "not present", "absent", "no payload", "archive false",
    )
    positive_tokens = (
        "exist", "record available", "file present", "true", "enabled",
        "enable", "available", "present", "proceed", "continue",
        "archive true", "logging enabled", "notification enabled",
    )
    if any(token in text for token in negative_tokens):
        return "No"
    if any(token in text for token in positive_tokens):
        return "Yes"
    return ""


def _normalize_gateway_connection_labels(stages, connections):
    """Force gateway labels to Yes/No while preserving the branch meaning."""
    stage_by_id = {stage["stage_id"]: stage for stage in stages}
    outgoing = defaultdict(list)
    for connection in connections:
        outgoing[connection["from"]].append(connection)

    for stage_id, stage in stage_by_id.items():
        if stage.get("node_type") != "Gateway":
            continue
        routes = outgoing.get(stage_id, [])
        if not routes:
            continue

        gateway_text = " ".join([
            _normalize_text(stage.get("heading")),
            _normalize_text(stage.get("name")),
            _normalize_text(stage.get("subtype")),
        ]).lower()

        normalized = [_yes_no_label(route.get("label")) for route in routes]

        if len(routes) == 2:
            # Determine which target represents the affirmative action. This is
            # authoritative for optional archive/notification/logging gateways.
            affirmative_index = None
            negative_index = None
            for index, route in enumerate(routes):
                target = stage_by_id.get(route.get("to"), {})
                target_text = " ".join([
                    _normalize_text(target.get("heading")),
                    _normalize_text(target.get("name")),
                    _normalize_text(target.get("subtype")),
                ]).lower()
                target_is_end = target.get("node_type") == "End"
                target_is_optional_action = bool(target.get("optional"))

                if "archive" in gateway_text:
                    if "archive" in target_text or "fetch source" in target_text or target_is_optional_action:
                        affirmative_index = index
                    elif target_is_end or "complete" in target_text:
                        negative_index = index
                elif any(token in gateway_text for token in ("notification", "email", "mail")):
                    if any(token in target_text for token in ("notification", "email", "mail")) or target_is_optional_action:
                        affirmative_index = index
                    elif target_is_end or "archive" in target_text or "complete" in target_text:
                        negative_index = index
                elif any(token in gateway_text for token in ("logging", "data store", "payload log")):
                    if any(token in target_text for token in ("data store", "save", "log")) or target_is_optional_action:
                        affirmative_index = index
                    elif "encrypt" in target_text or "send" in target_text or "deliver" in target_text:
                        negative_index = index
                elif any(token in gateway_text for token in ("record", "file", "payload", "present", "exist")):
                    if target_is_end or any(token in target_text for token in ("stop", "terminate", "no file", "no record")):
                        negative_index = index
                    else:
                        affirmative_index = index

            if affirmative_index is None and negative_index is not None:
                affirmative_index = 1 - negative_index
            if negative_index is None and affirmative_index is not None:
                negative_index = 1 - affirmative_index

            if affirmative_index is not None:
                normalized = ["No", "No"]
                normalized[affirmative_index] = "Yes"
            elif normalized.count("Yes") == 1 and normalized.count("No") == 1:
                pass
            else:
                # Last-resort structural rule: an optional action is Yes; a
                # direct End/skip branch is No. Do not infer from path_type alone.
                optional_target = next(
                    (index for index, route in enumerate(routes)
                     if stage_by_id.get(route.get("to"), {}).get("optional")),
                    None,
                )
                end_target = next(
                    (index for index, route in enumerate(routes)
                     if stage_by_id.get(route.get("to"), {}).get("node_type") == "End"),
                    None,
                )
                if optional_target is not None:
                    normalized = ["No", "No"]
                    normalized[optional_target] = "Yes"
                elif end_target is not None:
                    normalized = ["Yes", "Yes"]
                    normalized[end_target] = "No"
                else:
                    normalized = ["Yes", "No"]
        else:
            normalized = [label or ("Yes" if index == 0 else "No")
                          for index, label in enumerate(normalized)]

        for route, label in zip(routes, normalized):
            route["label"] = label

    return connections


def _validate_diagram_spec(result, data, processing_logic=None):
    """Validate a concise diagram without forcing low-level BPMN expansion."""
    if not isinstance(result, dict):
        return None, "diagram response is not a JSON object"
    stages = result.get("diagram_stages")
    connections = result.get("connections")
    if not isinstance(stages, list) or len(stages) < 2:
        return None, "diagram must contain at least two stages"
    if not isinstance(connections, list) or not connections:
        return None, "diagram connections are required"

    valid_bpmn_ids = set(data.get("main_flow_nodes", {}))
    allowed_types = {"Sender", "Processing", "Gateway", "Receiver", "End"}
    allowed_roles = {
        "input", "validation", "processing", "transformation",
        "communication", "governance", "completion", "exception",
    }
    allowed_lanes = {"main", "upper", "lower"}
    allowed_paths = {"main", "alternate", "optional", "exception"}

    cleaned = []
    stage_ids = set()
    covered_ids = set()
    for index, item in enumerate(stages):
        if not isinstance(item, dict):
            return None, "each diagram stage must be an object"
        stage_id = _normalize_text(item.get("stage_id")) or f"stage_{index + 1}"
        if stage_id in stage_ids:
            return None, f"duplicate stage_id: {stage_id}"
        stage_ids.add(stage_id)
        heading = _normalize_text(item.get("heading"))
        label = _normalize_text(item.get("label"))
        node_type = item.get("node_type") if item.get("node_type") in allowed_types else "Processing"
        role = str(item.get("visual_role") or "processing").lower()
        lane = str(item.get("lane") or "main").lower()
        if role not in allowed_roles:
            role = "processing"
        if lane not in allowed_lanes:
            lane = "main"
        source_ids = [str(v) for v in item.get("source_node_ids", []) if str(v) in valid_bpmn_ids]
        generic = f"{heading} {label}".lower()
        if not heading or not label:
            return None, f"stage {stage_id} requires heading and label"
        if "bpmn step" in generic or re.search(r"\b\d+\s+steps?\b", generic):
            return None, f"stage {stage_id} contains a generic heading"
        forbidden_visible_terms = (
            "sap cpi", "iflow", "groovy", "content modifier", "processdirect",
            "local process", "data store", "exchange property", "receiver call",
            "external call", "exclusive gateway", "router", "pgpencryptor",
            "pgpdecryptor", "bpmn", "activity id", "technical id",
        )
        if any(term in generic for term in forbidden_visible_terms):
            return None, f"stage {stage_id} exposes CPI implementation terminology"
        if node_type not in {"Sender", "End"} and not source_ids:
            return None, f"stage {stage_id} is not grounded in BPMN nodes"
        covered_ids.update(source_ids)
        cleaned.append({
            "stage_id": stage_id,
            "heading": heading,
            "name": label,
            "detail": "",
            "subtype": heading,
            "node_type": node_type,
            "visual_role": role,
            "lane": lane,
            "source_ids": source_ids,
            "optional": bool(item.get("optional", False)),
        })

    cleaned_connections = []
    outgoing = defaultdict(list)
    incoming = defaultdict(list)
    for item in connections:
        source = _normalize_text(item.get("from"))
        target = _normalize_text(item.get("to"))
        label = _normalize_text(item.get("label"))
        path_type = str(item.get("path_type") or "main").lower()
        if source not in stage_ids or target not in stage_ids:
            return None, f"connection {source}->{target} references an unknown stage"
        if source == target:
            return None, f"self-connection is not allowed for {source}"
        if path_type not in allowed_paths:
            path_type = "main"
        connection = {"from": source, "to": target, "label": label, "path_type": path_type}
        cleaned_connections.append(connection)
        outgoing[source].append(connection)
        incoming[target].append(connection)

    cleaned_connections = _normalize_gateway_connection_labels(
        cleaned,
        cleaned_connections,
    )
    # Rebuild route indexes because labels were normalized in-place.
    outgoing = defaultdict(list)
    incoming = defaultdict(list)
    for connection in cleaned_connections:
        outgoing[connection["from"]].append(connection)
        incoming[connection["to"]].append(connection)

    stage_by_id = {stage["stage_id"]: stage for stage in cleaned}
    roots = [stage_id for stage_id in stage_ids if not incoming[stage_id]]
    if not roots:
        return None, "diagram has no start stage"

    for stage_id, stage in stage_by_id.items():
        if stage["node_type"] != "End" and not outgoing[stage_id]:
            return None, f"non-end stage has no outgoing flow: {stage_id}"
        if stage["node_type"] == "Gateway":
            routes = outgoing[stage_id]
            if len(routes) < 2:
                return None, f"gateway {stage_id} does not show both outcomes"
            if any(not route["label"] for route in routes):
                return None, f"gateway {stage_id} has an unlabeled branch"

    reachable = set()
    queue = list(roots)
    while queue:
        current = queue.pop(0)
        if current in reachable:
            continue
        reachable.add(current)
        queue.extend(connection["to"] for connection in outgoing[current])
    if reachable != stage_ids:
        return None, "diagram contains unreachable stages"
    if not any(stage["node_type"] == "End" for stage in cleaned):
        return None, "diagram has no End outcome"

    # Every approved high-level logic point must contribute at least one BPMN ID
    # to the diagram, but individual low-level nodes may be grouped concisely.
    for index, logic_item in enumerate(processing_logic or [], 1):
        logic_ids = {
            str(value) for value in logic_item.get("source_node_ids", [])
            if str(value) in valid_bpmn_ids
        }
        if logic_ids and not (logic_ids & covered_ids):
            return None, f"diagram omitted approved business phase {index}"

    return {"diagram_stages": cleaned, "connections": cleaned_connections}, ""


# ---------------------------------------------------------------------------
# Main ZIP parser
# ---------------------------------------------------------------------------
def parse_cpi_iflow_zip(zip_file_path, user_inputs=None):
    base_filename = os.path.splitext(os.path.basename(zip_file_path))[0]
    bpmn_xml_str = None
    data = {
        "iflow_name": base_filename,
        "package_name": "SAP Integration Package",
        "doc_version": "1.0",
        "sender_system": "Sender System",
        "receiver_system": "Receiver System",
        "sender_adapter_type": "Adapter",
        "receiver_adapter_type": "Adapter",
        "execution_mode": "Real-time / Event-driven",
        "synchronous_asynchronous": "Synchronous",
        "frequency": "Immediate upon request",
        "direction": "Inbound",
        "prepared_by": "",
        "reviewed_by": "",
        "approved_by": "",
        "effective_date": "",
        "description": "",
        "business_overview": "",
        "interface_description": "",
        "processing_logic": [],
        "main_pipeline_nodes": [],
        "main_flow_nodes": {},
        "main_flow_edges": [],
        "error_subprocess_nodes": {},
        "error_subprocess_edges": [],
        "exception_flow_nodes": {},
        "exception_flow_edges": [],
        "diagram_positions": {},
        "groovy_scripts": [],
        "local_groovy_scripts": [],
        "global_groovy_scripts": [],
        "script_references": [],
        "xslt_scripts": [],
        "mappings": [],
        "properties": {},
        "sender_params": [],
        "receiver_params": [],
        "security_materials": [],
        "message_flows": [],
        "auxiliary_receivers": [],
        "conditional_features": {},
    }

    try:
        with zipfile.ZipFile(zip_file_path, "r") as archive:
            file_list = archive.namelist()

            metafiles = [name for name in file_list if name.lower().endswith("metainfo.prop")]
            if metafiles:
                content = archive.read(metafiles[0]).decode("utf-8", errors="ignore")
                for line in content.splitlines():
                    if "=" not in line or line.strip().startswith("#"):
                        continue
                    key, value = [part.strip() for part in line.split("=", 1)]
                    if key in {"SymbolicName", "Bundle-Name", "Name"} and value:
                        data["iflow_name"] = value
                    elif key in {"OriginBundle-SymbolicName", "PackageName"} and value:
                        data["package_name"] = value

            physical_groovy = []
            for member in file_list:
                normalized = member.replace("\\", "/")
                lowered = normalized.lower()
                if lowered.endswith(".groovy"):
                    physical_groovy.append(normalized)
                elif lowered.endswith((".xsl", ".xslt")):
                    if normalized not in data["xslt_scripts"]:
                        data["xslt_scripts"].append(normalized)
                elif lowered.endswith((".mmap", ".mmap.xml", ".map")):
                    if normalized not in data["mappings"]:
                        data["mappings"].append(normalized)

            param_files = [name for name in file_list if name.lower().endswith("parameters.prop")]
            if param_files:
                content = archive.read(param_files[0]).decode("utf-8", errors="ignore")
                for line in content.splitlines():
                    if "=" not in line or line.strip().startswith("#"):
                        continue
                    key, value = [part.strip() for part in line.split("=", 1)]
                    data["properties"][key] = value
                    if any(token in key.lower() for token in ("credential", "alias", "user", "key", "auth", "password", "certificate", "material")):
                        _add_security_material(data, key, value)

            bpmn_files = [
                name for name in file_list
                if name.lower().endswith((".iflw", ".ifbw")) or name.lower().endswith("component.xml")
            ]
            if not bpmn_files:
                raise ValueError("No BPMN iFlow artifact (.iflw/.ifbw/component.xml) was found.")

            bpmn_bytes = archive.read(bpmn_files[0])
            bpmn_xml_str = bpmn_bytes.decode("utf-8", errors="ignore")
            tree = ET.fromstring(bpmn_bytes)

            participants = {}
            participant_types = {}
            for element in tree.iter():
                if _local_name(element) != "participant":
                    continue
                participant_id = element.attrib.get("id", "")
                participants[participant_id] = element.attrib.get("name", "").strip()
                participant_types[participant_id] = element.attrib.get("{http:///com.sap.ifl.model/Ifl.xsd}type", "")

            data["diagram_positions"] = _parse_di_positions(tree)
            data["message_flows"] = _parse_message_flows(tree, participants)
            data["timer_events"] = _extract_timer_events(tree)

            sender_flows = []
            receiver_flows = []
            for flow in data["message_flows"]:
                direction = flow.get("direction", "").lower()
                params = list(flow.get("properties", {}).items())
                if direction == "sender":
                    sender_flows.append(flow)
                    data["sender_system"] = flow.get("system") or flow.get("name") or data["sender_system"]
                    data["sender_adapter_type"] = flow.get("adapter_type") or data["sender_adapter_type"]
                    data["sender_params"] = params
                    if flow.get("properties", {}).get("scheduleKey"):
                        data["execution_mode"] = "Scheduled / Batched"
                        data["synchronous_asynchronous"] = "Asynchronous"
                        data["frequency"] = "Configured Schedule / Polling"
                elif direction == "receiver":
                    receiver_flows.append(flow)

                for key, value in flow.get("properties", {}).items():
                    if any(token in key.lower() for token in ("credential", "alias", "user", "key", "auth", "password", "certificate", "material")):
                        _add_security_material(data, key, value)

            auxiliary_terms = {
                "email", "mail", "archive", "event-mesh", "event_mesh",
                "event mesh", "amqp", "exception", "monitor", "logging",
            }
            primary_candidates = []
            for flow in receiver_flows:
                searchable = " ".join([
                    flow.get("name", ""), flow.get("system", ""), flow.get("adapter_type", "")
                ]).lower()
                if any(term in searchable for term in auxiliary_terms):
                    data["auxiliary_receivers"].append(flow)
                else:
                    primary_candidates.append(flow)

            if primary_candidates:
                primary = primary_candidates[0]
                data["receiver_system"] = primary.get("system") or primary.get("name") or data["receiver_system"]
                data["receiver_adapter_type"] = primary.get("adapter_type") or data["receiver_adapter_type"]
                data["receiver_params"] = list(primary.get("properties", {}).items())

            if data["sender_adapter_type"] == "SFTP" and data["receiver_adapter_type"] == "SFTP":
                data["direction"] = "Outbound"

            processes = [element for element in tree.iter() if _local_name(element) == "process"]
            main_process, exception_processes = _select_main_process(
                processes,
                data["message_flows"],
            )
            if main_process is not None:
                print(
                    f"[PARSER] Main process selected: "
                    f"{main_process.attrib.get('name') or main_process.attrib.get('id')}"
                )

            if main_process is not None:
                data["main_flow_nodes"] = _parse_process_nodes(main_process)
                data["main_flow_edges"] = _parse_sequence_flows(main_process)
                _attach_positions(data["main_flow_nodes"], data["diagram_positions"])
                _resolve_interface_adapters(data)
                print(
                    f"[PARSER] Resolved adapters: source={data['sender_adapter_type']} "
                    f"({data['sender_system']}), target={data['receiver_adapter_type']} "
                    f"({data['receiver_system']})"
                )

                for child in list(main_process):
                    if _local_name(child) != "subProcess":
                        continue
                    props = _element_properties(child)
                    name = child.attrib.get("name", "").lower()
                    if "exception" in name or "error" in name or props.get("activityType", "").lower() == "erroreventsubprocesstemplate":
                        data["error_subprocess_nodes"] = _parse_process_nodes(child)
                        data["error_subprocess_edges"] = _parse_sequence_flows(child)
                        _attach_positions(data["error_subprocess_nodes"], data["diagram_positions"])

            for process in exception_processes:
                data["exception_flow_nodes"].update(_parse_process_nodes(process))
                data["exception_flow_edges"].extend(_parse_sequence_flows(process))
            _attach_positions(data["exception_flow_nodes"], data["diagram_positions"])

            for element in tree.iter():
                props = _element_properties(element)
                script_path = props.get("script", "")
                if not script_path:
                    continue
                scope = "Global" if props.get("scriptBundleId") else "Local"
                _add_script_reference(
                    data,
                    path=script_path,
                    scope=scope,
                    bundle_id=props.get("scriptBundleId", ""),
                    step_name=element.attrib.get("name", "").strip(),
                )

            _merge_physical_groovy_files(data, physical_groovy)
            data["conditional_features"] = _detect_features(data)
            security_mode = data["conditional_features"].get("security_mode", "none")
            print(
                f"[PARSER] Security operation detected for '{data['iflow_name']}': "
                f"{security_mode.upper()}"
            )
            data["processing_logic"] = build_detailed_processing_logic(data)

            data["main_pipeline_nodes"] = [
                (node["name"], node["subtype"])
                for node in data["main_flow_nodes"].values()
                if node["node_type"] not in {"Start", "End", "ErrorStart", "ErrorEnd"}
            ]

    except Exception as exc:
        print(f"WARNING: could not fully parse ZIP '{os.path.basename(zip_file_path)}': {exc}")

    for key, value in data["properties"].items():
        lowered = key.lower()
        if "receiver_system" == lowered or lowered == "receiver system":
            data["receiver_system"] = value
        elif lowered in {"sap_sender", "sender_system", "sender system"}:
            data["sender_system"] = value

    if not data["processing_logic"]:
        data["processing_logic"] = [
            f"The flow receives data from {data['sender_system']}, processes it in SAP Integration Suite, and delivers it to {data['receiver_system']}."
        ]

    wording = extract_interface_wording_with_llm(data) if bpmn_xml_str else None
    if wording:
        data["business_overview"] = _normalize_text(wording.get("business_overview"))
        data["interface_description"] = _normalize_text(wording.get("interface_description"))
        data["processing_logic"] = [
            f"{index}. {item['title']}: {item['explanation']}"
            for index, item in enumerate(wording.get("processing_logic", []), 1)
        ]
        data["diagram_stages"] = wording.get("diagram_stages", [])
        data["diagram_connections"] = wording.get("connections", [])
        data["diagram_outcome"] = wording.get("outcome", "")
        data["diagram_subtitle"] = wording.get("subtitle", "")
        data["llm_analysis_status"] = "success"
    else:
        data["llm_analysis_status"] = "fallback"

    if not data["business_overview"]:
        data["business_overview"] = (
            f"This integration automates the transfer of business data from {data['sender_system']} "
            f"to {data['receiver_system']} through SAP Integration Suite."
        )
    if not data["interface_description"]:
        data["interface_description"] = (
            f"The interface receives data through {data['sender_adapter_type']}, executes the parsed BPMN processing and conditional branches, "
            f"and delivers the result through {data['receiver_adapter_type']}."
        )

    return apply_user_metadata(data, user_inputs)


# ---------------------------------------------------------------------------
# Diagram rendering
# ---------------------------------------------------------------------------
def _wrap_label(text, width_chars=18, max_lines=4):
    text = (text or "").strip().replace("_", " ")
    return textwrap.wrap(text, width=width_chars, max_lines=max_lines, break_long_words=False) or [""]


def _node_color(node):
    role_colors = {
        "input": NAVY_SOFT,
        "validation": AMBER,
        "processing": NAVY_SOFT,
        "transformation": GREEN,
        "communication": TEAL,
        "governance": PURPLE,
        "completion": SLATE,
        "exception": RED,
    }
    if node.get("visual_role") in role_colors:
        return role_colors[node["visual_role"]]
    node_type = node.get("node_type")
    if node_type in {"Start", "End"}:
        return SLATE
    if node_type in {"ErrorStart", "ErrorEnd"}:
        return RED
    if node_type == "Gateway":
        return AMBER
    return TYPE_COLORS.get(node.get("subtype"), NAVY_SOFT)


def _draw_card(ax, x, y, width, height, node, number=None, scale=1.0):
    color = _node_color(node)
    if node.get("node_type") == "Gateway":
        diamond = patches.Polygon(
            [(x, y + height / 2), (x + width / 2, y), (x, y - height / 2), (x - width / 2, y)],
            closed=True, facecolor=f"#{WHITE}", edgecolor=f"#{color}", linewidth=1.3, zorder=5,
        )
        ax.add_patch(diamond)
        lines = _wrap_label(node.get("name"), 15, 3)
        ax.text(x, y, "\n".join(lines), ha="center", va="center", fontsize=6.2 * scale,
                fontweight="bold", color=f"#{SLATE}", zorder=6)
        return

    ax.add_patch(patches.FancyBboxPatch(
        (x - width / 2 + 0.035, y - height / 2 - 0.035), width, height,
        boxstyle="round,pad=0.02,rounding_size=0.08", linewidth=0,
        facecolor="#00102A", alpha=0.12, zorder=4,
    ))
    ax.add_patch(patches.FancyBboxPatch(
        (x - width / 2, y - height / 2), width, height,
        boxstyle="round,pad=0.02,rounding_size=0.08", linewidth=1.0,
        edgecolor=f"#{color}", facecolor=f"#{WHITE}", zorder=5,
    ))
    banner_h = 0.28
    ax.add_patch(patches.Rectangle(
        (x - width / 2, y + height / 2 - banner_h), width, banner_h,
        linewidth=0, facecolor=f"#{color}", zorder=6,
    ))
    ax.text(x, y + height / 2 - banner_h / 2, node.get("subtype", "STEP").upper(),
            fontsize=5.4 * scale, fontweight="bold", color=f"#{WHITE}",
            ha="center", va="center", zorder=7)
    lines = _wrap_label(node.get("name"), 18, 4)
    detail = _normalize_text(node.get("detail"))
    label_y = y + 0.02 if detail else y - 0.05
    ax.text(x, label_y, "\n".join(lines), fontsize=6.4 * scale, fontweight="bold",
            color=f"#{SLATE}", ha="center", va="center", zorder=7)
    if detail:
        detail_lines = _wrap_label(detail, 22, 2)
        ax.text(x, y - height * 0.27, "\n".join(detail_lines), fontsize=5.2 * scale,
                color=f"#{SLATE}", ha="center", va="center", zorder=7)
    if number is not None:
        ax.add_patch(patches.Circle((x - width / 2, y + height / 2), 0.13,
                                    facecolor=f"#{NAVY}", edgecolor=f"#{WHITE}", linewidth=1.0, zorder=8))
        ax.text(x - width / 2, y + height / 2, str(number), fontsize=6.0 * scale,
                fontweight="bold", color=f"#{WHITE}", ha="center", va="center", zorder=9)


def _graph_positions(nodes, edges, lane_y, x_start, x_end, branch_height=0.9):
    if not nodes:
        return {}
    positions_with_di = {
        node_id: node.get("position") for node_id, node in nodes.items() if node.get("position")
    }
    if len(positions_with_di) >= max(2, len(nodes) // 2):
        xs = [value["x"] for value in positions_with_di.values()]
        ys = [value["y"] for value in positions_with_di.values()]
        min_x, max_x = min(xs), max(xs)
        median_y = sorted(ys)[len(ys) // 2]
        span = max(max_x - min_x, 1.0)
        result = {}
        for node_id, node in nodes.items():
            pos = node.get("position")
            if pos:
                x = x_start + (pos["x"] - min_x) / span * (x_end - x_start)
                relative = (pos["y"] - median_y) / 110.0
                y = lane_y - max(-1.2, min(1.2, relative)) * branch_height
                result[node_id] = (x, y)
        return result

    layers = _topological_layers(nodes, edges)
    max_layer = max(layers.values()) if layers else 1
    grouped = defaultdict(list)
    for node_id, layer in layers.items():
        grouped[layer].append(node_id)
    result = {}
    for layer, node_ids in grouped.items():
        x = x_start + (layer / max(max_layer, 1)) * (x_end - x_start)
        offsets = [0] if len(node_ids) == 1 else [branch_height * (i - (len(node_ids) - 1) / 2) for i in range(len(node_ids))]
        for node_id, offset in zip(node_ids, offsets):
            result[node_id] = (x, lane_y - offset)
    return result


def _draw_graph(ax, nodes, edges, positions, color=NAVY, numbered=True):
    node_w, node_h = 1.55, 0.88
    for edge in edges:
        if edge["source"] not in positions or edge["target"] not in positions:
            continue
        x1, y1 = positions[edge["source"]]
        x2, y2 = positions[edge["target"]]
        ax.annotate("", xy=(x2 - node_w / 2 + 0.03, y2), xytext=(x1 + node_w / 2 - 0.03, y1),
                    arrowprops=dict(arrowstyle="-|>", color=f"#{color}", lw=1.15,
                                    mutation_scale=10, connectionstyle="arc3,rad=0.0"), zorder=3)
        label = edge.get("name") or ""
        if label:
            ax.text((x1 + x2) / 2, (y1 + y2) / 2 + 0.12, label, fontsize=5.5,
                    color=f"#{color}", ha="center", va="bottom", fontweight="bold", zorder=8)

    counter = 1
    for node_id, node in sorted(nodes.items(), key=lambda item: positions.get(item[0], (0, 0))[0]):
        if node_id not in positions:
            continue
        number = None
        if numbered and node.get("node_type") == "Processing":
            number = counter
            counter += 1
        _draw_card(ax, *positions[node_id], node_w, node_h, node, number=number)


def _summary_node(name, subtype, node_type="Processing"):
    return {"name": name, "subtype": subtype, "node_type": node_type}


def _draw_summary_arrow(ax, start, end, label="", color=NAVY, connectionstyle="arc3,rad=0.0"):
    x1, y1 = start
    x2, y2 = end
    ax.annotate(
        "",
        xy=(x2, y2),
        xytext=(x1, y1),
        arrowprops=dict(
            arrowstyle="-|>",
            color=f"#{color}",
            lw=1.35,
            mutation_scale=11,
            connectionstyle=connectionstyle,
        ),
        zorder=3,
    )
    if label:
        ax.text(
            (x1 + x2) / 2,
            (y1 + y2) / 2 + 0.12,
            label,
            fontsize=6.2,
            color=f"#{color}",
            ha="center",
            va="bottom",
            fontweight="bold",
            zorder=9,
        )


def _fit_story_text(text, width, height, is_gateway=False):
    """Return wrapped lines and a fitted font size that remain inside the shape."""
    value = _normalize_text(text)
    if not value:
        return [""], 6.0

    # Conservative character capacity derived from the actual rendered shape.
    if is_gateway:
        chars_per_line = max(9, int(width * 8.2))
        max_lines = max(2, int(height * 2.45))
        base_size = 7.9
        min_size = 5.7
    else:
        chars_per_line = max(11, int(width * 10.2))
        max_lines = max(2, int(height * 3.1))
        base_size = 9.0
        min_size = 6.2

    words = value.split()
    lines = []
    current = ""
    for word in words:
        candidate = word if not current else current + " " + word
        if len(candidate) <= chars_per_line:
            current = candidate
        else:
            if current:
                lines.append(current)
            # Split unusually long technical tokens rather than letting them overflow.
            while len(word) > chars_per_line:
                lines.append(word[:chars_per_line - 1] + "-")
                word = word[chars_per_line - 1:]
            current = word
    if current:
        lines.append(current)

    longest = max((len(line) for line in lines), default=1)
    line_factor = max(1.0, len(lines) / max_lines)
    width_factor = max(1.0, longest / chars_per_line)
    font_size = max(min_size, base_size / max(line_factor, width_factor))

    # If wrapping still exceeds the preferred line count, shrink further. Keep
    # all wording intact rather than truncating technical content.
    if len(lines) > max_lines:
        font_size = max(min_size, font_size * max_lines / len(lines))

    return lines, font_size


def _estimate_story_box_height(text, width, base_height=1.18):
    """Increase card height for long labels while retaining a compact layout."""
    lines, _ = _fit_story_text(text, width, max(base_height, 1.0), is_gateway=False)
    required = 0.38 + 0.245 * len(lines)
    return max(base_height, min(1.75, required))


def _draw_story_box(ax, x, y, w, h, node):
    """Render a clean enterprise card with polished shape, color, and shadow."""
    color = _node_color(node)
    fills = {
        "input": "EDF5FF",
        "validation": "FFF7DF",
        "processing": "F0F6FF",
        "transformation": "ECF9F0",
        "communication": "EAF9F7",
        "governance": "F5EEFC",
        "completion": "F1F3F5",
        "exception": "FDEEEE",
    }
    fill = fills.get(node.get("visual_role"), "F0F6FF")
    optional = node.get("optional", False)

    # Layered shadow preserves the polished card appearance.
    ax.add_patch(patches.FancyBboxPatch(
        (x - w / 2 + 0.075, y - h / 2 - 0.075), w, h,
        boxstyle="round,pad=0.015,rounding_size=0.13",
        facecolor="#16324F", edgecolor="none", alpha=0.08, zorder=3,
    ))
    ax.add_patch(patches.FancyBboxPatch(
        (x - w / 2 + 0.035, y - h / 2 - 0.035), w, h,
        boxstyle="round,pad=0.015,rounding_size=0.13",
        facecolor="#16324F", edgecolor="none", alpha=0.06, zorder=4,
    ))

    # Main card. Optional paths retain a dashed border but no extra text chip.
    ax.add_patch(patches.FancyBboxPatch(
        (x - w / 2, y - h / 2), w, h,
        boxstyle="round,pad=0.015,rounding_size=0.13",
        facecolor=f"#{fill}", edgecolor=f"#{color}", linewidth=2.15,
        linestyle="--" if optional else "-", zorder=5,
    ))

    # Thin accent line keeps the style without adding another text label.
    ax.add_patch(patches.FancyBboxPatch(
        (x - w / 2 + 0.11, y + h / 2 - 0.13), w - 0.22, 0.075,
        boxstyle="round,pad=0.008,rounding_size=0.035",
        facecolor=f"#{color}", edgecolor="none", zorder=6,
    ))

    # Fit the full process action inside the actual card dimensions.
    lines, font_size = _fit_story_text(
        node.get("name", ""),
        width=w - 0.06,
        height=h - 0.08,
        is_gateway=False,
    )
    ax.text(
        x, y - 0.015, "\n".join(lines), ha="center", va="center",
        fontsize=font_size, color="#13263A", fontweight="bold", zorder=8,
        linespacing=1.05,
    )


def _draw_story_gateway(ax, x, y, size, node):
    """Render a clean decision gateway without additional captions."""
    color = _node_color(node)
    shadow_points = [
        [x + 0.055, y + size - 0.055],
        [x + size + 0.055, y - 0.055],
        [x + 0.055, y - size - 0.055],
        [x - size + 0.055, y - 0.055],
    ]
    ax.add_patch(patches.Polygon(
        shadow_points, closed=True, facecolor="#16324F",
        edgecolor="none", alpha=0.10, zorder=3,
    ))
    points = [[x, y + size], [x + size, y], [x, y - size], [x - size, y]]
    ax.add_patch(patches.Polygon(
        points, closed=True, facecolor="#FFF7DF",
        edgecolor=f"#{color}", linewidth=2.45, zorder=5,
    ))
    inner = size * 0.82
    inner_points = [[x, y + inner], [x + inner, y], [x, y - inner], [x - inner, y]]
    ax.add_patch(patches.Polygon(
        inner_points, closed=True, fill=False,
        edgecolor=f"#{color}", linewidth=0.65, alpha=0.45, zorder=6,
    ))
    lines, font_size = _fit_story_text(
        node.get("name", ""),
        width=size * 1.30,
        height=size * 1.34,
        is_gateway=True,
    )
    ax.text(
        x, y, "\n".join(lines),
        ha="center", va="center", fontsize=font_size,
        color="#13263A", fontweight="bold", zorder=8,
        linespacing=0.96,
    )


def _stage_boundary(position,node,w=2.25,h=1.15,size=0.72):
    x,y=position
    if node.get("node_type")=="Gateway": return x-size,y,x+size,y
    return x-w/2,y,x+w/2,y


def _draw_story_connection(ax, start, end, label="", path_type="main"):
    """Draw a clean orthogonal connector with an unambiguous label position."""
    sx, sy = start
    tx, ty = end
    color = "#4A4A4A"
    if path_type in {"alternate", "exception"}:
        label_color = "#C0392B"
    elif path_type == "optional":
        label_color = "#7A4FC0"
    else:
        label_color = "#118A22"

    if abs(sy - ty) < 0.08:
        points = [(sx, sy), (tx, ty)]
    else:
        # Leave the source vertically, then travel horizontally at a dedicated
        # routing level, then enter the target vertically.
        direction = 1 if ty > sy else -1
        bend_y = sy + direction * min(0.72, max(0.42, abs(ty - sy) * 0.38))
        points = [(sx, sy), (sx, bend_y), (tx, bend_y), (tx, ty)]

    for index in range(len(points) - 1):
        p1, p2 = points[index], points[index + 1]
        arrow = index == len(points) - 2
        ax.annotate(
            "", xy=p2, xytext=p1,
            arrowprops=dict(
                arrowstyle="-|>" if arrow else "-", lw=1.85, color=color,
                shrinkA=0, shrinkB=0, mutation_scale=10,
            ),
            zorder=3,
        )

    if label:
        if len(points) == 2:
            lx = sx + (tx - sx) * 0.30
            ly = sy + 0.17
        else:
            # Label the short source segment so Yes/No clearly belongs to the
            # gateway, not to a distant optional activity.
            lx = sx + 0.20
            ly = sy + (points[1][1] - sy) * 0.45
        ax.text(
            lx, ly, label, fontsize=6.8, color=label_color,
            ha="left", va="center", fontweight="bold", zorder=9,
            bbox=dict(boxstyle="round,pad=0.10", facecolor="white",
                      edgecolor="none", alpha=0.92),
        )


def _snake_positions(items, x_left, x_right, top_y, columns=6, row_gap=1.65, reverse_first=False):
    """Lay out any number of stages in readable alternating rows."""
    positions = {}
    if not items:
        return positions, top_y
    for index, node in enumerate(items):
        row = index // columns
        column = index % columns
        items_in_row = min(columns, len(items) - row * columns)
        if items_in_row == 1:
            xs = [(x_left + x_right) / 2]
        else:
            xs = [
                x_left + i * (x_right - x_left) / (items_in_row - 1)
                for i in range(items_in_row)
            ]
        if (row % 2 == 1) != reverse_first:
            xs = list(reversed(xs))
        positions[node["stage_id"]] = (xs[column], top_y - row * row_gap)
    rows = (len(items) + columns - 1) // columns
    lowest_y = top_y - max(rows - 1, 0) * row_gap
    return positions, lowest_y


def _simplify_business_diagram(stages, connections):
    """Create a clean business graph and normalize all decision semantics."""
    implementation_terms = (
        "parameter", "filename generation", "custom header", "property setup",
        "content modifier", "script", "processdirect", "local process",
        "data store naming", "adapter configuration", "tracking header",
    )
    stages = [dict(stage) for stage in stages]
    connections = [dict(connection) for connection in connections]
    stage_by_id = {stage["stage_id"]: stage for stage in stages}

    # Remove implementation-only boxes.
    removed = {
        stage["stage_id"] for stage in stages
        if any(term in f"{stage.get('heading', '')} {stage.get('name', '')}".lower()
               for term in implementation_terms)
    }

    # Collapse optional yes/no gateways for email and archive. A business reader
    # needs the action, not a separate decision diamond and two unlabeled bypasses.
    collapsible_gateways = []
    for stage in stages:
        if stage.get("node_type") != "Gateway":
            continue
        label = f"{stage.get('heading', '')} {stage.get('name', '')}".lower()
        if any(term in label for term in ("email", "notification", "notify", "archive", "archiv")):
            collapsible_gateways.append(stage["stage_id"])

    for gateway_id in collapsible_gateways:
        incoming = [c for c in connections if c.get("to") == gateway_id]
        outgoing = [c for c in connections if c.get("from") == gateway_id]
        positive = None
        bypass = None
        for route in outgoing:
            target = stage_by_id.get(route.get("to"), {})
            target_text = f"{target.get('heading', '')} {target.get('name', '')}".lower()
            is_operational_action = target.get("node_type") != "End" and (
                target.get("optional") or any(term in target_text for term in (
                    "email", "notify", "notification", "archive", "save a copy",
                    "save copy", "records", "logging", "log", "retain source",
                    "source payload", "audit copy", "audit log"
                ))
            )
            if is_operational_action:
                positive = route
            else:
                bypass = route
        if positive:
            action_id = positive.get("to")
            if action_id in stage_by_id:
                stage_by_id[action_id]["lane"] = "lower"
                stage_by_id[action_id]["optional"] = True
            action_outgoing = [c for c in connections if c.get("from") == action_id]
            bypass_target = bypass.get("to") if bypass else None
            if not bypass_target and action_outgoing:
                bypass_target = action_outgoing[0].get("to")
            connections = [c for c in connections if c.get("from") != gateway_id and c.get("to") != gateway_id]
            # Preserve the main business route by bridging around the removed
            # operational gateway. The optional action is placed in the lower
            # operational sequence later and does not interrupt the happy path.
            if bypass_target:
                for source_route in incoming:
                    connections.append({
                        "from": source_route.get("from"), "to": bypass_target,
                        "label": "", "path_type": "main",
                    })
            removed.add(gateway_id)

    kept_stages = [stage for stage in stages if stage["stage_id"] not in removed]
    kept_ids = {stage["stage_id"] for stage in kept_stages}
    connections = [
        connection for connection in connections
        if connection.get("from") in kept_ids and connection.get("to") in kept_ids
    ]
    stage_by_id = {stage["stage_id"]: stage for stage in kept_stages}

    # Only genuine business decisions retain Yes/No. Positive continues through
    # the main flow. Negative terminates in the upper stop branch.
    outgoing_index = defaultdict(list)
    for connection in connections:
        outgoing_index[connection["from"]].append(connection)
    for stage in kept_stages:
        if stage.get("node_type") != "Gateway":
            continue
        routes = outgoing_index.get(stage["stage_id"], [])
        for route in routes:
            target = stage_by_id.get(route.get("to"), {})
            target_text = f"{target.get('heading', '')} {target.get('name', '')}".lower()
            negative = (
                target.get("lane") == "upper"
                or target.get("visual_role") == "exception"
                or any(term in target_text for term in (
                    "stop", "terminate", "no file", "no data", "not received",
                    "without sending", "reject", "failed",
                ))
            )
            route["label"] = "No" if negative else "Yes"

    # Rebuild the lower operational sequence deterministically. This prevents
    # duplicate, crossed, or orphaned arrows when optional gateways are collapsed.
    lower_stages = [stage for stage in kept_stages if stage.get("lane") == "lower"]
    lower_ids = {stage["stage_id"] for stage in lower_stages}
    if lower_stages:
        connections = [
            connection for connection in connections
            if connection.get("from") not in lower_ids
            and connection.get("to") not in lower_ids
        ]
        main_receivers = [
            stage for stage in kept_stages
            if stage.get("lane", "main") == "main" and stage.get("node_type") == "Receiver"
        ]
        main_ends = [
            stage for stage in kept_stages
            if stage.get("lane", "main") == "main" and stage.get("node_type") == "End"
        ]
        # Prefer the receiver as the operational anchor. If no receiver exists,
        # use the last main-row business action before End.
        main_actions = [
            stage for stage in kept_stages
            if stage.get("lane", "main") == "main"
            and stage.get("node_type") not in {"End", "Gateway"}
        ]
        anchor_stage = main_receivers[-1] if main_receivers else (main_actions[-1] if main_actions else None)
        if anchor_stage:
            connections.append({
                "from": anchor_stage["stage_id"],
                "to": lower_stages[0]["stage_id"],
                "label": "", "path_type": "optional",
            })
        for first, second in zip(lower_stages, lower_stages[1:]):
            connections.append({
                "from": first["stage_id"], "to": second["stage_id"],
                "label": "", "path_type": "optional",
            })
        if main_ends:
            connections.append({
                "from": lower_stages[-1]["stage_id"],
                "to": main_ends[-1]["stage_id"],
                "label": "", "path_type": "optional",
            })

    # Normalize completion: one business diagram has exactly one visible END.
    main_ends = [
        stage for stage in kept_stages
        if stage.get("node_type") == "End" and stage.get("lane", "main") == "main"
    ]
    canonical_end = main_ends[-1] if main_ends else next(
        (stage for stage in kept_stages if stage.get("node_type") == "End"), None
    )
    if canonical_end:
        canonical_end["lane"] = "main"
        canonical_end["visual_role"] = "completion"
        canonical_end["heading"] = "End"
        canonical_end["name"] = "END"
        duplicate_end_ids = {
            stage["stage_id"] for stage in kept_stages
            if stage.get("node_type") == "End" and stage["stage_id"] != canonical_end["stage_id"]
        }
        for connection in connections:
            if connection.get("to") in duplicate_end_ids:
                connection["to"] = canonical_end["stage_id"]
        kept_stages = [stage for stage in kept_stages if stage["stage_id"] not in duplicate_end_ids]
        connections = [
            connection for connection in connections
            if connection.get("from") not in duplicate_end_ids
        ]

    # Restore the optional source-copy branch as a real decision branch.
    # The business sequence is: logging decision -> Yes -> save copy -> data
    # validation; logging decision -> No -> data validation.
    stage_by_id = {stage["stage_id"]: stage for stage in kept_stages}
    copy_stage = next((
        stage for stage in kept_stages
        if any(term in f"{stage.get('heading', '')} {stage.get('name', '')}".lower()
               for term in ("save a copy", "source copy", "for records", "retain copy"))
    ), None)
    lower_gateway = next((
        stage for stage in kept_stages
        if stage.get("node_type") == "Gateway"
        and any(term in f"{stage.get('heading', '')} {stage.get('name', '')}".lower()
                for term in ("logging", "save copy", "source copy", "audit"))
    ), None)
    validation_gateway = next((
        stage for stage in kept_stages
        if stage.get("node_type") == "Gateway"
        and stage is not lower_gateway
        and any(term in f"{stage.get('heading', '')} {stage.get('name', '')}".lower()
                for term in ("record", "data available", "file present", "records exist"))
    ), None)

    if lower_gateway and copy_stage and validation_gateway:
        lower_gateway["lane"] = "lower"
        copy_stage["lane"] = "lower"
        copy_stage["optional"] = True
        involved = {lower_gateway["stage_id"], copy_stage["stage_id"]}
        incoming_to_logging = [
            c for c in connections
            if c.get("to") == lower_gateway["stage_id"] and c.get("from") not in involved
        ]
        connections = [
            c for c in connections
            if c.get("from") not in involved and c.get("to") not in involved
        ]
        if not incoming_to_logging:
            main_before_validation = [
                stage for stage in kept_stages
                if stage.get("lane", "main") == "main"
                and stage.get("node_type") not in {"End", "Gateway"}
            ]
            if main_before_validation:
                incoming_to_logging = [{
                    "from": main_before_validation[-1]["stage_id"],
                    "to": lower_gateway["stage_id"], "label": "", "path_type": "optional"
                }]
        connections.extend(incoming_to_logging)
        connections.extend([
            {"from": lower_gateway["stage_id"], "to": copy_stage["stage_id"],
             "label": "Yes", "path_type": "optional"},
            {"from": lower_gateway["stage_id"], "to": validation_gateway["stage_id"],
             "label": "No", "path_type": "main"},
            {"from": copy_stage["stage_id"], "to": validation_gateway["stage_id"],
             "label": "", "path_type": "optional"},
        ])

    # Remove duplicate connections after endpoint normalization.
    deduped = []
    seen_connections = set()
    for connection in connections:
        key = (connection.get("from"), connection.get("to"), connection.get("label", ""))
        if key in seen_connections or key[0] == key[1]:
            continue
        seen_connections.add(key)
        deduped.append(connection)
    connections = deduped

    # All nonconditional operational arrows are intentionally unlabeled.
    gateway_ids = {stage["stage_id"] for stage in kept_stages if stage.get("node_type") == "Gateway"}
    for connection in connections:
        if connection.get("from") not in gateway_ids:
            connection["label"] = ""

    return kept_stages, connections


def _repair_business_connections(stages, connections):
    """Guarantee visible connectivity after optional gateway simplification."""
    stage_by_id = {stage["stage_id"]: stage for stage in stages}
    connections = [dict(c) for c in connections
                   if c.get("from") in stage_by_id and c.get("to") in stage_by_id]

    def indexes():
        outgoing = defaultdict(list)
        incoming = defaultdict(list)
        for connection in connections:
            outgoing[connection["from"]].append(connection)
            incoming[connection["to"]].append(connection)
        return outgoing, incoming

    main = [s for s in stages if s.get("lane", "main") == "main"]
    upper = [s for s in stages if s.get("lane") == "upper"]
    lower = [s for s in stages if s.get("lane") == "lower"]
    outgoing, incoming = indexes()

    # Fill missing links along the declared business sequence only when no path
    # already exists between adjacent stages.
    for first, second in zip(main, main[1:]):
        if first.get("node_type") == "End":
            continue
        if not outgoing[first["stage_id"]]:
            connections.append({"from": first["stage_id"], "to": second["stage_id"],
                                "label": "", "path_type": "main"})
    outgoing, incoming = indexes()

    main_end = next((s for s in reversed(main) if s.get("node_type") == "End"), None)
    anchor_stage = next((s for s in reversed(main)
                         if s.get("node_type") not in {"End", "Gateway"}), None)
    if lower:
        if not incoming[lower[0]["stage_id"]] and anchor_stage:
            connections.append({"from": anchor_stage["stage_id"], "to": lower[0]["stage_id"],
                                "label": "", "path_type": "optional"})
        for first, second in zip(lower, lower[1:]):
            outgoing, incoming = indexes()
            if not outgoing[first["stage_id"]]:
                connections.append({"from": first["stage_id"], "to": second["stage_id"],
                                    "label": "", "path_type": "optional"})
        outgoing, incoming = indexes()
        if main_end and not outgoing[lower[-1]["stage_id"]]:
            connections.append({"from": lower[-1]["stage_id"], "to": main_end["stage_id"],
                                "label": "", "path_type": "optional"})

    # A genuine gateway must show both outcomes. Use upper exception as No and
    # the next main stage as Yes when an LLM omitted one connection.
    outgoing, incoming = indexes()
    for index, stage in enumerate(main):
        if stage.get("node_type") != "Gateway":
            continue
        routes = outgoing[stage["stage_id"]]
        targets = {route["to"] for route in routes}
        negative = next((s for s in upper if s["stage_id"] not in targets), None)
        positive = main[index + 1] if index + 1 < len(main) else main_end
        if positive and positive["stage_id"] not in targets:
            connections.append({"from": stage["stage_id"], "to": positive["stage_id"],
                                "label": "Yes", "path_type": "main"})
        if negative and negative["stage_id"] not in targets:
            connections.append({"from": stage["stage_id"], "to": negative["stage_id"],
                                "label": "No", "path_type": "alternate"})

    return connections


def _draw_business_system(ax, x, y, w, h, name, caption, color):
    """Draw a prominent source or target system card."""
    ax.add_patch(patches.FancyBboxPatch(
        (x - w / 2 + 0.06, y - h / 2 - 0.06), w, h,
        boxstyle="round,pad=0.02,rounding_size=0.10",
        facecolor="#16324F", edgecolor="none", alpha=0.10, zorder=3,
    ))
    ax.add_patch(patches.FancyBboxPatch(
        (x - w / 2, y - h / 2), w, h,
        boxstyle="round,pad=0.02,rounding_size=0.10",
        facecolor="#E8F4FB", edgecolor=color, linewidth=2.2, zorder=5,
    ))
    ax.add_patch(patches.FancyBboxPatch(
        (x - w / 2 + 0.08, y + h / 2 - 0.13), w - 0.16, 0.07,
        boxstyle="round,pad=0.005,rounding_size=0.03",
        facecolor=color, edgecolor="none", zorder=6,
    ))
    lines, font_size = _fit_story_text(name, w - 0.12, h - 0.16, False)
    ax.text(x, y + 0.04, "\n".join(lines), ha="center", va="center",
            fontsize=max(7.4, font_size), fontweight="bold", color="#13263A", zorder=7)


def _draw_compact_main_process(ax, iflow_data):
    """Render a professional business process flow without CPI implementation detail."""
    stages = iflow_data.get("diagram_stages") or []
    connections = iflow_data.get("diagram_connections") or []
    if not stages:
        return
    stages, connections = _simplify_business_diagram(stages, connections)
    if not stages:
        return
    connections = _repair_business_connections(stages, connections)

    for index, node in enumerate(stages):
        node.setdefault("stage_id", f"stage_{index + 1}")
        node.setdefault("lane", "main")
        # Terminal success/completion nodes always use one consistent label.
        # Exception End nodes keep their explanatory stop wording.
        if (
            node.get("node_type") == "End"
            and node.get("lane", "main") != "upper"
            and node.get("visual_role") != "exception"
        ):
            node["heading"] = "End"
            node["name"] = "END"
    stage_by_id = {node["stage_id"]: node for node in stages}

    sender_stages = [node for node in stages if node.get("node_type") == "Sender"]
    receiver_stages = [node for node in stages if node.get("node_type") in {"Receiver", "End"}]
    sender = sender_stages[0] if sender_stages else None

    # Keep secondary record-retention actions off the main row even if the LLM
    # assigned them to main. This prevents long flows from being clipped.
    receiver_present = any(node.get("node_type") == "Receiver" for node in stages)
    for node in stages:
        node_text = f"{node.get('heading', '')} {node.get('name', '')}".lower()
        is_record_copy = (
            node.get("visual_role") == "governance"
            or any(term in node_text for term in ("save a copy", "archive", "retain copy", "for records"))
        )
        if node.get("node_type") not in {"End", "Gateway"} and is_record_copy:
            node["lane"] = "lower"
            node["optional"] = True

    # Keep the first terminal target on the main row; branch End states remain branches.
    main = [node for node in stages if node.get("lane", "main") == "main"]
    upper = [node for node in stages if node.get("lane") == "upper"]
    lower = [node for node in stages if node.get("lane") == "lower"]
    if not main:
        main = stages
        upper = []
        lower = []

    main_y, upper_y, lower_y = 6.10, 8.25, 3.45
    gateway_size = 0.62
    minimum_gap = 0.30
    available_width = 13.45
    gateway_count = sum(1 for stage in main if stage.get("node_type") == "Gateway")
    box_count = max(1, len(main) - gateway_count)
    gap_count = max(0, len(main) - 1)
    remaining_for_boxes = (
        available_width
        - gateway_count * gateway_size * 2
        - gap_count * minimum_gap
    )
    box_w = min(1.88, max(1.42, remaining_for_boxes / box_count))

    def visual_width(stage):
        return gateway_size * 2 if stage.get("node_type") == "Gateway" else box_w

    # Geometry-aware spacing guarantees a visible gap between every pair.
    main_x = [0.75 + visual_width(main[0]) / 2] if main else []
    for previous, current in zip(main, main[1:]):
        distance = visual_width(previous) / 2 + minimum_gap + visual_width(current) / 2
        main_x.append(main_x[-1] + distance)

    # Final hard-fit check. Geometry must remain fully inside the canvas.
    if main_x:
        left_edge = main_x[0] - visual_width(main[0]) / 2
        right_edge = main_x[-1] + visual_width(main[-1]) / 2
        span = right_edge - left_edge
        maximum_span = 13.50
        if span > maximum_span and span > 0:
            scale = maximum_span / span
            center = (left_edge + right_edge) / 2
            main_x = [7.50 + (value - center) * scale for value in main_x]
            box_w *= scale
            gateway_size *= scale
        else:
            shift = 0.0
            if left_edge < 0.75:
                shift = 0.75 - left_edge
            elif right_edge > 14.25:
                shift = 14.25 - right_edge
            main_x = [value + shift for value in main_x]

    positions = {stage["stage_id"]: (x, main_y) for stage, x in zip(main, main_x)}

    # Stop outcomes sit directly above their source decision.
    for index, stage in enumerate(upper):
        incoming_route = next((c for c in connections if c.get("to") == stage["stage_id"]), None)
        source_x = positions.get((incoming_route or {}).get("from"), (7.5, main_y))[0]
        positions[stage["stage_id"]] = (source_x, upper_y + index * 0.12)

    # Lower operational actions are centered as a compact row with equal gaps.
    if lower:
        lower_gap = 0.48
        lower_box_w = box_w
        total_lower_width = len(lower) * lower_box_w + max(0, len(lower) - 1) * lower_gap
        if total_lower_width > 13.50:
            lower_box_w = max(1.35, (13.50 - max(0, len(lower) - 1) * 0.30) / len(lower))
            lower_gap = 0.30
            total_lower_width = len(lower) * lower_box_w + max(0, len(lower) - 1) * lower_gap
        receiver_candidates = [s for s in main if s.get("node_type") == "Receiver"]
        action_candidates = [s for s in main if s.get("node_type") not in {"End", "Gateway"}]
        anchor_candidate = receiver_candidates[-1] if receiver_candidates else (action_candidates[-1] if action_candidates else None)
        anchor_x = positions.get(anchor_candidate["stage_id"], (9.0, main_y))[0] if anchor_candidate else 9.0
        left = max(0.75, min(anchor_x - total_lower_width / 2, 14.25 - total_lower_width))
        for index, stage in enumerate(lower):
            positions[stage["stage_id"]] = (
                left + lower_box_w / 2 + index * (lower_box_w + lower_gap),
                lower_y,
            )
        box_w = min(box_w, lower_box_w)

    base_h = 1.22
    node_heights = {
        node["stage_id"]: _estimate_story_box_height(node.get("name", ""), box_w, base_h)
        for node in stages if node.get("node_type") != "Gateway"
    }

    # Draw subtle phase band behind the happy path.
    band_left = max(0.35, min(main_x) - 0.70) if main_x else 0.55
    band_right = min(14.65, max(main_x) + 0.70) if main_x else 14.45
    ax.add_patch(patches.FancyBboxPatch(
        (band_left, main_y - 1.05), band_right - band_left, 2.10,
        boxstyle="round,pad=0.02,rounding_size=0.12",
        facecolor="#F7FAFC", edgecolor="#D7E1EC", linewidth=1.0, zorder=1,
    ))

    for node in stages:
        position = positions.get(node["stage_id"])
        if not position:
            continue
        x, y = position
        if node.get("node_type") == "Gateway":
            _draw_story_gateway(ax, x, y, gateway_size, node)
        elif node is sender:
            _draw_business_system(ax, x, y, box_w, node_heights[node["stage_id"]],
                                  node.get("name", "Process Starts"), "", "#2475A7")
        elif node.get("node_type") in {"Receiver", "End"} and node.get("lane", "main") == "main":
            _draw_business_system(ax, x, y, box_w, node_heights[node["stage_id"]],
                                  node.get("name", "END"), "", "#2F855A")
        else:
            _draw_story_box(ax, x, y, box_w, node_heights[node["stage_id"]], node)

    for connection in connections:
        source = stage_by_id.get(connection.get("from"))
        target = stage_by_id.get(connection.get("to"))
        if not source or not target:
            continue
        source_pos = positions.get(source["stage_id"])
        target_pos = positions.get(target["stage_id"])
        if not source_pos or not target_pos:
            continue
        source_h = node_heights.get(source["stage_id"], base_h)
        target_h = node_heights.get(target["stage_id"], base_h)
        sx, sy, rx, ry = _stage_boundary(source_pos, source, box_w, source_h, gateway_size)
        lx, ly, tx, ty = _stage_boundary(target_pos, target, box_w, target_h, gateway_size)
        if abs(target_pos[1] - source_pos[1]) < 0.10:
            if target_pos[0] >= source_pos[0]:
                start_point, end_point = (rx, ry), (lx, ly)
            else:
                start_point, end_point = (sx, sy), (tx, ty)
        else:
            sign = 1 if target_pos[1] > source_pos[1] else -1
            source_offset = gateway_size if source.get("node_type") == "Gateway" else source_h / 2
            target_offset = gateway_size if target.get("node_type") == "Gateway" else target_h / 2
            start_point = (source_pos[0], source_pos[1] + sign * source_offset)
            end_point = (target_pos[0], target_pos[1] - sign * target_offset)
        _draw_story_connection(ax, start_point, end_point,
                               connection.get("label", ""), connection.get("path_type", "main"))


def _draw_exception_lane(ax, nodes, edges, x_left, x_right, y_center, panel_color=RED):
    """Draw an exception lane with evenly spaced nodes and guaranteed clear gaps."""
    if not nodes:
        return

    outgoing = {}
    incoming_count = {node_id: 0 for node_id in nodes}
    for edge in edges:
        source = edge.get("source")
        target = edge.get("target")
        if source in nodes and target in nodes:
            outgoing.setdefault(source, []).append(target)
            incoming_count[target] = incoming_count.get(target, 0) + 1

    starts = [node_id for node_id in nodes if incoming_count.get(node_id, 0) == 0]
    ordered_ids = []
    visited = set()
    queue = list(starts)

    while queue:
        node_id = queue.pop(0)
        if node_id in visited:
            continue
        visited.add(node_id)
        ordered_ids.append(node_id)
        queue.extend(outgoing.get(node_id, []))

    for node_id in nodes:
        if node_id not in visited:
            ordered_ids.append(node_id)

    count = len(ordered_ids)
    available_width = x_right - x_left
    gap = 0.42
    node_width = min(1.48, (available_width - gap * max(count - 1, 0)) / max(count, 1))
    node_width = max(1.12, node_width)
    node_height = 0.78

    if count == 1:
        x_positions = [(x_left + x_right) / 2]
    else:
        usable = available_width - node_width
        x_positions = [x_left + node_width / 2 + index * usable / (count - 1) for index in range(count)]

    coordinates = {
        node_id: (x_positions[index], y_center)
        for index, node_id in enumerate(ordered_ids)
    }

    for edge in edges:
        source = edge.get("source")
        target = edge.get("target")
        if source not in coordinates or target not in coordinates:
            continue
        x1, y1 = coordinates[source]
        x2, y2 = coordinates[target]
        ax.annotate(
            "",
            xy=(x2 - node_width / 2 - 0.05, y2),
            xytext=(x1 + node_width / 2 + 0.05, y1),
            arrowprops=dict(
                arrowstyle="-|>",
                color=f"#{panel_color}",
                lw=1.25,
                mutation_scale=10,
                shrinkA=0,
                shrinkB=0,
            ),
            zorder=3,
        )

    for node_id in ordered_ids:
        node = nodes[node_id]
        x, y = coordinates[node_id]
        _draw_card(ax, x, y, node_width, node_height, node, number=None, scale=0.92)


def draw_enterprise_diagram(iflow_data, output_img_path="cpi_flow_diagram.png"):
    os.makedirs(os.path.dirname(output_img_path) or ".", exist_ok=True)

    error_nodes = iflow_data.get("error_subprocess_nodes", {})
    error_edges = iflow_data.get("error_subprocess_edges", [])
    local_nodes = iflow_data.get("exception_flow_nodes", {})
    local_edges = iflow_data.get("exception_flow_edges", [])

    fig_w = 15.0
    fig_h = 10.0
    fig, ax = plt.subplots(figsize=(fig_w, fig_h), dpi=240)
    ax.set_xlim(0, fig_w)
    ax.set_ylim(0, fig_h)
    ax.axis("off")
    fig.patch.set_facecolor(f"#{WHITE}")


    ax.add_patch(patches.FancyBboxPatch(
        (0.25, 1.10), 14.5, fig_h - 1.40,
        boxstyle="round,pad=0.02,rounding_size=0.06",
        linewidth=1.1, edgecolor=f"#{WHITE}", facecolor=f"#{WHITE}", zorder=1,
    ))
    _draw_compact_main_process(ax, iflow_data)

    # Technical exception subprocesses are intentionally excluded from the
    # business process-flow diagram. They remain documented in technical sections.

    plt.savefig(
        output_img_path,
        bbox_inches="tight",
        pad_inches=0.08,
        facecolor=f"#{WHITE}",
    )
    plt.close(fig)
    return output_img_path


# ---------------------------------------------------------------------------
# Word helpers
# ---------------------------------------------------------------------------
def set_cell_bg(cell, color):
    cell._tc.get_or_add_tcPr().append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color}"/>'))


def set_cell_borders(cell, color=GRID, sz=4, sides=("top", "bottom", "left", "right")):
    tc_pr = cell._tc.get_or_add_tcPr()
    borders = OxmlElement("w:tcBorders")
    for side in sides:
        element = OxmlElement(f"w:{side}")
        element.set(qn("w:val"), "single")
        element.set(qn("w:sz"), str(sz))
        element.set(qn("w:color"), color)
        borders.append(element)
    tc_pr.append(borders)


def set_cell_margins(cell, top=90, bottom=90, left=140, right=140):
    cell._tc.get_or_add_tcPr().append(parse_xml(
        f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/>'
        f'<w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/>'
        f'<w:right w:w="{right}" w:type="dxa"/></w:tcMar>'
    ))


def vcenter(cell):
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER


def style_header_row(row, bg=NAVY):
    for cell in row.cells:
        set_cell_bg(cell, bg)
        set_cell_borders(cell, color=bg)
        set_cell_margins(cell, top=110, bottom=110)
        vcenter(cell)
        for run in cell.paragraphs[0].runs:
            run.font.bold = True
            run.font.size = Pt(9.5)
            run.font.color.rgb = rgb(WHITE)
            run.font.name = FONT


def zebra_rows(table, start=1, label_col_bold=True):
    for row_index, row in enumerate(table.rows[start:]):
        bg = WHITE if row_index % 2 == 0 else BG_LIGHT
        for column_index, cell in enumerate(row.cells):
            set_cell_bg(cell, bg)
            set_cell_borders(cell)
            set_cell_margins(cell)
            vcenter(cell)
            for paragraph in cell.paragraphs:
                paragraph.paragraph_format.space_before = Pt(1)
                paragraph.paragraph_format.space_after = Pt(1)
                for run in paragraph.runs:
                    run.font.size = Pt(9)
                    run.font.name = FONT
                    if label_col_bold and column_index == 0:
                        run.font.bold = True
                        run.font.color.rgb = rgb(SLATE)


def add_section_heading(doc, number, title):
    paragraph = doc.add_paragraph(style=doc.styles["Heading 1"])
    paragraph.paragraph_format.space_before = Pt(20)
    paragraph.paragraph_format.space_after = Pt(8)
    paragraph.paragraph_format.keep_with_next = True
    p_pr = paragraph._p.get_or_add_pPr()
    p_bdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "8")
    bottom.set(qn("w:space"), "4")
    bottom.set(qn("w:color"), ORANGE)
    p_bdr.append(bottom)
    p_pr.append(p_bdr)
    run = paragraph.add_run(f"{number}  ")
    run.font.color.rgb = rgb(ORANGE)
    run.font.bold = True
    run.font.size = Pt(14)
    run.font.name = FONT
    run = paragraph.add_run(title.upper())
    run.font.color.rgb = rgb(NAVY)
    run.font.bold = True
    run.font.size = Pt(14)
    run.font.name = FONT
    return paragraph


def add_sub_heading(doc, text):
    paragraph = doc.add_paragraph(style=doc.styles["Heading 2"])
    paragraph.paragraph_format.space_before = Pt(10)
    paragraph.paragraph_format.space_after = Pt(4)
    run = paragraph.add_run(text)
    run.font.color.rgb = rgb(NAVY_SOFT)
    run.font.bold = True
    run.font.size = Pt(11.5)
    run.font.name = FONT
    return paragraph


def body_text(doc, text):
    paragraph = doc.add_paragraph()
    paragraph.paragraph_format.space_after = Pt(8)
    run = paragraph.add_run(str(text))
    run.font.size = Pt(10)
    run.font.name = FONT
    run.font.color.rgb = rgb(TEXT_BODY)
    return paragraph


def add_bullets(doc, items):
    for item in items:
        paragraph = doc.add_paragraph(style="List Bullet")
        paragraph.paragraph_format.space_after = Pt(4)
        run = paragraph.add_run(str(item))
        run.font.size = Pt(10)
        run.font.name = FONT
        run.font.color.rgb = rgb(TEXT_BODY)


def set_col_widths(table, widths):
    table.autofit = False
    for index, width in enumerate(widths):
        table.columns[index].width = Inches(width)
    for row in table.rows:
        for index, width in enumerate(widths):
            if index < len(row.cells):
                row.cells[index].width = Inches(width)


def default_col_widths(count):
    if count == 2:
        return [2.3, CONTENT_WIDTH_IN - 2.3]
    width = CONTENT_WIDTH_IN / count
    return [width] * count


def new_kv_table(doc, rows, ncols=2, col_widths=None):
    table = doc.add_table(rows=0, cols=ncols)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for values in rows:
        cells = table.add_row().cells
        for index, value in enumerate(values):
            cells[index].text = str(value)
    zebra_rows(table, start=0)
    set_col_widths(table, col_widths or default_col_widths(ncols))
    return table


def new_headed_table(doc, headers, rows, col_widths=None):
    table = doc.add_table(rows=1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for index, header in enumerate(headers):
        table.rows[0].cells[index].text = header
    style_header_row(table.rows[0])
    for values in rows:
        cells = table.add_row().cells
        for index, value in enumerate(values):
            cells[index].text = str(value)
    zebra_rows(table, start=1, label_col_bold=False)
    set_col_widths(table, col_widths or default_col_widths(len(headers)))
    return table


def add_field(run, field_string, default_text=""):
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instruction = OxmlElement("w:instrText")
    instruction.set(qn("xml:space"), "preserve")
    instruction.text = f" {field_string} "
    separator = OxmlElement("w:fldChar")
    separator.set(qn("w:fldCharType"), "separate")
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    run._r.append(begin)
    run._r.append(instruction)
    run._r.append(separator)
    if default_text:
        text = OxmlElement("w:t")
        text.text = default_text
        run._r.append(text)
    run._r.append(end)


def add_toc(doc):
    paragraph = doc.add_paragraph()
    add_field(paragraph.add_run(), 'TOC \\o "1-2" \\h \\z \\u',
              "Right-click here and choose 'Update Field' to build the table of contents.")


def set_update_fields_on_open(doc):
    element = OxmlElement("w:updateFields")
    element.set(qn("w:val"), "true")
    doc.settings.element.append(element)


# ---------------------------------------------------------------------------
# DOCX builder
# ---------------------------------------------------------------------------
def build_visteon_ts_docx(iflow_data, logo_path=PRIMARY_LOGO_PATH, output_dir="."):
    doc = Document()
    iflow_data = apply_user_metadata(iflow_data)
    iflow_name = (iflow_data.get("iflow_name") or "SAP_CPI_IFLOW").strip()
    safe_name = sanitize_output_name(iflow_name)
    output_dir = os.path.abspath(output_dir or ".")
    os.makedirs(output_dir, exist_ok=True)
    output_doc_path = os.path.join(output_dir, f"UID_U057_{safe_name}.docx")
    logo = resolve_asset_path(logo_path)

    for section in doc.sections:
        section.top_margin = Inches(0.75)
        section.bottom_margin = Inches(0.75)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)

        header_table = section.header.add_table(rows=1, cols=2, width=Inches(6.9))
        header_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        left, right = header_table.rows[0].cells
        set_cell_borders(left, color=NAVY, sz=6, sides=("bottom",))
        set_cell_borders(right, color=NAVY, sz=6, sides=("bottom",))
        if logo and os.path.exists(logo):
            left.paragraphs[0].add_run().add_picture(logo, width=Inches(1.6))
        else:
            run = left.paragraphs[0].add_run("VISTEON")
            run.font.bold = True
            run.font.size = Pt(14)
            run.font.color.rgb = rgb(NAVY)
        right.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
        run = right.paragraphs[0].add_run(iflow_name[:44])
        run.font.size = Pt(9)
        run.font.bold = True
        run.font.color.rgb = rgb(NAVY)
        right.paragraphs[0].add_run().add_break()
        run = right.paragraphs[0].add_run("Technical Specification")
        run.font.size = Pt(7.5)
        run.font.color.rgb = rgb(SLATE)

        footer_table = section.footer.add_table(rows=1, cols=2, width=Inches(6.9))
        footer_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        f_left, f_right = footer_table.rows[0].cells
        set_cell_borders(f_left, color=GRID, sz=6, sides=("top",))
        set_cell_borders(f_right, color=GRID, sz=6, sides=("top",))
        run = f_left.paragraphs[0].add_run("Visteon Internal & Confidential  |  SAP Integration Suite")
        run.font.size = Pt(7.5)
        run.font.color.rgb = rgb("808080")
        f_right.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
        run = f_right.paragraphs[0].add_run("Page ")
        run.font.size = Pt(7.5)
        add_field(run, "PAGE", "1")
        run = f_right.paragraphs[0].add_run(" of ")
        run.font.size = Pt(7.5)
        add_field(run, "NUMPAGES", "1")

    doc.styles["Normal"].font.name = FONT
    doc.styles["Normal"].font.size = Pt(10)
    doc.styles["Normal"].font.color.rgb = rgb(TEXT_BODY)
    for style_name in ("Heading 1", "Heading 2", "Heading 3"):
        doc.styles[style_name].font.name = FONT
        doc.styles[style_name].font.color.rgb = rgb(NAVY)

    cover = doc.add_paragraph()
    cover.paragraph_format.space_before = Pt(130)
    if logo and os.path.exists(logo):
        cover.add_run().add_picture(logo, width=Inches(2.1))
    else:
        run = cover.add_run("VISTEON")
        run.font.bold = True
        run.font.size = Pt(22)
        run.font.color.rgb = rgb(NAVY)

    kicker = doc.add_paragraph()
    kicker.paragraph_format.space_before = Pt(36)
    run = kicker.add_run("SAP INTEGRATION SUITE  |  TECHNICAL SPECIFICATION")
    run.font.size = Pt(11)
    run.font.bold = True
    run.font.color.rgb = rgb(ORANGE)
    title = doc.add_paragraph()
    title.paragraph_format.space_before = Pt(6)
    run = title.add_run(iflow_name)
    run.font.size = Pt(26)
    run.font.bold = True
    run.font.color.rgb = rgb(NAVY)
    subtitle = doc.add_paragraph()
    run = subtitle.add_run(f"{iflow_data['sender_system']}  ->  {iflow_data['receiver_system']}")
    run.font.size = Pt(13)
    run.font.color.rgb = rgb(SLATE)
    doc.add_paragraph().paragraph_format.space_after = Pt(60)

    new_kv_table(doc, [
        ("Integration Package", iflow_data["package_name"]),
        ("Document Version", iflow_data["doc_version"]),
        ("Prepared By", iflow_data["prepared_by"]),
        ("Reviewed By", iflow_data["reviewed_by"]),
        ("Approved By", iflow_data["approved_by"]),
        ("Effective Date", iflow_data["effective_date"]),
        ("Classification", "Visteon Internal & Confidential"),
    ])
    doc.add_page_break()

    toc_head = doc.add_paragraph()
    run = toc_head.add_run("TABLE OF CONTENTS")
    run.font.size = Pt(16)
    run.font.bold = True
    run.font.color.rgb = rgb(NAVY)
    add_toc(doc)
    doc.add_page_break()

    add_section_heading(doc, "", "Revision History")
    new_headed_table(doc,
        ["Version", "Effective Date", "Brief Description", "Change Ref", "Affected Section", "Prepared By", "Reviewed By", "Approved By"],
        [[iflow_data["doc_version"], iflow_data["effective_date"], iflow_data["description"], "-", "All Sections", iflow_data["prepared_by"], iflow_data["reviewed_by"], iflow_data["approved_by"]]],
    )

    add_section_heading(doc, "1.", "Overview")
    body_text(doc, iflow_data["business_overview"])

    add_section_heading(doc, "2.", "Interface Flow")
    scheduled = "Scheduled" in iflow_data.get("execution_mode", "")
    new_kv_table(doc, [
        ("Source System Name", iflow_data["sender_system"]),
        ("Target System Name", iflow_data["receiver_system"]),
        ("Direction of Interface", iflow_data.get("direction", "Inbound")),
        ("Execution Mode", iflow_data["execution_mode"]),
        ("Synchronous / Asynchronous", iflow_data["synchronous_asynchronous"]),
        ("Source Transport Protocol", iflow_data["sender_adapter_type"]),
        ("Target Transport Protocol", iflow_data["receiver_adapter_type"]),
        ("Triggering Events", "Schedule / Polling" if scheduled else "Inbound Request"),
        ("Frequency of Run", iflow_data["frequency"]),
        ("Volume", "As per operational schedule"),
        ("Acknowledgement", "No"),
    ])

    add_section_heading(doc, "3.", "Assumptions")
    add_bullets(doc, [
        f"Data is processed from {iflow_data['sender_system']} to {iflow_data['receiver_system']} according to the configured iFlow routes and conditions."
    ])

    add_section_heading(doc, "4.", "Interface Description")
    body_text(doc, iflow_data["interface_description"])

    add_section_heading(doc, "5.", "Technical Artifacts")
    mapping_text = "\n".join(iflow_data["xslt_scripts"] + iflow_data["mappings"]) or "None"
    new_kv_table(doc, [
        ("Integration Package", iflow_data["package_name"]),
        ("Integration iFlow", iflow_name),
        ("Sender System", iflow_data["sender_system"]),
        ("Receiver System", iflow_data["receiver_system"]),
        ("XSLT / Message Mappings", mapping_text),
        ("Groovy Scripts", format_groovy_artifacts(iflow_data)),
        ("Business Identifier", "Application ID / Track ID"),
    ])

    add_section_heading(doc, "6.", "High Level Processing Logic")
    add_bullets(doc, iflow_data.get("processing_logic", []))

    add_section_heading(doc, "7.", "Process Flow Diagram")
    diagram_path = os.path.join(output_dir, f"diag_{safe_name}.png")
    draw_enterprise_diagram(iflow_data, diagram_path)
    doc.add_picture(diagram_path, width=Inches(6.6))

    add_section_heading(doc, "8.", "Connectivity & Authentications")
    connectivity_items = _build_connectivity_rows(iflow_data)
    # Render as numbered items with subpoints (Adapter Type, Authentication, Purpose, Details)
    def _to_roman(n):
        vals = [
            (1000, "M"), (900, "CM"), (500, "D"), (400, "CD"),
            (100, "C"), (90, "XC"), (50, "L"), (40, "XL"),
            (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I"),
        ]
        out = ""
        for val, sym in vals:
            while n >= val:
                out += sym
                n -= val
        return out

    def _resolve_value(raw, global_props):
        if raw is None:
            return ""
        text = str(raw).strip()
        # detect {{VAR}} or ${VAR} patterns and replace from global properties when possible
        m = re.search(r"\{\{\s*([^\}]+)\s*\}\}", text)
        if not m:
            m = re.search(r"\$\{\s*([^\}]+)\s*\}\}", text)
        if m:
            key = m.group(1).strip()
            val = global_props.get(key)
            if val:
                return val
            # try dot/underscore variants
            val = global_props.get(key.replace(".", "_").replace("-", "_"))
            if val:
                return val
            # fallback to raw token
            return key
        return text

    # mapping of detail keys we consider important with friendly labels
    # Order matters: show Port before Host as requested
    important_keys = [
        ("port", "Port"), ("host", "Host"), ("hostname", "Host"),
        ("directory", "Directory"), ("remotepath", "Remote Path"), ("remote_dir", "Remote Directory"),
        ("folder", "Folder"),
        ("proxy", "Proxy Type"), ("proxytype", "Proxy Type"), ("proxy_host", "Proxy Host"), ("proxyport", "Proxy Port"),
        ("user", "Username"), ("username", "Username"), ("user_name", "Username"),
    ]

    global_props = iflow_data.get("properties", {}) or {}
    for idx, item in enumerate(connectivity_items, start=1):
        roman = _to_roman(idx)
        # Roman heading (bold)
        p = doc.add_paragraph()
        run = p.add_run(f"{roman}. {item.get('title', '')}")
        run.bold = True
        run.font.size = Pt(11)
        run.font.name = FONT
        run.font.color.rgb = rgb(TEXT_BODY)

        # Prepare bullet subpoints with a left indent so they appear under the roman heading
        bullets = []
        bullets.append(("Adapter Type", item.get("adapter_type", "-")))
        bullets.append(("Authentication", item.get("authentication", "-")))
        bullets.append(("Purpose", item.get("purpose", "-")))

        details = item.get("details") or {}
        # pick important configured values only, deduplicate and prefer resolved configured values
        picked = []
        seen_labels = set()
        lower_keys = {k.lower(): k for k in details.keys()}
        for key_token, label in important_keys:
            # find the first matching detail key
            match_key = None
            for present_lower, orig_key in lower_keys.items():
                if key_token in present_lower:
                    match_key = orig_key
                    break
            if not match_key:
                continue
            raw_val = details.get(match_key)
            resolved = _resolve_value(raw_val, global_props)
            # If resolution only returned the token/key name (meaning not configured), skip
            if not resolved or resolved == match_key or re.search(r"^[A-Za-z0-9_\- ]+_Receiver|_Sender|_Receiver_|Receiver_", str(resolved)):
                continue
            # mask sensitive keys
            lk = match_key.lower()
            if any(secret in lk for secret in ("password", "secret", "private", "key")):
                resolved = "***"
            if label not in seen_labels:
                picked.append((label, resolved))
                seen_labels.add(label)
        # append picked details after bullets
        for label, val in picked:
            bullets.append((label, val))

        for label, val in bullets:
            p2 = doc.add_paragraph(style="List Bullet")
            p2.paragraph_format.left_indent = Inches(0.25)
            p2.paragraph_format.space_after = Pt(2)
            run2 = p2.add_run(f"{label} - {val}")
            run2.font.size = Pt(10)
            run2.font.name = FONT
            run2.font.color.rgb = rgb(TEXT_BODY)

    add_section_heading(doc, "9.", "Data Mapping Matrix")
    mapping_rows = _summarize_mapping_artifacts(iflow_data)
    if mapping_rows:
        new_headed_table(doc, ["Mapping Artifact", "Type", "Purpose"], mapping_rows)
    else:
        body_text(doc, "No XSLT or message-mapping artifact was detected in the uploaded iFlow package.")

    add_section_heading(doc, "11.", "Interface Alert Notification")
    success_feature = iflow_data.get("conditional_features", {}).get("success_email")
    alert_rows = [[
        "Error Event Subprocess",
        "Invokes the local exception process, constructs a structured error message, and publishes the error to SAP Event Mesh through AMQP.",
    ]]
    if success_feature:
        alert_rows.append([
            "Success Notification",
            "Conditionally invokes the common success-email integration through ProcessDirect when SendSuccessfullNotification is TRUE.",
        ])
    new_headed_table(doc, ["Alert Parameter", "Configuration Details"], alert_rows)

    add_section_heading(doc, "12.", "Error Conditions, Message Retry & Reprocessing")
    new_kv_table(doc, [
        ("Retry Configuration", "No explicit application-level retry subprocess detected in the parsed BPMN."),
        ("Reprocessing", "Use SAP Integration Suite monitoring and the approved operational reprocessing procedure."),
    ])

    add_section_heading(doc, "13.", "Unit Test Documents")
    body_text(doc, f"UT_DOCUMENT_U057_{iflow_name}.docx")

    add_section_heading(doc, "14.", "Interface Monitoring Procedures")
    body_text(doc, "Monitor message execution status, Message Processing Logs, Data Store activity, receiver delivery, and exception publications in SAP Integration Suite monitoring.")

    add_section_heading(doc, "15.", "Project References")
    new_headed_table(doc, ["Sr. No.", "Document Name", "Reference Link"], [["", "", ""]])

    add_section_heading(doc, "16.", "Comments & Approvals")
    new_headed_table(doc, ["Date", "Raised By", "Comment", "Status"], [["", "", "", ""]] * 3)

    set_update_fields_on_open(doc)
    doc.save(output_doc_path)
    print(f"Generated: {output_doc_path}")
    return output_doc_path


# ---------------------------------------------------------------------------
# Batch runner
# ---------------------------------------------------------------------------
def process_all_cpi_zips(directory_path):
    if not os.path.exists(directory_path):
        print(f"Target directory '{directory_path}' does not exist.")
        return []
    zip_files = glob.glob(os.path.join(directory_path, "*.zip"))
    if not zip_files:
        print(f"No .zip files found in {directory_path}")
        return []
    outputs = []
    print(f"Found {len(zip_files)} CPI ZIP file(s).")
    for zip_file in zip_files:
        parsed = parse_cpi_iflow_zip(zip_file)
        outputs.append(build_visteon_ts_docx(parsed, output_dir=directory_path))
    return outputs


if __name__ == "__main__":
    process_all_cpi_zips(ZIP_DIRECTORY_PATH)
