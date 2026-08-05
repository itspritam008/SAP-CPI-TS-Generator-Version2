import unittest

from generator import apply_user_metadata, apply_llm_interface_intelligence, build_processing_logic_from_pipeline


class ApplyUserMetadataTests(unittest.TestCase):
    def test_applies_enterprise_metadata_and_avoids_test_placeholders(self):
        iflow_data = {
            "sender_system": "test",
            "receiver_system": "test",
            "direction": "Inbound",
            "synchronous_asynchronous": "Synchronous",
            "description": "",
        }

        user_inputs = {
            "prepared_by": "Jane Doe",
            "reviewed_by": "John Smith",
            "approved_by": "Alex Chen",
            "effective_date": "2026-07-30",
            "direction": "Outbound",
            "sync_async": "Asynchronous",
            "description": "Customer payment file exchange",
            "source_system": "SAP S/4HANA",
            "target_system": "Fifth Third Bank",
        }

        result = apply_user_metadata(iflow_data, user_inputs)

        self.assertEqual(result["prepared_by"], "Jane Doe")
        self.assertEqual(result["reviewed_by"], "John Smith")
        self.assertEqual(result["approved_by"], "Alex Chen")
        self.assertEqual(result["effective_date"], "2026-07-30")
        self.assertEqual(result["direction"], "Outbound")
        self.assertEqual(result["synchronous_asynchronous"], "Asynchronous")
        self.assertEqual(result["description"], "Customer payment file exchange")
        self.assertEqual(result["sender_system"], "SAP S/4HANA")
        self.assertEqual(result["receiver_system"], "Fifth Third Bank")

    def test_builds_processing_logic_from_pipeline_nodes(self):
        iflow_data = {
            "sender_system": "SAP S/4HANA",
            "receiver_system": "Fifth Third Bank",
            "main_pipeline_nodes": [
                ("Validate Payload", "Groovy Script"),
                ("Map Payment Data", "Data Mapping"),
                ("Invoke Bank Endpoint", "Processing Step"),
            ],
        }

        result = build_processing_logic_from_pipeline(iflow_data)

        self.assertIn("Validate Payload", result)
        self.assertIn("Map Payment Data", result)
        self.assertIn("Invoke Bank Endpoint", result)
        self.assertIn("SAP S/4HANA", result)
        self.assertIn("Fifth Third Bank", result)

    def test_applies_llm_interface_intelligence(self):
        iflow_data = {
            "sender_system": "Source System",
            "receiver_system": "Target System",
            "sender_adapter_type": "HTTPS",
            "receiver_adapter_type": "REST",
            "description": "",
        }

        payload = {
            "sender_system": "SAP S/4HANA",
            "sender_adapter_type": "IDoc",
            "receiver_system": "Fifth Third Bank",
            "receiver_adapter_type": "HTTPS",
            "execution_mode": "Real-time / Event-driven",
            "business_overview": "This integration synchronizes payment status updates from SAP S/4HANA to Fifth Third Bank for settlement processing.",
            "interface_description": "The interface receives payment status notifications and delivers them through a secure HTTPS callout to the bank.",
            "processing_logic": "The flow validates the payload, maps the fields, and calls the target endpoint after enrichment.",
        }

        result = apply_llm_interface_intelligence(iflow_data, payload)

        self.assertEqual(result["sender_system"], "SAP S/4HANA")
        self.assertEqual(result["sender_adapter_type"], "IDoc")
        self.assertEqual(result["receiver_system"], "Fifth Third Bank")
        self.assertEqual(result["receiver_adapter_type"], "HTTPS")
        self.assertEqual(result["business_overview"], payload["business_overview"])
        self.assertEqual(result["interface_description"], payload["interface_description"])
        self.assertEqual(result["processing_logic"], payload["processing_logic"])


if __name__ == "__main__":
    unittest.main()
