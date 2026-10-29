import sys
import unittest
from unittest.mock import MagicMock, patch

from utils.credentials import read_windows_credential


class CredentialTests(unittest.TestCase):
    def test_reads_generic_credential(self):
        api = MagicMock()
        api.CRED_TYPE_GENERIC = 1
        api.CredRead.return_value = {
            "UserName": "test-user", "CredentialBlob": "test-password".encode("utf-16-le")
        }
        with patch.dict(sys.modules, {"win32cred": api}):
            self.assertEqual(read_windows_credential("JDE"), ("test-user", "test-password"))
        api.CredRead.assert_called_once_with("JDE", 1, 0)

    def test_read_error_has_no_credential_content(self):
        api = MagicMock()
        api.CredRead.side_effect = RuntimeError("sensitive detail")
        with patch.dict(sys.modules, {"win32cred": api}):
            with self.assertRaises(RuntimeError) as caught:
                read_windows_credential("JDE")
        self.assertNotIn("sensitive detail", str(caught.exception))

    def test_empty_target(self):
        with self.assertRaises(ValueError):
            read_windows_credential("")
