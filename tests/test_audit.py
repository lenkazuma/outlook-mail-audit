"""Offline regression tests: all mailbox data and clients are synthetic."""
import csv
import io
import json
import os
import subprocess
import sys
import tempfile
import unittest
from datetime import UTC, datetime, timedelta, timezone
from pathlib import Path
from types import SimpleNamespace as NS
from unittest.mock import Mock, patch
import audit_utils as utils
import graph_client as api
import scan_readonly as graph
import scan_outlook_local_readonly as local


class Reports(unittest.TestCase):
    def test_formula_cells_headers_and_unicode(self):
        values = ['=1+1', '+SUM(A1)', '-1+2', '@SUM(A1)', '  =1', '\ttext', '\rtext', '\ntext']
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / 'report.csv'
            utils.write_csv(path, [{'subject': v, '=header': '中文', 'count': -2} for v in values])
            with path.open(encoding='utf-8-sig', newline='') as stream:
                rows = list(csv.DictReader(stream))
            self.assertEqual([r['subject'] for r in rows], ["'" + v for v in values])
            self.assertEqual(rows[0]["'=header"], '中文')
            self.assertEqual(rows[0]['count'], '-2')
            utils.write_csv(path, [])
            self.assertEqual(path.read_bytes(), b'')

    def test_positive_parameters(self):
        self.assertEqual(utils.positive_int('7'), 7)
        for value in ['0', '-10', 'invalid']:
            with self.assertRaises((ValueError, __import__('argparse').ArgumentTypeError)):
                utils.positive_int(value)


class GraphClient(unittest.TestCase):
    def setUp(self):
        self.client = api.GraphClient.__new__(api.GraphClient)
        self.client._access_token = 'synthetic-test-token'
        self.client._requests = Mock()

    def test_only_read_scopes_and_reserved_scope_compatibility(self):
        self.assertEqual(api.normalize_scopes(['Mail.Read', 'offline_access', 'openid', 'profile']), ['Mail.Read'])
        self.assertEqual(api.normalize_scopes(None), api.DEFAULT_SCOPES)
        for scopes in [['Mail.ReadWrite'], ['offline_access'], ['https://graph.microsoft.com/.default']]:
            with self.assertRaises(ValueError):
                api.normalize_scopes(scopes)

    def test_untrusted_pagination_never_receives_auth(self):
        urls = ['https://example.com/v1.0/me', 'http://graph.microsoft.com/v1.0/me',
                'https://graph.microsoft.com.evil.example/v1.0/me',
                'https://user@graph.microsoft.com/v1.0/me', '//example.com/v1.0/me',
                'https://graph.microsoft.com/beta/me', 'https://graph.microsoft.com/v1.0/me#fragment']
        for url in urls:
            with self.assertRaises(ValueError):
                self.client.get(url)
        self.client._requests.get.assert_not_called()

    def test_get_only_and_redirects_disabled(self):
        self.client._requests.get.return_value = Mock(status_code=200, json=Mock(return_value={'value': []}))
        self.assertEqual(self.client.get('/me'), {'value': []})
        self.assertFalse(self.client._requests.get.call_args.kwargs['allow_redirects'])
        self.assertEqual(self.client._requests.get.call_args.args[0], 'https://graph.microsoft.com/v1.0/me')
        for method in ['post', 'patch', 'delete', 'put']:
            getattr(self.client._requests, method).assert_not_called()

    def test_errors_do_not_leak_response_body(self):
        self.client._requests.get.return_value = Mock(status_code=302, text='sensitive-body')
        with self.assertRaisesRegex(RuntimeError, 'HTTP 302') as caught:
            self.client.get('/me')
        self.assertNotIn('sensitive-body', str(caught.exception))

    def test_pagination_params_and_cycle_detection(self):
        self.client.get = Mock(side_effect=[{'value': [1], '@odata.nextLink': '/me?p=2'}, {'value': [2]}])
        self.assertEqual(self.client.get_all_pages('/me', {'$top': 1}), [1, 2])
        self.assertEqual(self.client.get.call_args_list[0].kwargs['params'], {'$top': 1})
        self.assertIsNone(self.client.get.call_args_list[1].kwargs['params'])
        self.client.get = Mock(return_value={'value': [], '@odata.nextLink': '/me'})
        with self.assertRaisesRegex(RuntimeError, 'repeated'):
            self.client.get_all_pages('/me')

    def test_dotenv_preserves_process_values(self):
        with tempfile.TemporaryDirectory() as tmp:
            env = Path(tmp) / '.env'
            env.write_text('# comment\nGRAPH_TEST_A="quoted value"\nGRAPH_TEST_B=new\n', encoding='utf-8')
            with patch.dict(os.environ, {'GRAPH_TEST_B': 'existing'}, clear=True):
                api.load_env(str(env))
                self.assertEqual(os.environ['GRAPH_TEST_A'], 'quoted value')
                self.assertEqual(os.environ['GRAPH_TEST_B'], 'existing')


class GraphScan(unittest.TestCase):
    def message(self, id, date):
        return {'id': id, 'receivedDateTime': date, 'subject': 'Hello\nWorld', 'from': {'emailAddress': {'address': 'NEWS@example.com'}}}

    def test_message_cap_and_cutoff(self):
        client = Mock()
        client.get.return_value = {'value': [self.message(str(i), '2026-01-02T00:00:00Z') for i in range(5)]}
        rows = graph.collect_messages(client, 'inbox', '收件箱', datetime(2026, 1, 1, tzinfo=UTC), 2)
        self.assertEqual(len(rows), 2)
        self.assertEqual(rows[0]['fromAddress'], 'news@example.com')
        self.assertEqual(rows[0]['subject'], 'Hello World')
        client.get.reset_mock()
        client.get.return_value = {'value': [self.message('recent', '2026-01-02T00:00:00Z'), self.message('old', '2025-12-31T00:00:00Z')], '@odata.nextLink': 'unused'}
        rows = graph.collect_messages(client, 'inbox', 'Inbox', datetime(2026, 1, 1, tzinfo=UTC), 20)
        self.assertEqual([r['id'] for r in rows], ['recent'])
        client.get.assert_called_once()

    def test_localized_inbox_and_null_rule(self):
        client = Mock()
        client.get.return_value = {'id': 'localized-inbox'}
        self.assertEqual(graph.get_inbox_id(client), 'localized-inbox')
        client.get.assert_called_once_with('/me/mailFolders/inbox', params={'$select': 'id'})
        self.assertEqual(graph.flatten_rule({'conditions': None, 'actions': None})['senderContains'], '')

    def test_folder_traversal_deduplicates_ids(self):
        client = Mock()
        client.get_all_pages.side_effect = [[{'id': 'root', 'childFolderCount': 1}], [{'id': 'root', 'childFolderCount': 1}, {'id': 'child'}]]
        self.assertEqual([f['id'] for f in graph.traverse_folders(client)], ['root', 'child'])
        self.assertEqual(client.get_all_pages.call_count, 2)

    def test_main_localized_inbox_and_existing_rules(self):
        client = Mock()
        client.get.side_effect = [{'displayName': 'Synthetic', 'mail': 'reader@example.com'}, {'id': 'inbox'}]
        client.get_all_pages.return_value = [{'isEnabled': True, 'conditions': {'senderContains': ['covered@example.com']}}]
        folders = [{'id': 'inbox', 'displayName': '收件箱'}, {'id': 'other', 'displayName': 'Inbox'}]
        messages = [[{'folderId': 'inbox', 'folderName': '收件箱', 'categories': '', 'fromAddress': a} for a in ['covered@example.com', 'news@example.org']], [{'folderId': 'other', 'folderName': 'Inbox', 'categories': '', 'fromAddress': 'wrong@example.com'}]]
        with tempfile.TemporaryDirectory() as tmp, patch.object(graph, 'load_env'), patch.dict(os.environ, {'GRAPH_CLIENT_ID': 'synthetic'}), patch.object(graph, 'GraphClient', return_value=client), patch.object(graph, 'traverse_folders', return_value=folders), patch.object(graph, 'collect_messages', side_effect=messages), patch.object(sys, 'argv', ['scan_readonly.py', '--output-dir', tmp, '--suggest-threshold', '1']), patch('sys.stdout', new_callable=io.StringIO):
            graph.main()
            profile = json.loads((Path(tmp) / 'mailbox_profile.json').read_text(encoding='utf-8'))
            proposed = json.loads((Path(tmp) / 'proposed_rules.json').read_text(encoding='utf-8'))
            self.assertEqual(profile['inboxUnsortedCount'], 2)
            self.assertEqual([p['match'] for p in proposed], ['news@example.org'])


class Items:
    def __init__(self, items, fail_sort=False):
        self.items, self.Count, self.fail_sort = items, len(items), fail_sort
    def Sort(self, *args):
        if self.fail_sort:
            raise RuntimeError('Synthetic sort failure')
    def Item(self, index):
        return self.items[index - 1]


class LocalScan(unittest.TestCase):
    def message(self, id, date):
        return NS(Class=43, ReceivedTime=date, EntryID=id, SenderEmailAddress='News@EXAMPLE.com', Subject='Example', Categories='', Attachments=NS(Count=0))

    def test_aware_datetimes_and_cap(self):
        date = datetime(2026, 1, 2, 11, tzinfo=timezone(timedelta(hours=11)))
        folder = NS(Items=Items([self.message('one', date), self.message('two', date)]), EntryID='inbox', StoreID='store')
        rows = local.collect_messages_from_folder(folder, '邮箱/收件箱', datetime(2026, 1, 1, tzinfo=UTC), 1)
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]['senderEmail'], 'news@example.com')
        self.assertEqual(rows[0]['folderEntryId'], 'inbox')

    def test_unsorted_old_messages_do_not_hide_recent_mail(self):
        folder = NS(Items=Items([self.message('old', datetime(2025, 1, 1, tzinfo=UTC)), self.message('new', datetime(2026, 1, 2, tzinfo=UTC))], fail_sort=True))
        rows = local.collect_messages_from_folder(folder, 'Inbox', datetime(2026, 1, 1, tzinfo=UTC), 5)
        self.assertEqual([r['entryId'] for r in rows], ['new'])

    def test_nonmail_and_naive_datetimes(self):
        folder = NS(Items=Items([NS(Class=26), self.message('one', datetime(2026, 1, 2, 12))]))
        self.assertEqual(len(local.collect_messages_from_folder(folder, 'Inbox', datetime(2025, 1, 1, tzinfo=UTC), 5)), 1)

    def test_main_uses_store_and_folder_ids(self):
        inbox = NS(EntryID='inbox', StoreID='store')
        store = NS(DisplayName='Synthetic', GetRootFolder=Mock(), GetDefaultFolder=Mock(return_value=inbox))
        outlook = NS(Session=NS(Stores=Items([store]), GetFolderFromID=Mock()))
        pythoncom = Mock()
        client = NS(gencache=NS(EnsureDispatch=Mock(return_value=outlook)))
        modules = {'pythoncom': pythoncom, 'win32com': NS(client=client), 'win32com.client': client}
        folders = [{'entryId': i, 'storeId': 'store', 'path': p} for i, p in [('inbox', '邮箱/收件箱'), ('child', '邮箱/Inbox/Child')]]
        messages = [[{'folderEntryId': i, 'storeId': 'store', 'categories': '', 'senderEmail': 'news@example.com'}] for i in ['inbox', 'child']]
        with tempfile.TemporaryDirectory() as tmp, patch.dict(sys.modules, modules), patch.object(local, 'collect_folders', return_value=folders), patch.object(local, 'collect_rules', return_value=[]), patch.object(local, 'collect_messages_from_folder', side_effect=messages), patch.object(sys, 'argv', ['scan_outlook_local_readonly.py', '--output-dir', tmp]), patch('sys.stdout', new_callable=io.StringIO):
            local.main()
            profile = json.loads((Path(tmp) / 'mailbox_profile.json').read_text(encoding='utf-8'))
            self.assertEqual(profile['inboxUnsortedCount'], 1)
            pythoncom.CoUninitialize.assert_called_once()


class CommandLine(unittest.TestCase):
    def test_help_without_any_site_packages(self):
        root = Path(__file__).resolve().parents[1]
        for script in ['scan_readonly.py', 'scan_outlook_local_readonly.py']:
            result = subprocess.run([sys.executable, '-S', str(root / script), '--help'], capture_output=True, text=True, timeout=10)
            self.assertEqual(result.returncode, 0, result.stderr)
            self.assertIn('--max-messages-per-folder', result.stdout)

    def test_bad_limits_fail_before_login_or_dependency_loading(self):
        root = Path(__file__).resolve().parents[1]
        for script in ['scan_readonly.py', 'scan_outlook_local_readonly.py']:
            result = subprocess.run([sys.executable, '-S', str(root / script), '--days', '0'], capture_output=True, text=True, timeout=10)
            self.assertEqual(result.returncode, 2)
            self.assertIn('positive integer', result.stderr)


if __name__ == '__main__':
    unittest.main()
