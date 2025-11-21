import unittest
from unittest.mock import patch, MagicMock
from pathlib import Path
import reprocessing

class TestReprocessingManager(unittest.TestCase):
    def setUp(self):
        self.manager = reprocessing.ReprocessingManager(attachments_root='test_attachments')

    @patch('reprocessing.get_failed_directories_for_reprocessing')
    def test_get_directories_for_reprocessing(self, mock_get_failed):
        mock_get_failed.return_value = [
            {'directory_name': 'dir1', 'is_reprocessing_enabled': True},
            {'directory_name': 'dir2', 'is_reprocessing_enabled': False}
        ]
        dirs = self.manager.get_directories_for_reprocessing()
        self.assertEqual(len(dirs), 2)
        self.assertEqual(dirs[0]['directory_name'], 'dir1')

    @patch('reprocessing.reset_processing_flags')
    @patch('reprocessing.update_reprocessing_enabled')
    @patch('reprocessing.ReprocessingManager._clean_metadata_flags')
    @patch('reprocessing.ReprocessingManager._remove_cleaned_files')
    def test_reset_directory_for_reprocessing_success(self, mock_remove, mock_clean, mock_update, mock_reset):
        mock_clean.return_value = True
        mock_reset.return_value = True
        mock_update.return_value = None
        # Create a dummy directory
        test_dir = Path('test_attachments/dir1')
        test_dir.mkdir(parents=True, exist_ok=True)
        result = self.manager.reset_directory_for_reprocessing('dir1')
        self.assertTrue(result)
        # Cleanup
        test_dir.rmdir()
        Path('test_attachments').rmdir()

    @patch('reprocessing.ReprocessingManager._clean_metadata_flags')
    def test_reset_directory_for_reprocessing_missing_dir(self, mock_clean):
        mock_clean.return_value = True
        result = self.manager.reset_directory_for_reprocessing('nonexistent_dir')
        self.assertFalse(result)

    @patch('reprocessing.ReprocessingManager._clean_metadata_flags')
    @patch('reprocessing.ReprocessingManager._remove_cleaned_files')
    @patch('reprocessing.reset_processing_flags')
    @patch('reprocessing.update_reprocessing_enabled')
    def test_reset_directory_for_reprocessing_clean_fail(self, mock_update, mock_reset, mock_remove, mock_clean):
        mock_clean.return_value = False
        result = self.manager.reset_directory_for_reprocessing('dir1')
        self.assertFalse(result)

    @patch('reprocessing.ReprocessingManager._clean_metadata_flags')
    @patch('reprocessing.ReprocessingManager._remove_cleaned_files')
    @patch('reprocessing.reset_processing_flags')
    @patch('reprocessing.update_reprocessing_enabled')
    def test_reset_directory_for_reprocessing_reset_fail(self, mock_update, mock_reset, mock_remove, mock_clean):
        mock_clean.return_value = True
        mock_reset.return_value = False
        result = self.manager.reset_directory_for_reprocessing('dir1')
        self.assertFalse(result)

    @patch('reprocessing.ReprocessingManager._clean_metadata_flags')
    @patch('reprocessing.ReprocessingManager._remove_cleaned_files')
    @patch('reprocessing.reset_processing_flags')
    @patch('reprocessing.update_reprocessing_enabled')
    def test_reset_directory_for_reprocessing_exception(self, mock_update, mock_reset, mock_remove, mock_clean):
        mock_clean.side_effect = Exception('fail')
        result = self.manager.reset_directory_for_reprocessing('dir1')
        self.assertFalse(result)

    @patch('reprocessing.get_failed_directories_for_reprocessing')
    @patch('reprocessing.ReprocessingManager.reset_directory_for_reprocessing')
    def test_reprocess_all_enabled_directories(self, mock_reset, mock_get_failed):
        mock_get_failed.return_value = [
            {'directory_name': 'dir1', 'is_reprocessing_enabled': True},
            {'directory_name': 'dir2', 'is_reprocessing_enabled': False}
        ]
        mock_reset.return_value = True
        results = self.manager.reprocess_all_enabled_directories()
        self.assertIn('dir1', results)
        self.assertTrue(results['dir1'])
        self.assertNotIn('dir2', results)

    @patch('reprocessing.Path.exists')
    @patch('reprocessing.json.load')
    @patch('reprocessing.Path.open')
    @patch('reprocessing.json.dump')
    def test_clean_metadata_flags_preserves_jera_fetched(self, mock_dump, mock_open, mock_load, mock_exists):
        mock_exists.return_value = True
        mock_load.return_value = {'jera_fetched': True, 'date_verification_ingestion': True}
        result = self.manager._clean_metadata_flags(Path('test_attachments/dir1'))
        self.assertTrue(result)

    @patch('reprocessing.Path.exists')
    def test_clean_metadata_flags_missing_metadata(self, mock_exists):
        mock_exists.return_value = False
        result = self.manager._clean_metadata_flags(Path('test_attachments/dir1'))
        self.assertFalse(result)

    @patch('reprocessing.Path.glob')
    def test_remove_cleaned_files(self, mock_glob):
        mock_file = MagicMock()
        mock_file.is_file.return_value = True
        mock_file.name = 'file_cleaned.xlsx'
        mock_glob.return_value = [mock_file]
        self.manager._remove_cleaned_files(Path('test_attachments/dir1'))
        mock_file.unlink.assert_called()

if __name__ == '__main__':
    unittest.main()
