import unittest
from unittest.mock import patch, MagicMock
from datetime import datetime
import pandas as pd
from outlook_handler import get_outlook_meetings


class TestOutlookHandler(unittest.TestCase):

    @patch('outlook_handler.win32com.client.Dispatch')
    def test_get_outlook_meetings(self, mock_dispatch):
        mock_outlook = MagicMock()
        mock_dispatch.return_value.GetNamespace.return_value.GetDefaultFolder.return_value.Items.Restrict.return_value = [
            MagicMock(Subject='Meeting 1', Start=datetime(
                2023, 8, 1, 9, 0), End=datetime(2023, 8, 1, 10, 0)),
            MagicMock(Subject='Meeting 2', Start=datetime(
                2023, 8, 1, 10, 0), End=datetime(2023, 8, 1, 11, 0))
        ]
        mock_dispatch.return_value.GetNamespace.return_value.GetDefaultFolder.return_value.Items.IncludeRecurrences = True
        start_date = datetime(2023, 8, 1)
        end_date = datetime(2023, 8, 7, 23, 59, 59)
        df = get_outlook_meetings(start_date, end_date)
        self.assertEqual(len(df), 2)
        self.assertEqual(df.iloc[0]['Subject'], 'Meeting 1')


if __name__ == '__main__':
    unittest.main()
