import pandas as pd
from app import group_activities

def test_group_activities_handles_empty_task_ref():
    df = pd.DataFrame({
        'Task Reference Number': ['', None, 'A'],
        'Date': ['2020-01-01', '2020-01-02', '2020-01-03'],
        'Modification Details': ['Fix', 'More', 'Other'],
    })
    grouped = group_activities(df)
    assert 'No Reference' in grouped
    assert 'A' in grouped
