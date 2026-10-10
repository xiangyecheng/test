import pytest
import allure

if __name__ == '__main__':
    pytest.main([
        'testcase/test_login.py',
        '-vs',
        '--alluredir=results ',
        '--clean-alluredir'
    ])
