import pytest

from base.read_yaml import Read_yaml
from base.read_excel import EXCEL

class TestPerson:
    yaml = Read_yaml().read_yaml('./data/login.yaml')
    excel1 = EXCEL().rd_excel('./data/person.xlsx', 'Sheet1')
    # excel2 = EXCEL().rd_excel('./data/person.xlsx', 'Sheet2')
    @pytest.mark.parametrize('person', excel1)
    def test_read_yaml(self, person):
        assert person['验证'] == person['姓名']
        # assert '向业成2' == process['name']
        # assert '向业成3' == process['name']
        # assert '湖南1' == process['address']
        # assert '湖南2' == process['address']
        # assert '湖南3' == process['address']
        # pass


