import logging
import logging.config
import os
from base.read_yaml import Read_yaml

class Log(Read_yaml):
    """
    日志
    """
    def get_log(self, logger='webUI'):
        logconf_path = os.path.join(os.path.dirname(__file__), 'logconfig.yaml')
        logconf_dict = self.read_yaml(logconf_path)
        print(logconf_dict)
        logging.config.dictConfig(logconf_dict)
        return logging.getLogger(logger)

