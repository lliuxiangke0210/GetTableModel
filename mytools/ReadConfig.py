import configparser
import os


class ReadConfig(object):
    """定义一个读取配置文件的类"""

    def __init__(self, filepath=None):
        if filepath:
            configpath = filepath
        else:
            
            root_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
            configpath = os.path.join(root_dir, "config/config.ini")
        self.cf = configparser.ConfigParser()
        self.cf.read(configpath)

    def get_param(self, param):
        value = self.cf.get("mysql-database", param)
        return value


if __name__ == '__main__':
    test = ReadConfig()
    t = test.get_param("host")
    print(t)
