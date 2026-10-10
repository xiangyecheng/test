import yaml

class Read_yaml:
    def read_yaml(self, filepath):
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                result = yaml.load(f.read(), Loader=yaml.FullLoader)
                return result
        except Exception as e:
            print(e)
