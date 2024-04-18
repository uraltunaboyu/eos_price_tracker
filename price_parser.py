from io import TextIOWrapper
import re

NAME_ENTRY_REGEX = r'\[\"(.+?)\"\]=\{\[\d*\]=(\d*),\}'
PRICE_ENTRY_REGEX = r'(?:\[(\d+?)\]=.*?)(\["A.*?\d+)\,\}'
PRICE_VALS_REGEX = r'(?:\["(\D+)"\]=(\d+\.?\d*),?)'

COMPARISON_FIELD = 'SaleAmountCount'

class PriceParser:
    def __init__(self):
        self.names = {}
        self.prices = {}

    def load_names(self, path: str):
        with open(path) as names_file:
            self.parse_names(names_file)
            return

    def parse_names(self, names_file: TextIOWrapper):
        name_regex = re.compile(NAME_ENTRY_REGEX)
        all_text = names_file.read()
        try:
            all_names = name_regex.findall(all_text)
            for name in all_names:
                self.names[name[0]] = name[1]
        except Exception as e:
            print("Unable to parse names file!", e)
            return
    
    def load_prices(self, path: str):
        with open(path) as prices_file:
            self.parse_prices(prices_file)

    def parse_prices(self, prices_file: TextIOWrapper):
        prices_regex = re.compile(PRICE_ENTRY_REGEX)
        entry_regex = re.compile(PRICE_VALS_REGEX)
        all_text = prices_file.read()
        try:
            all_entries = prices_regex.findall(all_text)
            for entry in all_entries:
                all_vals = entry_regex.findall(entry[1])
                vals_dict = {}
                for val in all_vals:
                    vals_dict[val[0]] = val[1]
                if entry[0] in self.prices:
                    self.prices[entry[0]] = self.get_better_entry(self.prices[entry[0]], vals_dict)
                else:
                    self.prices[entry[0]] = vals_dict
        except Exception as e:
            print("Unable to parse prices file!", e)
            return
        
    def get_better_entry(self, sample: dict, candidate: dict):
        if COMPARISON_FIELD not in sample and COMPARISON_FIELD not in candidate:
            return sample if len(sample.keys()) > len(candidate.keys()) else candidate
        if COMPARISON_FIELD not in sample:
            return candidate
        if COMPARISON_FIELD not in candidate:
            return sample
        return sample if sample[COMPARISON_FIELD] > candidate[COMPARISON_FIELD] else candidate
    
    def get_attr(self, name: str, attr: str):
        if name not in self.names:
            print(f"Unable to find {name}")
            return None
        if attr not in self.prices[self.names[name]]:
            print(f"Unable to find {attr} in {name}")
            return None
        return self.prices[self.names[name]][attr]