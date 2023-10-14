import requests
import pandas as pd

class Konjukturprognos:
    def __init__(self, uppladdfil, nerladdfil) -> None:
        dlurl = self.dlurl #str
        self.uppladdfil = uppladdfil
        self.nerladdfil = nerladdfil
        

    def laddaner(dlurl):
        
        nerladd = requests.get(dlurl)
        fil = nerladd.content #excel-fil
        df = pd.read_excel(fil) #data

        return df

    def uppladdfil
