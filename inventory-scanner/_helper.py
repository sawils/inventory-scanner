'''
Helper class

MUST ADD ERROR HANDLING.
'''

import re
from datetime import datetime

class Customer:
    def __init__(self, UUID=0, Name=''):
        self.uuid = UUID
        self.name = Name

    def get_name(self):
        return self.name

    def _pretty(self,data):
        for c in CUSTOMERS:
            if str(c.uuid) == str(customer_name):
                customer_name = c.name
        
class AppID:
    def __init__(self, AI='', description='', format='', length=0, unit=''):
        self.AI = AI
        self.desc = description
        self._format = format
        self._length = length
        self._unit = unit
        
    def get_RE(self):
        return re.compile('^%s([0-9]{%d})' %(self.AI, self._length))
    
    def get_description(self, data):
        return {self.desc: self._pretty(data)}
    
    def _pretty(self, data):
        data = data[len(self.AI):]
        if self._format == "weight":
            return self._get_readable_weight(data)
        elif self._format == "YYMMDD":
            return self._get_readable_yymmdd(data)
        else:
            return data
        
    def _get_readable_weight(self, string):
        '''Readable format for type weight'''
        try:
            decimal_places = int(self.AI[-1])
        except ValueError as e:
            raise e("Handle non-numeric last AI characters for weight!")
        weight = float(string) / (10**decimal_places)
        if self._unit == "kg":
            weight *= 2.2046
        return float('%.2f' % weight)

    
    def _get_readable_yymmdd(self, string):
        '''Readable format for type yymmdd'''
        if len(string) == 6:
            string = '20' + string    #convert to YYYYMMDD
            time = datetime.strptime(string, '%Y%m%d')
            return time.strftime("%b %d, %Y")
        
