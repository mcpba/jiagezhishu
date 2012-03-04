# -*- coding: utf8 -*-

class Goods:
        
    def __init__(self,ttl,amout,p0):
        self.ttl=ttl           #总数量
        self.amout=amout       #完税价格总价
        self.huilv=1.0         #汇率，考虑多月份计算用
        self.p0=p0
        self.p1=0.0
        self.affect=999.0        #对指标的影响度
    
    def setP1(self):
        if self.ttl != 0:
            self.p1=self.amout/(self.ttl*self.huilv)
        else:
            self.p1=0.0
