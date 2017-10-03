import attr

@attr.s
class InterimResult:
    p = attr.ib()
    s = attr.ib()
    v = attr.ib()

    def getP(self):
        return self.p

    def getS(self):
        return self.s

    def getV(self):
        return self.v

    def setP(self, p: float):
        self.p = p

    def setS(self, s: int):
        self.s = s

    def setV(self, v: int):
        self.v = v