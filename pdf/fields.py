class SearchField:
    def __init__(self, x=0, y=0, maskStr=None, maskState=None, fullString=""):
        if maskStr is None:
            maskStr = []
        if maskState is None:
            maskState = []
        self.x = x
        self.y = y
        self.maskStr = maskStr
        self.maskState = maskState
        self.fullString = fullString

    def getMaskEl(self, ind):
        return self.maskStr[ind]

    def getMaskState(self, ind):
        return self.maskState[ind]

class MinedField:
    def __init__(self, minedElement, x0=0, y0=0, w=0, h=0, text=None, UIElement=None):
        self.element = minedElement
        self.x0, self.y0 = x0, y0
        self.w, self.h = w, h

        self.text = text
        self.UIElement = UIElement

    def __repr__(self):
        return f"{self.text}, {self.x0}, {self.y0}"