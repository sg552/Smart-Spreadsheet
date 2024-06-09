class Table:
    def __init__(self, data):
        self.data = data
        self.top = 0
        self.left = 0
        self.right = 0
        self.bottom = 0


    def to_text(self):
        lines = []
        for i,row in enumerate(self.data):
            lines.append('|'.join(str(e) for e in self.data[i]))
        return "\r\n".join(lines)

