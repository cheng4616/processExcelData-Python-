class Account(object):
    """docstring for ClassName"""
    def __init__(self, date, bussinessContext, item, debit, credit):
        self.date = date
        self.bussinessContext = bussinessContext
        self.item = item
        self.debit = debit
        self.credit = credit
    def getAccount(self):
        print(self.date, self.bussinessContext, self.item, self.debit, self.credit)
