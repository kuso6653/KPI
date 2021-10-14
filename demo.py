import pandas as pd

a0 = {"number": range(10), "letter": ["a", "a", "b", "b", "c", "f", "f", "e", "h", "w"]}
a = pd.DataFrame(a0)
b0 = {"number": range(15), "letter": ["b", "a", "t", "b", "r", "f", "g", "e", "j", "w", "t", "h", "i", "y", "u"]}
b = pd.DataFrame(b0)
print(a)
print(b)
c = a.append(b)
c.drop_duplicates(keep=False, inplace=True)
c.reset_index()
print(c)
