import re
text = "A.  1945-1991. B.1991-2000. C.1945-1973. D.1973-1991."
lst = re.split(r'\s+(?=[a-dA-D]\.)', text.__str__).
for i in lst:
    print(i,"\n")