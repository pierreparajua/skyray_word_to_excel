

def modified_indentation(text):
    xs = str(text).split("\n")
    changes =[]
    indexes = []

    changes2 =[]
    indexes2 = []

    for x in xs:
        if x.startswith("\n"):
            indexes.append(xs.index(x))
            changes.append(x.replace("o\t", "\n\t\to "))
        if x.startswith("•\t"):
            ind2 = xs.index(x)
            indexes2.append(xs.index(x))
            changes2.append(x.replace("•\t", "\n\t• "))
              
    for i,c in zip(indexes, changes):
        xs[i] = c

    for i,c in zip(indexes2, changes2):
        xs[i] = c
            
    text = ''.join(xs)       
    return text
