

def modified_indentation(text: str) -> str:
    """
    Découpe les phrases provenant du fichier excel source à chaque retour à la ligne '\n'.
    Si la phrase commence par '•' rajoute une tabulation.
    Si la phrase commence par 'o' rajoute deux tabulations.
    """
    sentences = str(text).split("\n")
    changes = []
    indexes = []

    changes2 = []
    indexes2 = []

    for sentence in sentences:
        if sentence.startswith("o\t"):
            indexes.append(sentences.index(sentence))
            changes.append(sentence.replace("o\t", "\n\t\to "))
        if sentence.startswith("•\t"):
            indexes2.append(sentences.index(sentence))
            changes2.append(sentence.replace("•\t", "\n\t• "))

    for i, c in zip(indexes, changes):
        sentences[i] = c

    for i, c in zip(indexes2, changes2):
        sentences[i] = c

    text = ''.join(sentences)
    return text


